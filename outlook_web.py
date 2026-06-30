"""Localhost web dashboard for the Outlook MCP server.

A lightweight FastAPI app (served by uvicorn — both already ship as transitive
dependencies of ``mcp``, so no new packages are required) that puts a friendly,
triage-first UI in front of the very same tool functions the MCP server exposes.

Architecture / why it is safe and fast on busy mailboxes:

* **One persistent COM connection.** Outlook COM is single-threaded-apartment
  (STA). All Outlook access is funnelled through a single dedicated worker
  thread that initialises COM once and keeps the connection warm — so requests
  never pay the re-``Dispatch`` cost and never race across threads.
* **Zero logic duplication.** The worker re-points the MCP server's connection
  helpers at the persistent connection and then calls the *existing*, tested
  tool functions (search, triage, draft, send-with-confirm, archive, …). The
  web layer only does HTTP plumbing.
* **Localhost only.** Binds to 127.0.0.1 by default. An optional shared token
  (``OUTLOOK_WEB_TOKEN``) adds a second factor if you want it.

Run it::

    python outlook_web.py            # http://127.0.0.1:8765

Test it without Outlook by calling :func:`create_app` with your own async
``call`` shim (see tests/test_web.py).
"""

from __future__ import annotations

import asyncio
import json
import os
import queue
import threading
from concurrent.futures import Future
from typing import Any, Awaitable, Callable, Dict, Optional

from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, JSONResponse

import outlook_helpers as H
from outlook_helpers import ErrorCode, make_error

log = H.get_logger()

WEB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "web")
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = int(os.environ.get("OUTLOOK_WEB_PORT", "8765"))

# Type of the injectable bridge: (tool_function, **kwargs) -> envelope dict.
CallFn = Callable[..., Awaitable[Dict[str, Any]]]


# ---------------------------------------------------------------------------
# COM worker thread (persistent, serialised Outlook access)
# ---------------------------------------------------------------------------

def _run_tool(tool_fn: Callable[..., str], kwargs: Dict[str, Any]) -> Dict[str, Any]:
    """Execute an MCP tool function and parse its JSON envelope to a dict."""
    return json.loads(tool_fn(**kwargs))


class OutlookWorker:
    """Owns the single Outlook COM connection and serialises all access."""

    def __init__(self) -> None:
        self._queue: "queue.Queue" = queue.Queue()
        self._ready = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self.error: Optional[Exception] = None
        self.connected = False

    def start(self) -> None:
        self._thread = threading.Thread(
            target=self._run, name="outlook-com-worker", daemon=True)
        self._thread.start()

    def _run(self) -> None:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore

            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            # Re-point the MCP server at this warm, thread-owned connection so we
            # reuse all of its tool logic unchanged.
            import outlook_mcp_server as omcp
            omcp._connect = lambda: (outlook, namespace)   # type: ignore
            omcp._namespace = lambda: namespace            # type: ignore
            omcp._WIN32_AVAILABLE = True                   # type: ignore
            self.connected = True
            log.info("Web worker connected to Outlook (warm connection).")
        except Exception as exc:  # pragma: no cover - needs Windows/Outlook
            self.error = exc
            log.error("Web worker could not connect to Outlook: %s", exc)
            self._ready.set()
            return

        self._ready.set()
        while True:
            item = self._queue.get()
            if item is None:
                break
            fn, args, kwargs, fut = item
            try:
                fut.set_result(fn(*args, **kwargs))
            except Exception as exc:  # noqa: BLE001 - surfaced to caller
                fut.set_exception(exc)
        try:  # pragma: no cover
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass

    def _submit(self, fn: Callable[..., Any], *args: Any, **kwargs: Any) -> Future:
        fut: Future = Future()
        self._queue.put((fn, args, kwargs, fut))
        return fut

    async def acall(self, tool_fn: Callable[..., str], **kwargs: Any) -> Dict[str, Any]:
        if not self._ready.is_set():
            await asyncio.to_thread(self._ready.wait, 30)
        if self.error is not None:
            return make_error(
                ErrorCode.OUTLOOK_CONNECTION_FAILED,
                "The Outlook connection is not available.",
                details="Open the classic Outlook desktop app and sign in, then retry.",
                retryable=True,
            )
        return await asyncio.wrap_future(self._submit(_run_tool, tool_fn, kwargs))


# ---------------------------------------------------------------------------
# App factory
# ---------------------------------------------------------------------------

def create_app(call: CallFn) -> FastAPI:
    """Build the FastAPI app. ``call`` bridges to the MCP tool functions."""
    import outlook_mcp_server as omcp

    app = FastAPI(title="Outlook Assistant", docs_url=None, redoc_url=None)

    # ---- optional shared-token auth --------------------------------------
    @app.middleware("http")
    async def _auth(request: Request, call_next):
        token = os.environ.get("OUTLOOK_WEB_TOKEN", "").strip()
        open_paths = {"/", "/favicon.ico", "/api/health"}
        if token and request.url.path not in open_paths:
            provided = (request.headers.get("X-Outlook-Token")
                        or request.query_params.get("token"))
            if provided != token:
                return JSONResponse(
                    make_error(ErrorCode.PERMISSION_DENIED,
                               "Invalid or missing token."),
                    status_code=401)
        return await call_next(request)

    def _bad(field: str) -> JSONResponse:
        return JSONResponse(
            make_error(ErrorCode.INVALID_PARAMETER, f"Missing required field: {field}."),
            status_code=400)

    # ---- static page ------------------------------------------------------
    @app.get("/")
    async def index():
        return FileResponse(os.path.join(WEB_DIR, "index.html"))

    @app.get("/favicon.ico")
    async def favicon():
        return JSONResponse({}, status_code=204)

    # ---- read-only --------------------------------------------------------
    @app.get("/api/health")
    async def health():
        return {"success": True, "service": "outlook-assistant-web",
                "auth_required": bool(os.environ.get("OUTLOOK_WEB_TOKEN", "").strip())}

    @app.get("/api/unread_count")
    async def unread_count(folder: Optional[str] = None):
        return await call(omcp.count_unread_emails, folder_name=folder)

    @app.get("/api/folders")
    async def folders():
        return await call(omcp.list_folders)

    @app.get("/api/categories")
    async def categories():
        return await call(omcp.list_categories)

    @app.get("/api/drafts")
    async def drafts():
        return await call(omcp.list_drafts)

    @app.get("/api/triage")
    async def triage(days: int = 3, limit: int = 20, unread_only: bool = False,
                     folder: Optional[str] = None):
        return await call(omcp.triage_inbox, days=days, max_results=limit,
                          unread_only=unread_only, folder_name=folder)

    @app.get("/api/search")
    async def search(keyword: Optional[str] = None, sender: Optional[str] = None,
                     subject: Optional[str] = None, days: int = 14,
                     unread_only: bool = False, has_attachments: Optional[bool] = None,
                     folder: Optional[str] = None, max_results: int = 25,
                     offset: int = 0):
        return await call(omcp.search_emails, keyword=keyword, sender=sender,
                          subject=subject, days=days, unread_only=unread_only,
                          has_attachments=has_attachments, folder_name=folder,
                          max_results=max_results, offset=offset)

    @app.get("/api/email")
    async def email(entry_id: str):
        return await call(omcp.get_email_by_number, entry_id=entry_id)

    @app.get("/api/thread")
    async def thread(entry_id: str, days: int = 60):
        return await call(omcp.read_thread, entry_id=entry_id, days=days)

    @app.get("/api/attachments")
    async def attachments(entry_id: str):
        return await call(omcp.list_attachments, entry_id=entry_id)

    @app.get("/api/digest")
    async def digest():
        return _read_latest_digest()

    # ---- drafts & sending (sending always needs confirm) -----------------
    @app.post("/api/draft")
    async def create_draft(request: Request):
        d = await request.json()
        if not d.get("to"):
            return _bad("to")
        return await call(omcp.create_draft, to=d.get("to"),
                          subject=d.get("subject", ""), body=d.get("body", ""),
                          cc=d.get("cc"), bcc=d.get("bcc"))

    @app.post("/api/update_draft")
    async def update_draft(request: Request):
        d = await request.json()
        if not d.get("draft_id"):
            return _bad("draft_id")
        return await call(omcp.update_draft, draft_id=d["draft_id"],
                          to=d.get("to"), subject=d.get("subject"),
                          body=d.get("body"), cc=d.get("cc"), bcc=d.get("bcc"))

    @app.post("/api/reply")
    async def reply(request: Request):
        d = await request.json()
        if not d.get("entry_id"):
            return _bad("entry_id")
        return await call(omcp.reply_to_email_by_number, entry_id=d["entry_id"],
                          reply_text=d.get("reply_text", ""),
                          reply_all=bool(d.get("reply_all", False)),
                          send=False)

    @app.post("/api/forward")
    async def forward(request: Request):
        d = await request.json()
        if not d.get("entry_id") or not d.get("to"):
            return _bad("entry_id/to")
        return await call(omcp.forward_email, entry_id=d["entry_id"],
                          to=d["to"], comment=d.get("comment", ""), send=False)

    @app.post("/api/send_draft")
    async def send_draft(request: Request):
        d = await request.json()
        if not d.get("draft_id"):
            return _bad("draft_id")
        return await call(omcp.send_draft, draft_id=d["draft_id"],
                          confirm=bool(d.get("confirm", False)))

    @app.post("/api/send_email")
    async def send_email(request: Request):
        d = await request.json()
        if not d.get("to"):
            return _bad("to")
        return await call(omcp.send_email, to=d["to"], subject=d.get("subject", ""),
                          body=d.get("body", ""), cc=d.get("cc"),
                          confirm=bool(d.get("confirm", False)))

    # ---- organise ---------------------------------------------------------
    @app.post("/api/archive")
    async def archive(request: Request):
        d = await request.json()
        if not d.get("entry_id"):
            return _bad("entry_id")
        return await call(omcp.archive_email, entry_id=d["entry_id"])

    @app.post("/api/trash")
    async def trash(request: Request):
        d = await request.json()
        if not d.get("entry_id"):
            return _bad("entry_id")
        return await call(omcp.move_to_trash, entry_id=d["entry_id"])

    @app.post("/api/mark")
    async def mark(request: Request):
        d = await request.json()
        if not d.get("entry_id"):
            return _bad("entry_id")
        tool = omcp.mark_as_read if d.get("read", True) else omcp.mark_as_unread
        return await call(tool, entry_id=d["entry_id"])

    @app.post("/api/category")
    async def category(request: Request):
        d = await request.json()
        if not d.get("entry_id") or not d.get("category"):
            return _bad("entry_id/category")
        tool = omcp.remove_category if d.get("op") == "remove" else omcp.apply_category
        return await call(tool, entry_id=d["entry_id"], category=d["category"])

    @app.post("/api/move")
    async def move(request: Request):
        d = await request.json()
        if not d.get("entry_id") or not d.get("folder"):
            return _bad("entry_id/folder")
        return await call(omcp.move_email_by_number, entry_id=d["entry_id"],
                          destination_folder_name=d["folder"])

    return app


def _read_latest_digest() -> Dict[str, Any]:
    """Return the most recent scheduled digest JSON, if one exists."""
    path = os.path.join(_state_dir(), "digest.json")
    if not os.path.exists(path):
        return {"success": True, "digest": None,
                "message": "No scheduled digest has been generated yet."}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return {"success": True, "digest": json.load(fh)}
    except Exception as exc:
        return make_error(ErrorCode.ACTION_FAILED, "Could not read digest.",
                          details=str(exc))


def _state_dir() -> str:
    base = (os.environ.get("OUTLOOK_STATE_DIR")
            or os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                            "outlook-mcp"))
    return base


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

def main() -> None:
    import argparse
    import uvicorn

    parser = argparse.ArgumentParser(description="Outlook Assistant web dashboard")
    parser.add_argument("--host", default=os.environ.get("OUTLOOK_WEB_HOST", DEFAULT_HOST))
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    args = parser.parse_args()

    if args.host not in ("127.0.0.1", "localhost", "::1"):
        log.warning("Binding to %s exposes the dashboard beyond localhost — "
                    "set OUTLOOK_WEB_TOKEN to require a token.", args.host)

    worker = OutlookWorker()
    worker.start()
    app = create_app(worker.acall)

    log.info("Outlook Assistant web UI on http://%s:%d", args.host, args.port)
    uvicorn.run(app, host=args.host, port=args.port, log_level="warning")


if __name__ == "__main__":
    main()
