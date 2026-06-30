"""Outlook MCP Server.

A Model Context Protocol server that lets an AI agent operate a local Microsoft
Outlook desktop profile (Windows + Outlook, via COM/``pywin32``) as a careful
human email assistant: search, read, summarise, draft, reply, forward, organise.

Safety model (see README for the full contract):

* Reading and searching are always allowed.
* Composing / replying / forwarding default to **creating a draft**, never
  sending.  The only tools that put mail on the wire are ``send_email`` and
  ``send_draft``, and both require an explicit ``confirm=True``.
* Deletion is intentionally limited to **moving to Deleted Items**
  (``move_to_trash``); there is no permanent-delete tool.
* Bulk actions are capped and previewed.
* Logs go to **stderr** only (stdout is reserved for the MCP protocol) and never
  contain message bodies unless ``OUTLOOK_MCP_DEBUG`` is set.

Authentication note: this server does not handle OAuth, tokens or passwords.
It drives the Outlook application that the signed-in Windows user has already
authenticated, so there are no secrets for this process to store or leak.
"""

from __future__ import annotations

import datetime
import functools
import os
import re
import sys
from collections import Counter
from typing import Any, Callable, Dict, List, Optional

from mcp.server.fastmcp import FastMCP

import outlook_helpers as H
from outlook_helpers import ErrorCode, make_error, make_success, to_json

# Optional COM dependency – guarded so the module imports on any platform
# (e.g. for running the test-suite on Linux/CI).
try:  # pragma: no cover - exercised only on Windows with Outlook installed
    import win32com.client  # type: ignore
    import pywintypes  # type: ignore
    _WIN32_AVAILABLE = True
    _COM_ERROR = pywintypes.com_error
except Exception:  # pragma: no cover
    win32com = None  # type: ignore
    _WIN32_AVAILABLE = False
    _COM_ERROR = Exception

log = H.get_logger()
mcp = FastMCP("outlook-assistant")

# ---------------------------------------------------------------------------
# In-process listing cache.
# Listing/search tools assign each result a small ``email_number`` and stash the
# full formatted record here so follow-up tools can reference it ergonomically.
# Every record also carries a stable Outlook ``id`` (EntryID); action tools
# accept either ``email_number`` (from the most recent listing) or ``entry_id``
# (stable across listings) so an agent is never forced to rely on cache state.
# ---------------------------------------------------------------------------
_email_cache: Dict[int, Dict[str, Any]] = {}

MAX_SCAN = 1000  # hard ceiling on items inspected per call (bounds latency)


class OutlookError(Exception):
    """Raised inside tools to produce a structured error envelope."""

    def __init__(self, code: str, message: str, details: Optional[str] = None,
                 retryable: bool = False, **extra: Any):
        super().__init__(message)
        self.code = code
        self.message = message
        self.details = details
        self.retryable = retryable
        self.extra = extra

    def to_payload(self) -> Dict[str, Any]:
        return make_error(self.code, self.message, self.details,
                          self.retryable, **self.extra)


def email_tool(fn: Callable[..., Any]) -> Callable[..., str]:
    """Wrap a tool so it always returns a serialized, structured envelope.

    Catches :class:`OutlookError` (→ its structured payload) and any other
    exception (→ a sanitized ``ACTION_FAILED``), guaranteeing no raw stack
    traces or stdout writes ever reach the agent or corrupt the protocol.
    """

    @functools.wraps(fn)
    def wrapper(*args: Any, **kwargs: Any) -> str:
        try:
            result = fn(*args, **kwargs)
            return result if isinstance(result, str) else to_json(result)
        except OutlookError as exc:
            log.warning("%s failed: %s (%s)", fn.__name__, exc.message, exc.code)
            return to_json(exc.to_payload())
        except _COM_ERROR as exc:  # pragma: no cover - COM-specific
            log.error("%s COM error", fn.__name__, exc_info=H.debug_enabled())
            return to_json(make_error(
                ErrorCode.ACTION_FAILED,
                "Outlook reported an error while performing this action.",
                details=_com_error_hint(exc),
                retryable=True,
            ))
        except Exception as exc:  # noqa: BLE001 - last-resort safety net
            log.error("%s unexpected error", fn.__name__, exc_info=H.debug_enabled())
            return to_json(make_error(
                ErrorCode.ACTION_FAILED,
                "An unexpected error occurred while performing this action.",
                details=str(exc) if H.debug_enabled() else None,
            ))

    return wrapper


def _com_error_hint(exc: Exception) -> str:
    text = str(exc)
    if "0x80040111" in text or "rejected by the user" in text.lower():
        return ("Outlook blocked programmatic access (security prompt). "
                "Ensure Outlook is running and trusted.")
    return "Outlook is busy or returned an error. Try again shortly."


# ---------------------------------------------------------------------------
# Connection / lookup helpers (COM-specific; raise OutlookError on failure)
# ---------------------------------------------------------------------------

def _connect():
    """Return ``(outlook, namespace)`` or raise a structured OutlookError."""
    if not _WIN32_AVAILABLE:
        raise OutlookError(
            ErrorCode.OUTLOOK_NOT_AVAILABLE,
            "Outlook COM bindings are not available on this machine.",
            details="This server requires Windows with Outlook and pywin32 installed.",
        )
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace
    except Exception as exc:  # pragma: no cover
        raise OutlookError(
            ErrorCode.OUTLOOK_CONNECTION_FAILED,
            "Could not connect to Outlook.",
            details="Make sure Outlook is installed, running, and signed in.",
            retryable=True,
        ) from exc


def _namespace():
    return _connect()[1]


def _get_folder_by_name(namespace, folder_name: str):
    """Find a folder by (case-insensitive) name across inbox subfolders and
    store roots. Returns the folder or ``None``."""
    target = folder_name.strip().lower()
    try:
        inbox = namespace.GetDefaultFolder(H.OL_FOLDER_INBOX)
        for folder in inbox.Folders:
            if folder.Name.lower() == target:
                return folder
        for folder in namespace.Folders:
            if folder.Name.lower() == target:
                return folder
            try:
                for sub in folder.Folders:
                    if sub.Name.lower() == target:
                        return sub
            except Exception:
                continue
    except Exception as exc:  # pragma: no cover
        raise OutlookError(ErrorCode.ACTION_FAILED,
                           f"Failed while searching for folder '{folder_name}'.",
                           details=str(exc) if H.debug_enabled() else None)
    return None


def _require_folder(namespace, folder_name: Optional[str]):
    """Return the named folder, or the Inbox when ``folder_name`` is None."""
    if not folder_name:
        return namespace.GetDefaultFolder(H.OL_FOLDER_INBOX), "Inbox"
    folder = _get_folder_by_name(namespace, folder_name)
    if not folder:
        raise OutlookError(
            ErrorCode.FOLDER_NOT_FOUND,
            f"Folder '{folder_name}' was not found.",
            details="Call list_folders to see valid folder names.",
        )
    return folder, folder_name


def _resolve_mail(namespace, email_number: Optional[int] = None,
                  entry_id: Optional[str] = None):
    """Resolve a live Outlook MailItem from a cache number or a stable EntryID."""
    if entry_id:
        try:
            item = namespace.GetItemFromID(entry_id)
        except Exception:
            item = None
        if not item:
            raise OutlookError(ErrorCode.EMAIL_NOT_FOUND,
                               "No email matches the provided entry_id.",
                               details="The message may have been moved or deleted.")
        return item
    if email_number is not None:
        if not _email_cache:
            raise OutlookError(
                ErrorCode.NO_LISTING_CONTEXT,
                "No emails have been listed yet.",
                details="Call search_emails, list_recent_emails or get_unread_emails first, "
                        "or pass a stable entry_id.",
            )
        record = _email_cache.get(int(email_number))
        if not record:
            raise OutlookError(
                ErrorCode.EMAIL_NOT_FOUND,
                f"Email #{email_number} is not in the current listing.",
                details="Re-run a listing/search tool to refresh the numbering, "
                        "or pass entry_id.",
            )
        try:
            item = namespace.GetItemFromID(record["id"])
        except Exception:
            item = None
        if not item:
            raise OutlookError(ErrorCode.EMAIL_NOT_FOUND,
                               f"Email #{email_number} could not be retrieved from Outlook.",
                               details="It may have been moved or deleted since listing.")
        return item
    raise OutlookError(ErrorCode.INVALID_PARAMETER,
                       "Provide either email_number or entry_id.")


def _cache_listing(emails: List[Dict[str, Any]]) -> None:
    """Replace the listing cache with these records, numbered from 1."""
    _email_cache.clear()
    for i, email in enumerate(emails, 1):
        email["email_number"] = i
        _email_cache[i] = email


def _summarize(email: Dict[str, Any]) -> Dict[str, Any]:
    """Token-lean summary view of a formatted email for listings."""
    return {
        "email_number": email.get("email_number"),
        "entry_id": email.get("id"),
        "thread_id": email.get("conversation_id"),
        "subject": email.get("subject"),
        "from": _from_field(email),
        "to": email.get("recipients", []),
        "date": email.get("received_time") or email.get("sent_time"),
        "snippet": email.get("snippet"),
        "labels": H.parse_categories(email.get("categories")),
        "unread": email.get("unread"),
        "has_attachments": email.get("has_attachments"),
        "attachment_count": email.get("attachment_count"),
        "importance": email.get("importance"),
        "flagged": email.get("flagged"),
    }


def _from_field(email: Dict[str, Any]) -> str:
    name = email.get("sender") or ""
    addr = email.get("sender_email") or ""
    return f"{name} <{addr}>".strip() if addr else name


REQUEST_PHRASES = (
    "can you", "could you", "would you", "will you", "please", "pls",
    "let me know", "feedback", "follow up", "send me", "share", "review",
    "confirm", "approve", "approval", "need your", "need you to",
    "waiting for", "respond", "reply", "sign off", "sign-off",
)


def _email_dt(email: Dict[str, Any]) -> datetime.datetime:
    """Best-effort comparable timestamp for a formatted email record."""
    parsed = [
        dt for dt in (
            H._parse_received(email.get("received_time")),
            H._parse_received(email.get("sent_time")),
        )
        if dt is not None
    ]
    return max(parsed) if parsed else datetime.datetime.min


def _email_date(email: Dict[str, Any]) -> Optional[str]:
    return email.get("sent_time") or email.get("received_time")


def _email_from_me(email: Dict[str, Any], my_email: Optional[str]) -> bool:
    if "from_me" in email:
        return bool(email.get("from_me"))
    sender_email = (email.get("sender_email") or "").lower()
    return bool(my_email and my_email.lower() in sender_email)


def _message_text(email: Dict[str, Any]) -> str:
    return "\n".join([
        email.get("subject") or "",
        email.get("body") or "",
        email.get("snippet") or "",
    ]).lower()


def _message_has_request(email: Dict[str, Any]) -> bool:
    body = (email.get("body") or email.get("snippet") or "").lower()
    text = _message_text(email)
    return (
        "?" in body
        or any(phrase in text for phrase in REQUEST_PHRASES)
        or any(word in text for word in H.APPROVAL_KEYWORDS)
        or any(word in text for word in H.DEADLINE_KEYWORDS)
    )


def _base_subject(subject: Optional[str]) -> str:
    out = (subject or "(no subject)").strip() or "(no subject)"
    for _ in range(8):
        lowered = out.lower()
        for prefix in ("re:", "fw:", "fwd:"):
            if lowered.startswith(prefix):
                out = out[len(prefix):].strip() or "(no subject)"
                break
        else:
            return out
    return out


def _conversation_themes(messages: List[Dict[str, Any]]) -> List[str]:
    text = "\n".join(_message_text(m) for m in messages)
    themes: List[str] = []
    checks = (
        ("urgent", H.URGENT_KEYWORDS),
        ("deadline", H.DEADLINE_KEYWORDS),
        ("money", H.MONEY_KEYWORDS),
        ("approval", H.APPROVAL_KEYWORDS),
        ("meeting", H.MEETING_KEYWORDS),
    )
    for label, words in checks:
        if any(word in text for word in words):
            themes.append(label)
    if any("?" in (m.get("body") or m.get("snippet") or "") for m in messages):
        themes.append("questions")
    if any(m.get("has_attachments") for m in messages):
        themes.append("attachments")
    if any(m.get("unread") for m in messages):
        themes.append("unread")
    if any(m.get("flagged") == "flagged" for m in messages):
        themes.append("flagged")
    if any(m.get("importance") == "High" for m in messages):
        themes.append("high importance")
    return themes


def _compact_question(text: str) -> str:
    text = " ".join((text or "").split())
    if "?" not in text:
        return H.make_snippet(text, 180)
    for part in re.split(r"(?<=\?)\s+", text):
        if "?" in part:
            return H.make_snippet(part, 180)
    return H.make_snippet(text, 180)


def _open_questions(messages: List[Dict[str, Any]], my_email: Optional[str]) -> List[Dict[str, Any]]:
    questions: List[Dict[str, Any]] = []
    for email in reversed(messages):
        text = email.get("body") or email.get("snippet") or ""
        if "?" not in text:
            continue
        questions.append({
            "from": "You" if _email_from_me(email, my_email) else (_from_field(email) or "Someone"),
            "date": _email_date(email),
            "question": _compact_question(text),
        })
        if len(questions) >= 3:
            break
    return questions


def _age_days(email: Dict[str, Any], now: datetime.datetime) -> int:
    dt = _email_dt(email)
    if dt == datetime.datetime.min:
        return 0
    return max(0, int((now - dt).total_seconds() // 86400))


def _follow_up_state(
    messages: List[Dict[str, Any]],
    my_email: Optional[str],
    follow_up_days: int,
    now: datetime.datetime,
) -> Dict[str, Any]:
    last_outbound_idx: Optional[int] = None
    for idx, email in enumerate(messages):
        if _email_from_me(email, my_email):
            last_outbound_idx = idx

    inbound_after_last_outbound = messages[(last_outbound_idx + 1) if last_outbound_idx is not None else 0:]
    inbound_requests = [
        email for email in inbound_after_last_outbound
        if not _email_from_me(email, my_email) and _message_has_request(email)
    ]
    if inbound_requests:
        email = inbound_requests[-1]
        age = _age_days(email, now)
        due = age >= follow_up_days
        sender = _from_field(email) or "the sender"
        return {
            "state": "reply_owed",
            "direction": "them_to_you",
            "label": "Reply owed",
            "due": due,
            "age_days": age,
            "entry_id": email.get("id"),
            "hint": f"{sender} asked for something {age} day(s) ago; send a reply or explicitly close the loop.",
        }

    if last_outbound_idx is not None:
        email = messages[last_outbound_idx]
        if _message_has_request(email):
            age = _age_days(email, now)
            due = age >= follow_up_days
            return {
                "state": "waiting_on_them",
                "direction": "you_to_them",
                "label": "Waiting on them",
                "due": due,
                "age_days": age,
                "entry_id": email.get("id"),
                "hint": f"You asked for something {age} day(s) ago and no later response is in this conversation.",
            }

    latest = messages[-1] if messages else {}
    if latest and _email_from_me(latest, my_email):
        age = _age_days(latest, now)
        due = age >= follow_up_days
        return {
            "state": "last_from_you",
            "direction": "you_to_them",
            "label": "Last from you",
            "due": due,
            "age_days": age,
            "entry_id": latest.get("id"),
            "hint": "No later response is visible in this conversation.",
        }

    return {
        "state": "none",
        "direction": "none",
        "label": "No follow-up due",
        "due": False,
        "age_days": 0,
        "entry_id": None,
        "hint": "No unanswered request was detected in this conversation.",
    }


def _conversation_timeline(
    messages: List[Dict[str, Any]],
    my_email: Optional[str],
    limit: int = 5,
) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for email in messages[-limit:]:
        msg = _thread_msg(email)
        from_me = _email_from_me(email, my_email)
        msg["from_me"] = from_me
        msg["direction"] = "outbound" if from_me else "inbound"
        out.append(msg)
    return out


def _conversation_insight(
    messages: List[Dict[str, Any]],
    *,
    manager_name: Optional[str],
    my_email: Optional[str],
    follow_up_days: int,
    now: datetime.datetime,
) -> Dict[str, Any]:
    messages = sorted(messages, key=_email_dt)
    for email in messages:
        email["from_me"] = _email_from_me(email, my_email)

    latest = messages[-1]
    scored = H.rank_for_triage(messages, manager_name=manager_name,
                               my_email=my_email, now=now)
    max_score = max((int(e.get("triage_score") or 0) for e in scored), default=0)

    reasons: List[str] = []
    for email in scored:
        for reason in email.get("triage_reasons") or []:
            if reason not in reasons:
                reasons.append(reason)
            if len(reasons) >= 6:
                break
        if len(reasons) >= 6:
            break

    participants: List[str] = []
    for email in messages:
        person = "You" if _email_from_me(email, my_email) else _from_field(email)
        if person and person not in participants:
            participants.append(person)

    unread_count = sum(1 for email in messages if email.get("unread") and not _email_from_me(email, my_email))
    inbound_count = sum(1 for email in messages if not _email_from_me(email, my_email))
    outbound_count = len(messages) - inbound_count
    attachment_count = sum(int(email.get("attachment_count") or 0) for email in messages)
    themes = _conversation_themes(messages)
    follow_up = _follow_up_state(messages, my_email, follow_up_days, now)

    attention_score = max_score + min(unread_count, 3) * 2
    if len(messages) > 1:
        attention_score += min(len(messages), 5) - 1
    if follow_up["state"] == "reply_owed":
        attention_score += 8 if follow_up.get("due") else 5
        if "reply owed" not in reasons:
            reasons.insert(0, "reply owed")
    elif follow_up["state"] == "waiting_on_them":
        attention_score += 6 if follow_up.get("due") else 3
        if "waiting on response" not in reasons:
            reasons.insert(0, "waiting on response")
    elif follow_up["state"] == "last_from_you" and follow_up.get("due"):
        attention_score += 2
        if "last message was yours" not in reasons:
            reasons.append("last message was yours")

    latest_from = "You" if _email_from_me(latest, my_email) else _from_field(latest)
    if follow_up["state"] == "reply_owed":
        suggested = "Reply to the open request or mark it handled."
    elif follow_up["state"] == "waiting_on_them":
        suggested = "Send a follow-up if this is still needed."
    elif unread_count:
        suggested = "Read the latest unread message and decide whether to reply."
    elif attachment_count:
        suggested = "Review the attachment context before archiving."
    else:
        suggested = "No immediate action detected; keep for context or archive."

    theme_text = ", ".join(themes[:4]) if themes else "general mail"
    summary = (
        f"{len(messages)} message(s), {inbound_count} inbound and {outbound_count} outbound. "
        f"Latest from {latest_from}. Main signals: {theme_text}."
    )

    return {
        "conversation_id": latest.get("conversation_id") or latest.get("id"),
        "thread_id": latest.get("conversation_id"),
        "entry_id": latest.get("id"),
        "subject": _base_subject(latest.get("subject") or messages[0].get("subject")),
        "participants": participants[:8],
        "message_count": len(messages),
        "inbound_count": inbound_count,
        "outbound_count": outbound_count,
        "unread_count": unread_count,
        "attachment_count": attachment_count,
        "latest_from": latest_from,
        "latest_date": _email_date(latest),
        "latest_snippet": latest.get("snippet"),
        "attention_score": max(0, attention_score),
        "triage_score": max(0, attention_score),
        "triage_reasons": reasons[:7],
        "themes": themes,
        "follow_up": follow_up,
        "open_questions": _open_questions(messages, my_email),
        "insight_summary": summary,
        "suggested_action": suggested,
        "timeline": _conversation_timeline(messages, my_email),
        "_latest_record": latest,
    }


def _conversation_mailbox_summary(
    conversations: List[Dict[str, Any]],
    *,
    days: int,
    scanned: int,
    capped: bool,
) -> Dict[str, Any]:
    theme_counts: Counter = Counter()
    for conv in conversations:
        theme_counts.update(conv.get("themes") or [])
    return {
        "scan_window_days": days,
        "messages_scanned": scanned,
        "scan_capped": capped,
        "total_conversations": len(conversations),
        "reply_owed": sum(1 for c in conversations if c.get("follow_up", {}).get("state") == "reply_owed"),
        "waiting_on_them": sum(1 for c in conversations if c.get("follow_up", {}).get("state") == "waiting_on_them"),
        "unread_conversations": sum(1 for c in conversations if c.get("unread_count", 0) > 0),
        "high_attention": sum(1 for c in conversations if c.get("attention_score", 0) >= 10),
        "top_themes": [
            {"theme": theme, "count": count}
            for theme, count in theme_counts.most_common(6)
        ],
    }


def _fetch_emails(folder, days: int, *, unread_only: bool = False,
                  subject: Optional[str] = None, include_body: bool = True,
                  scan_cap: int = MAX_SCAN,
                  date_field: str = "ReceivedTime") -> Tuple[List[Dict[str, Any]], bool]:
    """Fetch + format emails newer than ``days`` from a folder, newest first.

    Optimised for busy mailboxes: pushes a date floor (and, when safe, an
    unread/subject filter) to Outlook via ``Restrict`` so the whole folder is
    never iterated in Python, and stops early once ``scan_cap`` items have been
    collected. Returns ``(emails, capped)`` where ``capped`` indicates the scan
    hit the cap (so totals are a lower bound).
    """
    now = datetime.datetime.now()
    threshold = now - datetime.timedelta(days=days)
    items = folder.Items
    try:
        items.Sort(f"[{date_field}]", True)  # newest first
    except Exception:
        pass

    # Push safe filters to Outlook (huge win on large folders). Falls back to
    # an unrestricted scan if the provider rejects the filter.
    if date_field == "ReceivedTime":
        restriction = H.build_inbox_restriction(
            threshold=threshold, unread_only=unread_only, subject=subject)
    else:
        restriction = f"[{date_field}] >= '{threshold.strftime('%m/%d/%Y %I:%M %p')}'"
    if restriction:
        try:
            items = items.Restrict(restriction)
        except Exception as exc:
            log.debug("Restrict failed (%s); scanning unfiltered.", exc)

    out: List[Dict[str, Any]] = []
    capped = False
    scanned = 0
    for item in items:
        scanned += 1
        if len(out) >= scan_cap:
            capped = True
            log.info("Scan cap (%d) reached in folder fetch.", scan_cap)
            break
        if scanned > MAX_SCAN:
            capped = True
            break
        try:
            received = (H._safe_getattr(item, date_field)
                        or H._safe_getattr(item, "ReceivedTime")
                        or H._safe_getattr(item, "SentOn"))
            if received is not None and received.replace(tzinfo=None) < threshold:
                # Sorted newest-first: once we cross the floor, stop.
                break
            out.append(H.format_email_item(item, include_body=include_body))
        except Exception as exc:
            log.debug("Skipping unreadable item: %s", exc)
            continue
    return out, capped


def _validate_days(days: int, maximum: int = H.MAX_DAYS) -> None:
    if not isinstance(days, int) or days < 1 or days > maximum:
        raise OutlookError(ErrorCode.INVALID_PARAMETER,
                           f"'days' must be an integer between 1 and {maximum}.")


def _set_body(mail, body: Optional[str], html_body: Optional[str]) -> None:
    if html_body:
        mail.HTMLBody = html_body
    elif body is not None:
        mail.Body = body


def _draft_payload(mail, action: str) -> Dict[str, Any]:
    """Structured payload describing a saved draft."""
    return make_success(
        action=action,
        status="draft_saved",
        draft_id=H._safe_getattr(mail, "EntryID"),
        to=H.extract_recipients(mail),
        subject=H._safe_getattr(mail, "Subject"),
        body_preview=H.make_snippet(H._safe_getattr(mail, "Body", "") or "", 400),
        next_safe_action="Review the draft, then call send_draft(draft_id, confirm=true) to send.",
    )


# ===========================================================================
# READ-ONLY TOOLS
# ===========================================================================

@mcp.tool()
@email_tool
def list_folders() -> str:
    """[READ-ONLY] List available Outlook mail folders (and sub-folders).

    Use this to discover valid ``folder_name`` values for search/list/move
    tools. Returns a JSON object: ``{"success": true, "folders": [...]}``.
    """
    namespace = _namespace()
    folders: List[Dict[str, Any]] = []

    def walk(folder, depth: int):
        if depth > 2:
            return
        try:
            unread = int(H._safe_getattr(folder.Items, "Count", 0) or 0) if False else None
        except Exception:
            unread = None
        folders.append({"name": folder.Name, "depth": depth})
        try:
            for sub in folder.Folders:
                walk(sub, depth + 1)
        except Exception:
            pass

    for root in namespace.Folders:
        walk(root, 0)
    return make_success(action="list_folders", count=len(folders), folders=folders)


@mcp.tool()
@email_tool
def search_emails(
    keyword: Optional[str] = None,
    sender: Optional[str] = None,
    subject: Optional[str] = None,
    recipient: Optional[str] = None,
    days: int = 7,
    unread_only: bool = False,
    has_attachments: Optional[bool] = None,
    category: Optional[str] = None,
    exact_phrase: Optional[str] = None,
    exclude: Optional[str] = None,
    folder_name: Optional[str] = None,
    max_results: int = H.DEFAULT_PAGE_SIZE,
    offset: int = 0,
) -> str:
    """[READ-ONLY] Search emails with structured filters; results newest-first.

    All filters are optional and AND-combined. This is the primary discovery
    tool — search first, then read with get_email_by_number / read_thread.

    Args:
        keyword: Free-text matched against subject, sender and body. Supports
            simple "a OR b" alternation (e.g. "invoice OR statement").
        sender: Substring matched against sender name/email (e.g. "john@acme").
        subject: Substring matched against the subject only.
        recipient: Substring matched against To/CC recipients.
        days: Look-back window in days (1–180, default 7). Newest first.
        unread_only: If true, only unread messages.
        has_attachments: If true/false, filter on attachment presence.
        category: Only emails carrying this Outlook category (label).
        exact_phrase: Require this exact phrase in subject or body.
        exclude: Comma-separated terms; emails containing any are dropped.
        folder_name: Folder to search (default Inbox). See list_folders.
        max_results: Page size (1–100, default 25).
        offset: Number of matches to skip (for pagination).

    Returns JSON: ``{success, query, results:[summary...], page_info}``. Each
    result has email_number (for this listing), entry_id (stable), thread_id,
    subject, from, to, date, snippet, labels, unread, has_attachments, etc.
    """
    _validate_days(days)
    namespace = _namespace()
    folder, folder_disp = _require_folder(namespace, folder_name)

    exclude_terms = [t.strip() for t in (exclude or "").split(",") if t.strip()]
    # Push the date floor + unread + subject narrowing to Outlook; keep the
    # Python predicate authoritative for the rest.
    emails, capped = _fetch_emails(
        folder, days, unread_only=unread_only, subject=subject,
        scan_cap=H.MAX_PAGE_SIZE * 8)
    matched = [
        e for e in emails
        if H.email_matches(
            e, keyword=keyword, sender=sender, subject=subject,
            recipient=recipient, unread_only=unread_only,
            has_attachments=has_attachments, category=category,
            exact_phrase=exact_phrase, exclude=exclude_terms or None,
        )
    ]
    _cache_listing(matched)
    page, page_info = H.paginate(matched, offset, max_results)
    page_info["capped"] = capped

    return make_success(
        action="search_emails",
        query={
            "keyword": keyword, "sender": sender, "subject": subject,
            "recipient": recipient, "days": days, "unread_only": unread_only,
            "has_attachments": has_attachments, "category": category,
            "exact_phrase": exact_phrase, "exclude": exclude_terms,
            "folder": folder_disp,
        },
        page_info=page_info,
        results=[_summarize(e) for e in page],
        next_safe_action="Use get_email_by_number(email_number) to read a result, "
                         "or read_thread to see the full conversation.",
    )


@mcp.tool()
@email_tool
def list_recent_emails(days: int = 7, folder_name: Optional[str] = None,
                       unread_only: bool = False,
                       max_results: int = H.DEFAULT_PAGE_SIZE,
                       offset: int = 0) -> str:
    """[READ-ONLY] List recent emails (newest first) from a folder.

    Args:
        days: Look-back window (1–180, default 7).
        folder_name: Folder to list (default Inbox).
        unread_only: Only unread messages when true.
        max_results: Page size (1–100, default 25).
        offset: Matches to skip for pagination.

    Returns the same summary shape as search_emails.
    """
    _validate_days(days)
    namespace = _namespace()
    folder, folder_disp = _require_folder(namespace, folder_name)
    emails, capped = _fetch_emails(folder, days, unread_only=unread_only,
                                   scan_cap=H.MAX_PAGE_SIZE * 8)
    _cache_listing(emails)
    page, page_info = H.paginate(emails, offset, max_results)
    page_info["capped"] = capped
    return make_success(
        action="list_recent_emails",
        folder=folder_disp, days=days, unread_only=unread_only,
        page_info=page_info,
        results=[_summarize(e) for e in page],
        next_safe_action="Use get_email_by_number(email_number) to read one.",
    )


@mcp.tool()
@email_tool
def get_unread_emails(days: int = 7, folder_name: Optional[str] = None,
                      max_results: int = H.DEFAULT_PAGE_SIZE,
                      offset: int = 0) -> str:
    """[READ-ONLY] List only unread emails (newest first). Convenience wrapper
    over list_recent_emails with unread_only=true."""
    return list_recent_emails(days=days, folder_name=folder_name,
                              unread_only=True, max_results=max_results,
                              offset=offset)


@mcp.tool()
@email_tool
def count_unread_emails(folder_name: Optional[str] = None) -> str:
    """[READ-ONLY] Count unread emails in a folder (default Inbox)."""
    namespace = _namespace()
    folder, folder_disp = _require_folder(namespace, folder_name)
    try:
        count = int(folder.Items.Restrict("[UnRead] = True").Count)
    except Exception as exc:
        raise OutlookError(ErrorCode.ACTION_FAILED,
                           "Could not count unread emails.",
                           details=str(exc) if H.debug_enabled() else None)
    return make_success(action="count_unread_emails", folder=folder_disp,
                        unread_count=count)


@mcp.tool()
@email_tool
def get_email_by_number(email_number: Optional[int] = None,
                        entry_id: Optional[str] = None) -> str:
    """[READ-ONLY] Read one email in full (body, recipients, attachments meta).

    Pass either ``email_number`` (from the most recent listing/search) or a
    stable ``entry_id``. Returns the full message including plain-text body,
    attachment metadata, labels and thread id.
    """
    namespace = _namespace()
    mail = _resolve_mail(namespace, email_number, entry_id)
    email = H.format_email_item(mail, include_body=True)
    attachments = _attachment_metadata(mail)
    payload = make_success(
        action="get_email_by_number",
        email_number=email_number,
        message_id=email.get("id"),
        thread_id=email.get("conversation_id"),
        subject=email.get("subject"),
        sender=_from_field(email),
        recipients=email.get("recipients"),
        date=email.get("received_time") or email.get("sent_time"),
        unread=email.get("unread"),
        labels=H.parse_categories(email.get("categories")),
        importance=email.get("importance"),
        flagged=email.get("flagged"),
        body=email.get("body"),
        attachments=attachments,
        attachment_count=len(attachments),
    )
    if email.get("warnings"):
        payload["warnings"] = email["warnings"]
    payload["next_safe_action"] = (
        "To respond, call reply_to_email_by_number (creates a draft by default). "
        "Sending requires an explicit send step."
    )
    return payload


@mcp.tool()
@email_tool
def read_thread(email_number: Optional[int] = None,
                entry_id: Optional[str] = None,
                days: int = 60, max_messages: int = 30) -> str:
    """[READ-ONLY] Read a full conversation thread in chronological order.

    Given any message in a thread (by email_number or entry_id), gathers the
    other messages in the same Outlook conversation from the Inbox and Sent
    Items within ``days``, sorted oldest→newest, so an agent can understand who
    said what before drafting a reply.

    Returns: thread_id, subject, participants, message_count, messages[] (each
    with from/date/snippet/unread), latest_message summary, and attachments
    seen across the thread.
    """
    _validate_days(days, H.MAX_DAYS)
    namespace = _namespace()
    anchor = _resolve_mail(namespace, email_number, entry_id)
    conv_id = H._safe_getattr(anchor, "ConversationID")
    if not conv_id:
        # Single-message "thread".
        e = H.format_email_item(anchor, include_body=True)
        return make_success(action="read_thread", thread_id=None,
                            subject=e.get("subject"), message_count=1,
                            participants=[_from_field(e)],
                            messages=[_thread_msg(e)],
                            warnings=["Message has no conversation id; returned alone."])

    threshold = datetime.datetime.now() - datetime.timedelta(days=days)
    collected: Dict[str, Dict[str, Any]] = {}
    for folder_idx in (H.OL_FOLDER_INBOX, H.OL_FOLDER_SENT):
        try:
            folder = namespace.GetDefaultFolder(folder_idx)
        except Exception:
            continue
        items = folder.Items
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass
        scanned = 0
        for item in items:
            scanned += 1
            if scanned > MAX_SCAN:
                break
            try:
                dt = H._safe_getattr(item, "ReceivedTime") or H._safe_getattr(item, "SentOn")
                if dt is not None and dt.replace(tzinfo=None) < threshold:
                    break
                if H._safe_getattr(item, "ConversationID") != conv_id:
                    continue
                e = H.format_email_item(item, include_body=True)
                if e.get("id"):
                    collected[e["id"]] = e
            except Exception:
                continue

    messages = sorted(
        collected.values(),
        key=lambda e: e.get("sent_time") or e.get("received_time") or "",
    )[:max_messages]

    participants: List[str] = []
    attachments: List[str] = []
    for e in messages:
        f = _from_field(e)
        if f and f not in participants:
            participants.append(f)
        if e.get("has_attachments"):
            attachments.append(e.get("subject") or "(no subject)")

    latest = messages[-1] if messages else H.format_email_item(anchor, include_body=True)
    return make_success(
        action="read_thread",
        thread_id=conv_id,
        subject=(messages[0].get("subject") if messages else None),
        participants=participants,
        message_count=len(messages),
        messages=[_thread_msg(e) for e in messages],
        latest_message={
            "from": _from_field(latest),
            "date": latest.get("received_time") or latest.get("sent_time"),
            "snippet": latest.get("snippet"),
            "unread": latest.get("unread"),
        },
        attachments_in_thread=attachments,
        next_safe_action="Draft a reply with reply_to_email_by_number; it will not send automatically.",
    )


def _thread_msg(e: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "entry_id": e.get("id"),
        "from": _from_field(e),
        "to": e.get("recipients"),
        "date": e.get("sent_time") or e.get("received_time"),
        "unread": e.get("unread"),
        "snippet": e.get("snippet"),
        "has_attachments": e.get("has_attachments"),
    }


# ===========================================================================
# ATTACHMENTS
# ===========================================================================

def _attachment_metadata(mail) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    att = H._safe_getattr(mail, "Attachments")
    if not att:
        return out
    try:
        count = int(H._safe_getattr(att, "Count", 0) or 0)
    except Exception:
        return out
    for i in range(1, count + 1):
        try:
            a = att(i)
            filename = H._safe_getattr(a, "FileName", "") or ""
            _, ext = os.path.splitext(filename.lower())
            out.append({
                "index": i,
                "filename": filename,
                "mime_type": ext.lstrip(".") or "unknown",
                "size_bytes": int(H._safe_getattr(a, "Size", 0) or 0),
                "is_readable_text": H.is_safe_text_attachment(filename),
            })
        except Exception:
            continue
    return out


@mcp.tool()
@email_tool
def list_attachments(email_number: Optional[int] = None,
                     entry_id: Optional[str] = None) -> str:
    """[READ-ONLY] List attachments on an email (filename, type, size, index).

    Use the returned ``index`` with save_attachment / read_attachment.
    """
    namespace = _namespace()
    mail = _resolve_mail(namespace, email_number, entry_id)
    attachments = _attachment_metadata(mail)
    return make_success(action="list_attachments",
                        email_number=email_number,
                        attachment_count=len(attachments),
                        attachments=attachments)


@mcp.tool()
@email_tool
def save_attachment(attachment_index: int, email_number: Optional[int] = None,
                    entry_id: Optional[str] = None,
                    destination_dir: Optional[str] = None) -> str:
    """[WRITES FILE] Save one attachment to disk (no execution, size-limited).

    Saves the attachment at ``attachment_index`` (from list_attachments) to
    ``destination_dir`` (or the OUTLOOK_ATTACHMENT_DIR env var, else the system
    temp dir). The filename is sanitised and never executed. Attachments larger
    than OUTLOOK_MAX_ATTACHMENT_MB (default 25 MB) are refused.

    Returns the saved file path and metadata.
    """
    namespace = _namespace()
    mail = _resolve_mail(namespace, email_number, entry_id)
    a = _get_attachment(mail, attachment_index)
    filename = H._safe_getattr(a, "FileName", "") or f"attachment_{attachment_index}"
    size = int(H._safe_getattr(a, "Size", 0) or 0)
    limit = H.max_attachment_bytes()
    if size > limit:
        raise OutlookError(ErrorCode.ATTACHMENT_TOO_LARGE,
                           f"Attachment is {size} bytes, over the {limit}-byte limit.",
                           details="Raise OUTLOOK_MAX_ATTACHMENT_MB to allow larger files.")
    dest_dir = destination_dir or os.environ.get("OUTLOOK_ATTACHMENT_DIR") or _temp_dir()
    os.makedirs(dest_dir, exist_ok=True)
    safe_name = H.sanitize_filename(filename)
    path = os.path.join(dest_dir, safe_name)
    try:
        a.SaveAsFile(path)
    except Exception as exc:
        raise OutlookError(ErrorCode.ACTION_FAILED, "Failed to save attachment.",
                           details=str(exc) if H.debug_enabled() else None)
    return make_success(action="save_attachment", filename=safe_name,
                        saved_path=path, size_bytes=size,
                        warning="File saved but NOT opened or executed.")


@mcp.tool()
@email_tool
def read_attachment(attachment_index: int, email_number: Optional[int] = None,
                    entry_id: Optional[str] = None,
                    max_chars: int = 20000) -> str:
    """[READ-ONLY] Read a *text* attachment's contents inline.

    Only safe text types (.txt/.csv/.md/.json/.xml/.html/...) under the size
    limit are read; binaries are refused. Content is truncated to ``max_chars``.
    """
    namespace = _namespace()
    mail = _resolve_mail(namespace, email_number, entry_id)
    a = _get_attachment(mail, attachment_index)
    filename = H._safe_getattr(a, "FileName", "") or ""
    if not H.is_safe_text_attachment(filename):
        raise OutlookError(ErrorCode.UNSUPPORTED_ATTACHMENT,
                           f"'{filename}' is not a supported text attachment.",
                           details="Use save_attachment to download non-text files.")
    size = int(H._safe_getattr(a, "Size", 0) or 0)
    if size > H.max_attachment_bytes():
        raise OutlookError(ErrorCode.ATTACHMENT_TOO_LARGE,
                           "Attachment exceeds the configured size limit.")
    tmp_path = os.path.join(_temp_dir(), H.sanitize_filename(filename) or "att.tmp")
    try:
        a.SaveAsFile(tmp_path)
        with open(tmp_path, "r", encoding="utf-8", errors="replace") as fh:
            content = fh.read(max_chars + 1)
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
    truncated = len(content) > max_chars
    return make_success(action="read_attachment", filename=filename,
                        truncated=truncated, content=content[:max_chars])


def _get_attachment(mail, index: int):
    att = H._safe_getattr(mail, "Attachments")
    count = int(H._safe_getattr(att, "Count", 0) or 0) if att else 0
    if not att or index < 1 or index > count:
        raise OutlookError(ErrorCode.ATTACHMENT_NOT_FOUND,
                           f"Attachment #{index} does not exist on this email.",
                           details=f"This email has {count} attachment(s).")
    return att(index)


def _temp_dir() -> str:
    import tempfile
    return os.environ.get("OUTLOOK_ATTACHMENT_DIR") or tempfile.gettempdir()


# ===========================================================================
# DRAFTS  (safe – never send)
# ===========================================================================

@mcp.tool()
@email_tool
def create_draft(to: str, subject: str, body: str,
                 cc: Optional[str] = None, bcc: Optional[str] = None,
                 html_body: Optional[str] = None) -> str:
    """[WRITES DRAFT — does not send] Create a new email draft in Outlook.

    This is the preferred way to compose. It validates recipients and saves a
    draft; nothing is sent until you call send_draft(draft_id, confirm=true).

    Args:
        to: Recipient address(es), comma/semicolon separated.
        subject: Subject line.
        body: Plain-text body (preferred, always set).
        cc / bcc: Optional carbon-copy / blind-copy recipients.
        html_body: Optional HTML body; if given it is used instead of body.

    Returns draft_id and a preview.
    """
    outlook, _ = _connect()
    _validate_outbound(to, cc, bcc)
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.To = _join(to)
    if cc:
        mail.CC = _join(cc)
    if bcc:
        mail.BCC = _join(bcc)
    _set_body(mail, body, html_body)
    mail.Save()
    return _draft_payload(mail, "create_draft")


@mcp.tool()
@email_tool
def update_draft(draft_id: str, to: Optional[str] = None,
                 subject: Optional[str] = None, body: Optional[str] = None,
                 cc: Optional[str] = None, bcc: Optional[str] = None,
                 html_body: Optional[str] = None) -> str:
    """[WRITES DRAFT — does not send] Edit an existing draft.

    Only the fields you pass are changed. Recipient changes are validated.
    """
    namespace = _namespace()
    mail = _resolve_draft(namespace, draft_id)
    if to is not None or cc is not None or bcc is not None:
        _validate_outbound(to or _addr_of(mail.To), cc, bcc)
    if to is not None:
        mail.To = _join(to)
    if cc is not None:
        mail.CC = _join(cc)
    if bcc is not None:
        mail.BCC = _join(bcc)
    if subject is not None:
        mail.Subject = subject
    if body is not None or html_body is not None:
        _set_body(mail, body, html_body)
    mail.Save()
    return _draft_payload(mail, "update_draft")


@mcp.tool()
@email_tool
def list_drafts(max_results: int = H.DEFAULT_PAGE_SIZE) -> str:
    """[READ-ONLY] List saved drafts (most recent first)."""
    namespace = _namespace()
    folder = namespace.GetDefaultFolder(H.OL_FOLDER_DRAFTS)
    items = folder.Items
    try:
        items.Sort("[LastModificationTime]", True)
    except Exception:
        pass
    drafts: List[Dict[str, Any]] = []
    for item in items:
        if len(drafts) >= max(1, min(max_results, H.MAX_PAGE_SIZE)):
            break
        try:
            drafts.append({
                "draft_id": H._safe_getattr(item, "EntryID"),
                "to": H.extract_recipients(item),
                "subject": H._safe_getattr(item, "Subject"),
                "body_preview": H.make_snippet(H._safe_getattr(item, "Body", "") or ""),
            })
        except Exception:
            continue
    return make_success(action="list_drafts", count=len(drafts), drafts=drafts)


@mcp.tool()
@email_tool
def send_draft(draft_id: str, confirm: bool = False) -> str:
    """[SENDS EMAIL — explicit] Send a previously created draft.

    Safety guard: with ``confirm=false`` (default) this only returns a preview
    of what would be sent. Re-call with ``confirm=true`` to actually send.
    Requires explicit user intent to send.
    """
    namespace = _namespace()
    mail = _resolve_draft(namespace, draft_id)
    recipients = H.extract_recipients(mail)
    if not recipients:
        raise OutlookError(ErrorCode.INVALID_RECIPIENT,
                           "Draft has no recipients.",
                           details="Add recipients with update_draft before sending.")
    if not confirm:
        return make_error(
            ErrorCode.CONFIRMATION_REQUIRED,
            "Confirmation required before sending.",
            details="Re-call send_draft with confirm=true to send this draft.",
            preview={"to": recipients,
                     "subject": H._safe_getattr(mail, "Subject"),
                     "body_preview": H.make_snippet(H._safe_getattr(mail, "Body", "") or "", 400)},
        )
    mail.Send()
    log.info("Draft %s sent to %d recipient(s).", draft_id, len(recipients))
    return make_success(action="send_draft", status="sent",
                        sent_to=recipients,
                        subject="(sent)" )


@mcp.tool()
@email_tool
def delete_draft(draft_id: str) -> str:
    """[SAFE DELETE] Delete a draft (moves it to Deleted Items). Drafts only —
    this refuses to delete sent or received mail."""
    namespace = _namespace()
    mail = _resolve_draft(namespace, draft_id)
    subject = H._safe_getattr(mail, "Subject")
    mail.Delete()
    return make_success(action="delete_draft", status="deleted",
                        subject=subject)


def _resolve_draft(namespace, draft_id: str):
    try:
        mail = namespace.GetItemFromID(draft_id)
    except Exception:
        mail = None
    if not mail:
        raise OutlookError(ErrorCode.DRAFT_NOT_FOUND,
                           "No draft matches that draft_id.",
                           details="Call list_drafts to see current drafts.")
    # Sent==False indicates an unsent (draft) item.
    if H._safe_getattr(mail, "Sent", True):
        raise OutlookError(ErrorCode.DRAFT_NOT_FOUND,
                           "That item is not an editable draft (already sent/received).",
                           details="Use forward_email or reply_to_email_by_number instead.")
    return mail


# ===========================================================================
# SENDING  (explicit + confirmed)
# ===========================================================================

@mcp.tool()
@email_tool
def send_email(to: str, subject: str, body: str, cc: Optional[str] = None,
               bcc: Optional[str] = None, html_body: Optional[str] = None,
               confirm: bool = False) -> str:
    """[SENDS EMAIL — explicit] Compose and send a NEW email immediately.

    Only use this when the user has clearly asked to *send*. Prefer create_draft
    otherwise. Safety guard: ``confirm=false`` (default) returns a preview only;
    re-call with ``confirm=true`` to actually send. Recipients are validated.
    """
    outlook, _ = _connect()
    _validate_outbound(to, cc, bcc)
    if not confirm:
        return make_error(
            ErrorCode.CONFIRMATION_REQUIRED,
            "Confirmation required before sending a new email.",
            details="Re-call send_email with confirm=true, or use create_draft to stage it.",
            preview={"to": H.parse_recipients(to), "cc": H.parse_recipients(cc),
                     "subject": subject, "body_preview": H.make_snippet(body, 400)},
        )
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.To = _join(to)
    if cc:
        mail.CC = _join(cc)
    if bcc:
        mail.BCC = _join(bcc)
    _set_body(mail, body, html_body)
    mail.Send()
    log.info("New email sent to %d recipient(s).", len(H.parse_recipients(to)))
    return make_success(action="send_email", status="sent",
                        sent_to=H.parse_recipients(to),
                        cc=H.parse_recipients(cc), subject=subject)


@mcp.tool()
@email_tool
def compose_email(recipient_email: str, subject: str, body: str,
                  cc_email: Optional[str] = None, send: bool = False) -> str:
    """[LEGACY] Compose an email. Drafts by default; only sends if send=true.

    Kept for backward compatibility. New agents should prefer create_draft (to
    stage) and send_email (to send, with confirm). With ``send=false`` (default)
    this saves a draft and is completely safe.
    """
    if send:
        return send_email(to=recipient_email, subject=subject, body=body,
                          cc=cc_email, confirm=True)
    return create_draft(to=recipient_email, subject=subject, body=body, cc=cc_email)


@mcp.tool()
@email_tool
def reply_to_email_by_number(email_number: Optional[int] = None,
                             reply_text: str = "",
                             entry_id: Optional[str] = None,
                             reply_all: bool = False,
                             send: bool = False) -> str:
    """[WRITES DRAFT by default] Reply to an email, preserving the quoted thread.

    By default this creates a DRAFT reply (safe) and returns its draft_id —
    review it, then send_draft(draft_id, confirm=true). Pass ``send=true`` only
    when the user explicitly asked to send the reply in one step.

    Args:
        email_number: Target email from the latest listing (or use entry_id).
        reply_text: Your reply text, inserted above the quoted original.
        entry_id: Stable id alternative to email_number.
        reply_all: Reply to all recipients instead of just the sender.
        send: If true, send immediately; otherwise save as a draft (default).
    """
    if not reply_text or not reply_text.strip():
        raise OutlookError(ErrorCode.INVALID_PARAMETER, "reply_text is required.")
    namespace = _namespace()
    mail = _resolve_mail(namespace, email_number, entry_id)
    reply = mail.ReplyAll() if reply_all else mail.Reply()
    # Insert our text above Outlook's quoted original (preserved in reply.Body).
    reply.Body = f"{reply_text}\n\n{H._safe_getattr(reply, 'Body', '') or ''}"
    if send:
        recipients = H.extract_recipients(reply)
        reply.Send()
        log.info("Reply sent (reply_all=%s).", reply_all)
        return make_success(action="reply_to_email", status="sent",
                            reply_all=reply_all, sent_to=recipients)
    reply.Save()
    payload = _draft_payload(reply, "reply_to_email")
    payload["reply_all"] = reply_all
    return payload


@mcp.tool()
@email_tool
def forward_email(to: str, email_number: Optional[int] = None,
                  entry_id: Optional[str] = None, comment: str = "",
                  send: bool = False) -> str:
    """[WRITES DRAFT by default] Forward an email (attachments preserved).

    Creates a DRAFT forward by default (safe); review then send_draft. Pass
    ``send=true`` to forward immediately when explicitly requested. Recipients
    are validated.

    Args:
        to: Where to forward, comma/semicolon separated.
        email_number / entry_id: The email to forward.
        comment: Optional note added above the forwarded content.
        send: Send now (true) or save draft (false, default).
    """
    namespace = _namespace()
    _validate_outbound(to, None, None)
    mail = _resolve_mail(namespace, email_number, entry_id)
    fwd = mail.Forward()  # preserves the original body and attachments
    fwd.To = _join(to)
    if comment:
        fwd.Body = f"{comment}\n\n{H._safe_getattr(fwd, 'Body', '') or ''}"
    if send:
        fwd.Send()
        log.info("Email forwarded to %d recipient(s).", len(H.parse_recipients(to)))
        return make_success(action="forward_email", status="sent",
                            sent_to=H.parse_recipients(to))
    fwd.Save()
    return _draft_payload(fwd, "forward_email")


def _validate_outbound(to: Optional[str], cc: Optional[str], bcc: Optional[str]) -> None:
    """Validate all outbound recipients; raise INVALID_RECIPIENT on any bad one."""
    invalid_all: List[str] = []
    any_valid = False
    for field in (to, cc, bcc):
        valid, invalid = H.validate_recipients(field)
        any_valid = any_valid or bool(valid)
        invalid_all.extend(invalid)
    if invalid_all:
        raise OutlookError(
            ErrorCode.INVALID_RECIPIENT,
            "One or more recipient addresses are invalid.",
            details=f"Invalid: {', '.join(invalid_all)}",
            invalid_recipients=invalid_all,
        )
    if not any_valid:
        raise OutlookError(ErrorCode.INVALID_RECIPIENT,
                           "At least one valid recipient is required.")


def _join(value: Optional[str]) -> str:
    return "; ".join(H.parse_recipients(value))


def _addr_of(value: Any) -> str:
    return str(value or "")


# ===========================================================================
# ORGANISE  (move / archive / trash / read state / labels)
# ===========================================================================

@mcp.tool()
@email_tool
def move_email_by_number(destination_folder_name: str,
                         email_number: Optional[int] = None,
                         entry_id: Optional[str] = None) -> str:
    """[MOVES EMAIL] Move an email to another folder.

    Args:
        destination_folder_name: Exact target folder (see list_folders).
        email_number / entry_id: The email to move.
    """
    if not destination_folder_name:
        raise OutlookError(ErrorCode.INVALID_PARAMETER,
                           "destination_folder_name is required.")
    namespace = _namespace()
    mail = _resolve_mail(namespace, email_number, entry_id)
    folder, _ = _require_folder(namespace, destination_folder_name)
    subject = H._safe_getattr(mail, "Subject")
    mail.Move(folder)
    if email_number in _email_cache:
        del _email_cache[email_number]
    return make_success(action="move_email", status="moved",
                        subject=subject, destination=destination_folder_name,
                        affected_count=1)


def _conversation_mail_items(namespace, anchor, days: int = H.MAX_DAYS) -> List[Any]:
    """Return live Outlook items in the same conversation from Inbox + Sent."""
    conv_id = H._safe_getattr(anchor, "ConversationID")
    if not conv_id:
        return [anchor]

    threshold = datetime.datetime.now() - datetime.timedelta(days=days)
    collected: Dict[str, Any] = {}
    for folder_idx, sort_field in ((H.OL_FOLDER_INBOX, "ReceivedTime"),
                                   (H.OL_FOLDER_SENT, "SentOn")):
        try:
            folder = namespace.GetDefaultFolder(folder_idx)
            items = folder.Items
            try:
                items.Sort(f"[{sort_field}]", True)
            except Exception:
                pass
            scanned = 0
            for item in list(items):
                scanned += 1
                if scanned > MAX_SCAN:
                    break
                try:
                    dt = (H._safe_getattr(item, sort_field)
                          or H._safe_getattr(item, "ReceivedTime")
                          or H._safe_getattr(item, "SentOn"))
                    if dt is not None and dt.replace(tzinfo=None) < threshold:
                        break
                    if H._safe_getattr(item, "ConversationID") != conv_id:
                        continue
                    item_id = H._safe_getattr(item, "EntryID") or f"{folder_idx}:{scanned}"
                    collected[item_id] = item
                except Exception:
                    continue
        except Exception:
            continue
    return list(collected.values()) or [anchor]


def _clear_cache_entries(entry_ids: List[str]) -> None:
    entry_id_set = set(entry_ids)
    for number, record in list(_email_cache.items()):
        if record.get("id") in entry_id_set:
            del _email_cache[number]


@mcp.tool()
@email_tool
def attend_conversation(entry_id: str, days: int = H.MAX_DAYS,
                        destination_folder_name: Optional[str] = None) -> str:
    """[MOVES EMAILS] Move every found message in a conversation to Attended.

    The destination defaults to the ``OUTLOOK_ATTENDED_FOLDER`` environment
    variable, or ``Attended``. Matching messages are gathered from Inbox and
    Sent Items within the supplied window, marked read, then moved together.
    """
    if not entry_id:
        raise OutlookError(ErrorCode.INVALID_PARAMETER, "entry_id is required.")
    _validate_days(days, H.MAX_DAYS)
    destination = destination_folder_name or os.environ.get("OUTLOOK_ATTENDED_FOLDER", "Attended")
    namespace = _namespace()
    anchor = _resolve_mail(namespace, entry_id=entry_id)
    folder = _get_folder_by_name(namespace, destination)
    if not folder:
        raise OutlookError(
            ErrorCode.FOLDER_NOT_FOUND,
            f"Attended folder '{destination}' was not found.",
            details="Create an Outlook folder named 'Attended' or set OUTLOOK_ATTENDED_FOLDER.",
        )

    items = _conversation_mail_items(namespace, anchor, days)
    moved_ids: List[str] = []
    subjects: List[str] = []
    for item in items:
        try:
            item_id = H._safe_getattr(item, "EntryID")
            subject = H._safe_getattr(item, "Subject")
            try:
                item.UnRead = False
                item.Save()
            except Exception:
                pass
            item.Move(folder)
            if item_id:
                moved_ids.append(item_id)
            if subject and subject not in subjects:
                subjects.append(subject)
        except Exception as exc:
            log.debug("Could not move conversation item to Attended: %s", exc)
            continue

    if not moved_ids:
        raise OutlookError(ErrorCode.ACTION_FAILED,
                           "No messages in the conversation could be moved.",
                           retryable=True)
    _clear_cache_entries(moved_ids)
    return make_success(
        action="attend_conversation",
        status="attended",
        destination=destination,
        affected_count=len(moved_ids),
        moved_entry_ids=moved_ids,
        subject=subjects[0] if subjects else H._safe_getattr(anchor, "Subject"),
    )


@mcp.tool()
@email_tool
def archive_email(email_number: Optional[int] = None,
                  entry_id: Optional[str] = None) -> str:
    """[MOVES EMAIL] Archive an email by moving it to the Archive folder.

    The target folder name defaults to "Archive" (override with the
    OUTLOOK_ARCHIVE_FOLDER env var). Prefer archiving over deleting.
    """
    folder_name = os.environ.get("OUTLOOK_ARCHIVE_FOLDER", "Archive")
    namespace = _namespace()
    mail = _resolve_mail(namespace, email_number, entry_id)
    folder = _get_folder_by_name(namespace, folder_name)
    if not folder:
        raise OutlookError(
            ErrorCode.FOLDER_NOT_FOUND,
            f"Archive folder '{folder_name}' was not found.",
            details="Create an 'Archive' folder in Outlook or set OUTLOOK_ARCHIVE_FOLDER.",
        )
    subject = H._safe_getattr(mail, "Subject")
    mail.Move(folder)
    if email_number in _email_cache:
        del _email_cache[email_number]
    return make_success(action="archive_email", status="archived",
                        subject=subject, destination=folder_name,
                        affected_count=1)


@mcp.tool()
@email_tool
def move_to_trash(email_number: Optional[int] = None,
                  entry_id: Optional[str] = None) -> str:
    """[DESTRUCTIVE — recoverable] Move an email to Deleted Items (the trash).

    This does NOT permanently delete: the message goes to Deleted Items and can
    be recovered. There is no permanent-delete tool by design.
    """
    namespace = _namespace()
    mail = _resolve_mail(namespace, email_number, entry_id)
    deleted = namespace.GetDefaultFolder(H.OL_FOLDER_DELETED_ITEMS)
    subject = H._safe_getattr(mail, "Subject")
    mail.Move(deleted)
    if email_number in _email_cache:
        del _email_cache[email_number]
    return make_success(action="move_to_trash", status="moved_to_deleted_items",
                        subject=subject, affected_count=1,
                        note="Recoverable from Deleted Items; not permanently deleted.")


@mcp.tool()
@email_tool
def mark_as_read(email_number: Optional[int] = None,
                 entry_id: Optional[str] = None) -> str:
    """[UPDATES STATE] Mark an email as read."""
    return _set_read_state(email_number, entry_id, read=True)


@mcp.tool()
@email_tool
def mark_as_unread(email_number: Optional[int] = None,
                   entry_id: Optional[str] = None) -> str:
    """[UPDATES STATE] Mark an email as unread."""
    return _set_read_state(email_number, entry_id, read=False)


def _set_read_state(email_number, entry_id, read: bool):
    namespace = _namespace()
    mail = _resolve_mail(namespace, email_number, entry_id)
    mail.UnRead = not read
    mail.Save()
    return make_success(action="mark_as_read" if read else "mark_as_unread",
                        status="read" if read else "unread",
                        subject=H._safe_getattr(mail, "Subject"))


@mcp.tool()
@email_tool
def list_categories() -> str:
    """[READ-ONLY] List Outlook categories (the closest thing to labels)."""
    namespace = _namespace()
    cats: List[Dict[str, Any]] = []
    try:
        for c in namespace.Categories:
            cats.append({"name": H._safe_getattr(c, "Name"),
                         "color": H._safe_getattr(c, "Color")})
    except Exception as exc:
        raise OutlookError(ErrorCode.ACTION_FAILED, "Could not read categories.",
                           details=str(exc) if H.debug_enabled() else None)
    return make_success(action="list_categories", count=len(cats), categories=cats)


@mcp.tool()
@email_tool
def apply_category(category: str, email_number: Optional[int] = None,
                   entry_id: Optional[str] = None) -> str:
    """[UPDATES LABEL] Add an Outlook category (label) to an email."""
    if not category or not category.strip():
        raise OutlookError(ErrorCode.INVALID_PARAMETER, "category is required.")
    namespace = _namespace()
    mail = _resolve_mail(namespace, email_number, entry_id)
    cats = H.parse_categories(H._safe_getattr(mail, "Categories", ""))
    if category not in cats:
        cats.append(category)
    mail.Categories = H.join_categories(cats)
    mail.Save()
    return make_success(action="apply_category", category=category,
                        labels=cats, subject=H._safe_getattr(mail, "Subject"))


@mcp.tool()
@email_tool
def remove_category(category: str, email_number: Optional[int] = None,
                    entry_id: Optional[str] = None) -> str:
    """[UPDATES LABEL] Remove an Outlook category (label) from an email."""
    namespace = _namespace()
    mail = _resolve_mail(namespace, email_number, entry_id)
    cats = [c for c in H.parse_categories(H._safe_getattr(mail, "Categories", ""))
            if c.lower() != category.strip().lower()]
    mail.Categories = H.join_categories(cats)
    mail.Save()
    return make_success(action="remove_category", category=category,
                        labels=cats, subject=H._safe_getattr(mail, "Subject"))


# ===========================================================================
# AI-ANALYSIS AGGREGATORS  (data providers – the agent does the reasoning)
# ===========================================================================

@mcp.tool()
@email_tool
def prioritize_inbox(days: int = 1, max_emails_to_scan: int = 25) -> str:
    """[READ-ONLY] Gather recent inbox emails for the agent to triage.

    Returns raw per-email data (sender, subject, snippet, importance) for you to
    rank — it does not rank for you. Use for "what needs my attention?".

    Args:
        days: Look-back window (1–31, default 1).
        max_emails_to_scan: Cap on emails returned (5–50, default 25).
    """
    if not isinstance(days, int) or not 1 <= days <= 31:
        raise OutlookError(ErrorCode.INVALID_PARAMETER,
                           "'days' must be an integer between 1 and 31.")
    if not isinstance(max_emails_to_scan, int) or not 5 <= max_emails_to_scan <= 50:
        raise OutlookError(ErrorCode.INVALID_PARAMETER,
                           "'max_emails_to_scan' must be between 5 and 50.")
    namespace = _namespace()
    manager = _get_manager_name(namespace)
    inbox = namespace.GetDefaultFolder(H.OL_FOLDER_INBOX)
    emails, _capped = _fetch_emails(inbox, days, scan_cap=max_emails_to_scan)
    emails = emails[:max_emails_to_scan]
    _cache_listing(emails)
    out = []
    for e in emails:
        out.append({
            "email_number": e["email_number"],
            "entry_id": e.get("id"),
            "sender": e.get("sender"),
            "subject": e.get("subject"),
            "snippet": e.get("snippet"),
            "received_time": e.get("received_time"),
            "unread": e.get("unread"),
            "importance": e.get("importance"),
            "is_from_manager": bool(manager and manager.lower() in (e.get("sender") or "").lower()),
        })
    return make_success(action="prioritize_inbox", count=len(out), emails=out,
                        analysis_instructions="Rank these by urgency/importance and explain why.")


@mcp.tool()
@email_tool
def triage_inbox(days: int = 3, max_results: int = 20,
                 folder_name: Optional[str] = None,
                 unread_only: bool = False) -> str:
    """[READ-ONLY] Rank what needs attention using a deterministic scorer.

    Unlike prioritize_inbox (which hands raw data to the agent), this returns a
    ready-made, explainable ranking — most-urgent first — using transparent
    rules (unread, sender importance, urgent/deadline/money/approval language,
    questions, recency; automated/bulk mail is down-ranked). Ideal for
    "what should I deal with first?" and powers the dashboard + scheduled digest.

    Args:
        days: Look-back window (1–60, default 3).
        max_results: How many ranked emails to return (1–100, default 20).
        folder_name: Folder to triage (default Inbox).
        unread_only: Only consider unread mail when true.

    Each result includes triage_score and triage_reasons plus the usual summary
    fields (entry_id, subject, from, date, snippet, labels, …).
    """
    _validate_days(days, H.ACTIONABLE_EMAIL_MAX_DAYS)
    namespace = _namespace()
    manager = _get_manager_name(namespace)
    my_email = _get_my_email(namespace)
    folder, folder_disp = _require_folder(namespace, folder_name)
    emails, capped = _fetch_emails(folder, days, unread_only=unread_only,
                                   scan_cap=H.MAX_PAGE_SIZE * 8)
    ranked = H.rank_for_triage(emails, manager_name=manager, my_email=my_email)
    _cache_listing(ranked)
    page = ranked[:max(1, min(max_results, H.MAX_PAGE_SIZE))]
    results = []
    for e in page:
        summary = _summarize(e)
        summary["triage_score"] = e.get("triage_score")
        summary["triage_reasons"] = e.get("triage_reasons")
        results.append(summary)
    return make_success(
        action="triage_inbox", folder=folder_disp, days=days,
        scanned=len(emails), capped=capped, count=len(results),
        results=results,
        next_safe_action="Open the top item with get_email_by_number, then draft a "
                         "reply or archive it. Sending always needs confirmation.",
    )


@mcp.tool()
@email_tool
def conversation_insights(
    days: int = H.MAX_DAYS,
    max_results: int = 40,
    offset: int = 0,
    folder_name: Optional[str] = None,
    unread_only: bool = False,
    keyword: Optional[str] = None,
    sender: Optional[str] = None,
    subject: Optional[str] = None,
    recipient: Optional[str] = None,
    has_attachments: Optional[bool] = None,
    category: Optional[str] = None,
    exact_phrase: Optional[str] = None,
    exclude: Optional[str] = None,
    include_sent: bool = True,
    follow_up_days: int = 3,
) -> str:
    """[READ-ONLY] Deep, conversation-level mailbox insight for the web manager.

    Scans a broad window (default: the maximum supported look-back), gathers
    Inbox plus Sent Items, groups by Outlook conversation id, and returns one
    ranked insight object per conversation instead of one card per message.
    Follow-up hints are bidirectional: the tool flags both conversations where
    you asked for something and have not received feedback, and conversations
    where someone asked you for something and you have not responded.
    """
    _validate_days(days, H.MAX_DAYS)
    if not isinstance(follow_up_days, int) or follow_up_days < 1:
        raise OutlookError(ErrorCode.INVALID_PARAMETER,
                           "'follow_up_days' must be a positive integer.")

    namespace = _namespace()
    manager = _get_manager_name(namespace)
    my_email = _get_my_email(namespace)
    folder, folder_disp = _require_folder(namespace, folder_name)
    max_results = max(1, min(int(max_results), H.MAX_PAGE_SIZE))
    scan_cap = min(MAX_SCAN, max(400, max_results * 12))

    primary_emails, primary_capped = _fetch_emails(
        folder, days, include_body=True, scan_cap=scan_cap)

    records: Dict[str, Dict[str, Any]] = {}

    def add_record(email: Dict[str, Any], source: str, from_me: Optional[bool] = None) -> None:
        record = dict(email)
        record["source_folder"] = source
        if from_me is not None:
            record["from_me"] = from_me
        else:
            record["from_me"] = _email_from_me(record, my_email)
        key = record.get("id") or f"{source}:{len(records)}"
        records[key] = record

    for email in primary_emails:
        add_record(email, folder_disp)

    sent_capped = False
    if include_sent:
        try:
            sent = namespace.GetDefaultFolder(H.OL_FOLDER_SENT)
            sent_emails, sent_capped = _fetch_emails(
                sent, days, include_body=True, scan_cap=scan_cap,
                date_field="SentOn")
            for email in sent_emails:
                add_record(email, "Sent Items", from_me=True)
        except Exception as exc:
            log.debug("Could not scan Sent Items for conversation insights: %s", exc)

    exclude_terms = [t.strip() for t in (exclude or "").split(",") if t.strip()]
    filters_active = any([
        keyword, sender, subject, recipient, unread_only,
        has_attachments is not None, category, exact_phrase, exclude_terms,
    ])

    def matches_filters(email: Dict[str, Any]) -> bool:
        return H.email_matches(
            email, keyword=keyword, sender=sender, subject=subject,
            recipient=recipient, unread_only=unread_only,
            has_attachments=has_attachments, category=category,
            exact_phrase=exact_phrase, exclude=exclude_terms or None,
        )

    groups: Dict[str, List[Dict[str, Any]]] = {}
    for email in records.values():
        key = email.get("conversation_id") or email.get("id") or f"message:{len(groups)}"
        groups.setdefault(key, []).append(email)

    now = datetime.datetime.now()
    conversations: List[Dict[str, Any]] = []
    for messages in groups.values():
        if filters_active and not any(matches_filters(email) for email in messages):
            continue
        conversations.append(_conversation_insight(
            messages, manager_name=manager, my_email=my_email,
            follow_up_days=follow_up_days, now=now))

    conversations.sort(
        key=lambda conv: (conv.get("attention_score", 0), conv.get("latest_date") or ""),
        reverse=True,
    )
    summary = _conversation_mailbox_summary(
        conversations, days=days, scanned=len(records),
        capped=primary_capped or sent_capped)

    page, page_info = H.paginate(conversations, offset, max_results)
    latest_records = []
    for conv in page:
        latest = conv.pop("_latest_record", None)
        if latest:
            latest_records.append(latest)
    _cache_listing(latest_records)
    for i, conv in enumerate(page, 1):
        conv["email_number"] = i

    return make_success(
        action="conversation_insights",
        folder=folder_disp,
        days=days,
        follow_up_days=follow_up_days,
        scanned=len(records),
        capped=primary_capped or sent_capped,
        query={
            "keyword": keyword, "sender": sender, "subject": subject,
            "recipient": recipient, "unread_only": unread_only,
            "has_attachments": has_attachments, "category": category,
            "exact_phrase": exact_phrase, "exclude": exclude_terms,
            "include_sent": include_sent,
        },
        mailbox_insights=summary,
        page_info=page_info,
        count=len(page),
        conversations=page,
        analysis_instructions=(
            "Use the conversation objects as the primary work queue. Prioritise "
            "reply_owed, waiting_on_them, unread, deadline, approval and money "
            "signals before low-context single-message mail."
        ),
        next_safe_action="Open a conversation with read_thread(entry_id=...) before replying.",
    )


@mcp.tool()
@email_tool
def generate_morning_briefing(days_to_scan: int = 3, follow_up_days: int = 2) -> str:
    """[READ-ONLY] Aggregate calendar, tasks and recent conversation threads.

    Returns a structured JSON payload for the agent to turn into a briefing.
    Includes today's calendar, due Outlook tasks, and active email threads with
    follow-up hints.

    Args:
        days_to_scan: Days back to scan email threads (1–14, default 3).
        follow_up_days: Days of silence before flagging "awaiting reply".
    """
    if not isinstance(days_to_scan, int) or not 1 <= days_to_scan <= 14:
        raise OutlookError(ErrorCode.INVALID_PARAMETER,
                           "'days_to_scan' must be between 1 and 14.")
    if not isinstance(follow_up_days, int) or follow_up_days < 1:
        raise OutlookError(ErrorCode.INVALID_PARAMETER,
                           "'follow_up_days' must be a positive integer.")
    namespace = _namespace()
    my_email = _get_my_email(namespace)
    manager = _get_manager_name(namespace)
    appointments = _todays_appointments(namespace)
    tasks = _todays_tasks(namespace)

    inbox = namespace.GetDefaultFolder(H.OL_FOLDER_INBOX)
    sent = namespace.GetDefaultFolder(H.OL_FOLDER_SENT)
    start = datetime.datetime.now() - datetime.timedelta(days=days_to_scan)
    start_str = start.strftime("%m/%d/%Y %H:%M %p")
    try:
        inbox_items = list(inbox.Items.Restrict(f"[ReceivedTime] >= '{start_str}'"))
        sent_items = list(sent.Items.Restrict(f"[SentOn] >= '{start_str}'"))
    except Exception:
        inbox_items, sent_items = [], []

    conversations: Dict[Any, List[Any]] = {}
    for item in inbox_items + sent_items:
        try:
            conversations.setdefault(item.ConversationID, []).append(item)
        except Exception:
            continue

    def item_dt(item):
        dt = H._safe_getattr(item, "ReceivedTime") or H._safe_getattr(item, "SentOn")
        return dt.replace(tzinfo=None) if dt else datetime.datetime.min

    threads = []
    for thread in conversations.values():
        if not thread:
            continue
        thread.sort(key=item_dt)
        last = H.format_email_item(thread[-1], include_body=True)
        is_from_me = bool(my_email and my_email in (last.get("sender_email") or "").lower())
        age_days = (datetime.datetime.now() - item_dt(thread[-1])).days
        status = {
            "subject": last.get("subject"),
            "last_from": "me" if is_from_me else last.get("sender"),
            "last_timestamp": item_dt(thread[-1]).strftime("%Y-%m-%d %H:%M"),
            "unread": bool(last.get("unread") and not is_from_me),
            "from_manager": bool(manager and manager.lower() in (last.get("sender") or "").lower()),
            "contains_question": "?" in (last.get("body") or ""),
            "days_since_last": age_days,
        }
        if is_from_me and age_days >= follow_up_days:
            status["follow_up_suggestion"] = f"Awaiting reply for {age_days} days."
        threads.append(status)

    threads.sort(key=lambda x: (not x["unread"], x["last_timestamp"]), reverse=True)
    return make_success(
        action="generate_morning_briefing",
        briefing_metadata={"date": datetime.date.today().strftime("%A, %B %d, %Y"),
                           "user_email": my_email, "manager_name": manager or "Not found"},
        todays_calendar=appointments,
        todays_tasks=tasks,
        conversation_threads=threads,
        analysis_instructions="Synthesise calendar, tasks and threads into a prioritised briefing.",
    )


@mcp.tool()
@email_tool
def inbox_load_estimator(days_to_scan: int = 30) -> str:
    """[READ-ONLY] Compute inbox load metrics for the agent to interpret.

    Returns counts/averages (unread-urgent, flagged, avg response delay, active
    conversations) without judging them — you assess 'calm/busy/overloaded'.

    Args:
        days_to_scan: Days back to analyse (1–60, default 30).
    """
    if not isinstance(days_to_scan, int) or not 1 <= days_to_scan <= H.ACTIONABLE_EMAIL_MAX_DAYS:
        raise OutlookError(ErrorCode.INVALID_PARAMETER,
                           f"'days_to_scan' must be between 1 and {H.ACTIONABLE_EMAIL_MAX_DAYS}.")
    namespace = _namespace()
    my_email = _get_my_email(namespace)
    inbox = namespace.GetDefaultFolder(H.OL_FOLDER_INBOX)
    sent = namespace.GetDefaultFolder(H.OL_FOLDER_SENT)
    start = datetime.datetime.now() - datetime.timedelta(days=days_to_scan)
    start_str = start.strftime("%m/%d/%Y %H:%M %p")
    try:
        items = list(inbox.Items.Restrict(f"[ReceivedTime] >= '{start_str}'")) + \
                list(sent.Items.Restrict(f"[SentOn] >= '{start_str}'"))
    except Exception:
        items = []

    urgent = ["urgent", "action required", "asap", "deadline", "critical",
              "dringend", "aksie vereis", "sperdatum", "krities", "belangrik", "spoedig"]
    conversations: Dict[Any, List[Any]] = {}
    for item in items:
        try:
            conversations.setdefault(item.ConversationID, []).append(item)
        except Exception:
            continue

    def item_dt(item):
        dt = H._safe_getattr(item, "ReceivedTime") or H._safe_getattr(item, "SentOn")
        return dt.replace(tzinfo=None) if dt else datetime.datetime.min

    unread_urgent = flagged = replied = 0
    total_delay = 0.0
    for thread in conversations.values():
        thread.sort(key=item_dt)
        last = thread[-1]
        last_addr = (H._safe_getattr(last, "SenderEmailAddress", "") or "").lower()
        if H._safe_getattr(last, "UnRead", False) and last_addr != my_email:
            if any(kw in (H._safe_getattr(last, "Subject", "") or "").lower() for kw in urgent):
                unread_urgent += 1
        if any(H._safe_getattr(i, "FlagStatus", 0) == 2 for i in thread):
            flagged += 1
        for i in range(len(thread) - 1):
            a, b = thread[i], thread[i + 1]
            if (H._safe_getattr(a, "SenderEmailAddress", "") or "").lower() != my_email and \
               (H._safe_getattr(b, "SenderEmailAddress", "") or "").lower() == my_email:
                delay = (item_dt(b) - item_dt(a)).total_seconds()
                if delay > 0:
                    total_delay += delay
                    replied += 1
    avg_delay_hours = round(total_delay / replied / 3600, 2) if replied else 0
    return make_success(
        action="inbox_load_estimator",
        scan_period_days=days_to_scan,
        inbox_metrics={"unread_urgent_count": unread_urgent,
                       "flagged_threads_count": flagged,
                       "average_response_delay_hours": avg_delay_hours,
                       "total_active_conversations": len(conversations)},
        ai_instructions="Interpret these metrics qualitatively and recommend next actions.",
    )


# ===========================================================================
# TASKS  (Outlook To-Do)
# ===========================================================================

@mcp.tool()
@email_tool
def create_outlook_task(subject: str, due_date_str: str,
                        reminder_time_str: Optional[str] = None) -> str:
    """[CREATES TASK] Create an Outlook To-Do task.

    Args:
        subject: Task title.
        due_date_str: Due date ("tomorrow", "next Friday", "2026-07-15").
        reminder_time_str: Optional reminder time (e.g. "9:00 AM").
    """
    if not subject or not subject.strip():
        raise OutlookError(ErrorCode.INVALID_PARAMETER, "subject is required.")
    outlook, _ = _connect()
    task = outlook.CreateItem(3)
    task.Subject = subject
    task.DueDate = due_date_str
    if reminder_time_str:
        task.ReminderSet = True
        task.ReminderTime = f"{due_date_str} {reminder_time_str}"
    task.Save()
    return make_success(action="create_outlook_task", subject=subject,
                        due_date=task.DueDate.strftime("%Y-%m-%d"),
                        reminder_set=bool(task.ReminderSet))


@mcp.tool()
@email_tool
def get_outlook_tasks(due: str = "today") -> str:
    """[READ-ONLY] List incomplete Outlook tasks by due window.

    Args:
        due: 'today', 'tomorrow', 'this week', or 'all' (default 'today').
    """
    due = (due or "today").lower().strip()
    if due not in {"today", "tomorrow", "this week", "all"}:
        raise OutlookError(ErrorCode.INVALID_PARAMETER,
                           "due must be 'today', 'tomorrow', 'this week' or 'all'.")
    namespace = _namespace()
    items = namespace.GetDefaultFolder(H.OL_FOLDER_TASKS).Items
    try:
        items.Sort("[DueDate]")
        items.IncludeRecurrences = True
    except Exception:
        pass
    today = datetime.date.today()
    today_str = today.strftime("%m/%d/%Y")
    if due == "today":
        restriction = f"[Complete] = false AND [DueDate] <= '{today_str}'"
    elif due == "tomorrow":
        t = (today + datetime.timedelta(days=1)).strftime("%m/%d/%Y")
        restriction = f"[Complete] = false AND [DueDate] = '{t}'"
    elif due == "this week":
        start = (today - datetime.timedelta(days=today.weekday())).strftime("%m/%d/%Y")
        end = (today - datetime.timedelta(days=today.weekday()) + datetime.timedelta(days=6)).strftime("%m/%d/%Y")
        restriction = f"[Complete] = false AND [DueDate] >= '{start}' AND [DueDate] <= '{end}'"
    else:
        restriction = "[Complete] = false"
    tasks = []
    for item in items.Restrict(restriction):
        if H._safe_getattr(item, "Class", None) == 48:
            tasks.append({"subject": item.Subject,
                          "due_date": item.DueDate.strftime("%Y-%m-%d") if H._safe_getattr(item, "DueDate") else "No due date",
                          "reminder_set": bool(H._safe_getattr(item, "ReminderSet", False))})
    return make_success(action="get_outlook_tasks", due=due, count=len(tasks), tasks=tasks)


@mcp.tool()
@email_tool
def mark_task_complete(task_subject: str) -> str:
    """[UPDATES TASK] Mark the first matching incomplete task complete."""
    namespace = _namespace()
    folder = namespace.GetDefaultFolder(H.OL_FOLDER_TASKS)
    safe_subject = task_subject.replace("'", "''")
    tasks = folder.Items.Restrict(f"[Subject] = '{safe_subject}' AND [Complete] = False")
    if tasks.Count == 0:
        raise OutlookError(ErrorCode.EMAIL_NOT_FOUND,
                           f"No active task with subject '{task_subject}'.")
    if tasks.Count > 1:
        log.info("Multiple tasks named '%s'; completing the first.", task_subject)
    tasks.Item(1).MarkComplete()
    return make_success(action="mark_task_complete", subject=task_subject,
                        status="complete")


# ---------------------------------------------------------------------------
# Internal helpers for the aggregator/task tools
# ---------------------------------------------------------------------------

def _get_my_email(namespace) -> Optional[str]:
    try:
        cu = namespace.CurrentUser
        if cu:
            ex = cu.AddressEntry.GetExchangeUser()
            if ex and ex.PrimarySmtpAddress:
                return ex.PrimarySmtpAddress.lower()
        for account in namespace.Accounts:
            if H._safe_getattr(account, "SmtpAddress"):
                return account.SmtpAddress.lower()
    except Exception as exc:
        log.debug("Could not determine user email: %s", exc)
    return None


def _get_manager_name(namespace) -> Optional[str]:
    try:
        cu = namespace.CurrentUser
        if cu:
            ex = cu.GetExchangeUser()
            if ex:
                mgr = ex.GetManager()
                if mgr:
                    return mgr.Name
    except Exception as exc:
        log.debug("Could not determine manager: %s", exc)
    return None


def _todays_appointments(namespace) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    try:
        calendar = namespace.GetDefaultFolder(H.OL_FOLDER_CALENDAR)
        items = calendar.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")
        start = datetime.datetime.now().replace(hour=0, minute=0, second=0)
        end = start + datetime.timedelta(days=1)
        restriction = (f"[Start] < '{end.strftime('%m/%d/%Y %H:%M %p')}' AND "
                       f"[End] > '{start.strftime('%m/%d/%Y %H:%M %p')}'")
        for item in items.Restrict(restriction):
            out.append({"subject": item.Subject,
                        "start": item.Start.strftime("%I:%M %p").lstrip("0"),
                        "end": item.End.strftime("%I:%M %p").lstrip("0"),
                        "location": item.Location or "Not specified",
                        "all_day": bool(item.AllDayEvent)})
    except Exception as exc:
        log.debug("Could not read appointments: %s", exc)
    return out


def _todays_tasks(namespace) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    try:
        items = namespace.GetDefaultFolder(H.OL_FOLDER_TASKS).Items
        items.Sort("[DueDate]")
        items.IncludeRecurrences = True
        today_str = datetime.date.today().strftime("%m/%d/%Y")
        for item in items.Restrict(f"[Complete] = false AND [DueDate] <= '{today_str}'"):
            if H._safe_getattr(item, "Class", None) == 48:
                out.append({"subject": item.Subject,
                            "due_date": item.DueDate.strftime("%Y-%m-%d") if H._safe_getattr(item, "DueDate") else "No due date"})
    except Exception as exc:
        log.debug("Could not read tasks: %s", exc)
    return out


# ===========================================================================
# Startup / health check
# ===========================================================================

def _startup_health_check() -> None:
    """Log a startup banner + connectivity status to stderr (never stdout)."""
    log.info("Starting Outlook MCP Server (debug=%s).", H.debug_enabled())
    if not _WIN32_AVAILABLE:
        log.warning("pywin32/Outlook COM not available — tools will return "
                    "OUTLOOK_NOT_AVAILABLE until run on Windows with Outlook.")
        return
    try:
        _, namespace = _connect()
        inbox = namespace.GetDefaultFolder(H.OL_FOLDER_INBOX)
        log.info("Connected to Outlook. Inbox contains %s items.",
                 H._safe_getattr(inbox.Items, "Count", "?"))
    except OutlookError as exc:
        log.warning("Outlook connectivity check failed: %s", exc.message)
    except Exception as exc:  # never let a flaky COM probe block startup
        log.warning("Outlook connectivity check error: %s", exc)


if __name__ == "__main__":
    _startup_health_check()
    log.info("MCP server ready. Press Ctrl+C to stop.")
    mcp.run()
