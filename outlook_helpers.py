"""Pure, side-effect-free helpers for the Outlook MCP server.

Everything in this module is deliberately free of any hard dependency on
``win32com``/Outlook so it can be unit-tested on any platform without a live
mailbox.  The server module (``outlook_mcp_server.py``) imports these helpers
and supplies the COM-specific glue.

Design goals:

* **Structured responses** – every tool returns a predictable JSON envelope
  with ``success`` plus either result fields or a machine-readable
  ``error_code``.  An LLM agent can branch on these without parsing prose.
* **No secret/PII leakage** – logging is routed to stderr and email bodies are
  never logged unless debug mode is explicitly enabled.
* **Safe-by-construction** – recipient validation, search predicates and body
  cleaning live here as small, well-tested functions.
"""

from __future__ import annotations

import json
import logging
import os
import re
import sys
from typing import Any, Dict, Iterable, List, Optional, Tuple

# ---------------------------------------------------------------------------
# Constants / limits
# ---------------------------------------------------------------------------

MAX_DAYS = 180                      # absolute ceiling for look-back windows
DEFAULT_DAYS = 7
ACTIONABLE_EMAIL_MAX_DAYS = 60
DEFAULT_PAGE_SIZE = 25
MAX_PAGE_SIZE = 100
SNIPPET_LENGTH = 300

# Bulk-operation safety ceiling.  Tools that can touch many messages must never
# exceed this in a single call without an explicit, higher limit from the user.
MAX_BULK_OPERATION = 50

# Outlook MAPI default-folder indexes (olDefaultFolders enumeration).
OL_FOLDER_DELETED_ITEMS = 3
OL_FOLDER_SENT = 5
OL_FOLDER_INBOX = 6
OL_FOLDER_DRAFTS = 16
OL_FOLDER_CALENDAR = 9
OL_FOLDER_TASKS = 13

# Outlook importance / flag enumerations, mapped to friendly labels.
IMPORTANCE_LABELS = {0: "Low", 1: "Normal", 2: "High"}
FLAG_LABELS = {
    0: "none",
    1: "complete",
    2: "flagged",
}

# Attachment handling.
SAFE_TEXT_EXTENSIONS = {
    ".txt", ".csv", ".tsv", ".md", ".markdown", ".log", ".json", ".xml",
    ".yaml", ".yml", ".ini", ".html", ".htm",
}
DEFAULT_MAX_ATTACHMENT_MB = 25


# ---------------------------------------------------------------------------
# Error codes – stable identifiers an agent can branch on.
# ---------------------------------------------------------------------------

class ErrorCode:
    OUTLOOK_NOT_AVAILABLE = "OUTLOOK_NOT_AVAILABLE"
    OUTLOOK_CONNECTION_FAILED = "OUTLOOK_CONNECTION_FAILED"
    INVALID_PARAMETER = "INVALID_PARAMETER"
    INVALID_SEARCH_QUERY = "INVALID_SEARCH_QUERY"
    INVALID_RECIPIENT = "INVALID_RECIPIENT"
    EMAIL_NOT_FOUND = "EMAIL_NOT_FOUND"
    FOLDER_NOT_FOUND = "FOLDER_NOT_FOUND"
    DRAFT_NOT_FOUND = "DRAFT_NOT_FOUND"
    THREAD_NOT_FOUND = "THREAD_NOT_FOUND"
    ATTACHMENT_NOT_FOUND = "ATTACHMENT_NOT_FOUND"
    ATTACHMENT_TOO_LARGE = "ATTACHMENT_TOO_LARGE"
    UNSUPPORTED_ATTACHMENT = "UNSUPPORTED_ATTACHMENT"
    NO_LISTING_CONTEXT = "NO_LISTING_CONTEXT"
    CONFIRMATION_REQUIRED = "CONFIRMATION_REQUIRED"
    BULK_LIMIT_EXCEEDED = "BULK_LIMIT_EXCEEDED"
    ACTION_FAILED = "ACTION_FAILED"
    PERMISSION_DENIED = "PERMISSION_DENIED"


# ---------------------------------------------------------------------------
# Logging / redaction
# ---------------------------------------------------------------------------

_LOGGER_NAME = "outlook_mcp"


def debug_enabled() -> bool:
    """True when verbose/debug logging is explicitly opted into via env var."""
    return os.environ.get("OUTLOOK_MCP_DEBUG", "").strip().lower() in {
        "1", "true", "yes", "on",
    }


def get_logger() -> logging.Logger:
    """Return the shared logger configured to write to **stderr**.

    Writing logs to stdout would corrupt the MCP stdio JSON-RPC stream, so all
    logging must go to stderr.  Bodies/PII are never logged at INFO level.
    """
    logger = logging.getLogger(_LOGGER_NAME)
    if not logger.handlers:
        handler = logging.StreamHandler(stream=sys.stderr)
        handler.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)s [outlook-mcp] %(message)s")
        )
        logger.addHandler(handler)
        logger.propagate = False
    logger.setLevel(logging.DEBUG if debug_enabled() else logging.INFO)
    return logger


def redact(text: Optional[str]) -> str:
    """Return a non-sensitive placeholder for free-text content.

    Used so that body/reply text never lands in logs unless debug mode is on,
    in which case a short, length-capped preview is allowed.
    """
    if text is None:
        return "<none>"
    if debug_enabled():
        preview = " ".join(str(text).split())
        return preview[:120] + ("…" if len(preview) > 120 else "")
    return f"<{len(str(text))} chars redacted>"


# ---------------------------------------------------------------------------
# Response envelopes
# ---------------------------------------------------------------------------

def make_success(action: Optional[str] = None, **fields: Any) -> Dict[str, Any]:
    """Build a success envelope. ``action`` names the operation performed."""
    payload: Dict[str, Any] = {"success": True}
    if action is not None:
        payload["action"] = action
    payload.update(fields)
    return payload


def make_error(
    code: str,
    message: str,
    details: Optional[str] = None,
    retryable: bool = False,
    **extra: Any,
) -> Dict[str, Any]:
    """Build a structured error envelope.

    Args:
        code: One of :class:`ErrorCode` – stable and machine-branchable.
        message: Short, human/agent friendly description (no raw stack traces).
        details: Optional actionable hint (e.g. "Use YYYY-MM-DD").
        retryable: Whether retrying the same call could succeed (rate limits,
            transient network), as opposed to a caller error.
    """
    payload: Dict[str, Any] = {
        "success": False,
        "error_code": code,
        "message": message,
        "retryable": retryable,
    }
    if details:
        payload["details"] = details
    payload.update(extra)
    return payload


def to_json(payload: Any) -> str:
    """Serialize a payload to a stable, pretty JSON string for tool output."""
    return json.dumps(payload, indent=2, default=str, ensure_ascii=False)


# ---------------------------------------------------------------------------
# Recipient validation
# ---------------------------------------------------------------------------

# Pragmatic address check: exactly one '@', non-empty local part, a dotted
# domain, no whitespace.  Intentionally not RFC-5322-complete – it is meant to
# catch obvious mistakes before they reach Outlook, not to be a parser.
_EMAIL_RE = re.compile(r"^[^@\s,;<>]+@[^@\s,;<>]+\.[^@\s,;<>]+$")


def is_valid_email(address: str) -> bool:
    """Return True if ``address`` looks like a deliverable SMTP address."""
    if not address:
        return False
    return bool(_EMAIL_RE.match(address.strip()))


def parse_recipients(value: Optional[str]) -> List[str]:
    """Split a recipient string on commas/semicolons into trimmed addresses."""
    if not value:
        return []
    parts = re.split(r"[;,]", value)
    return [p.strip() for p in parts if p.strip()]


def validate_recipients(value: Optional[str]) -> Tuple[List[str], List[str]]:
    """Return ``(valid, invalid)`` recipient lists parsed from ``value``."""
    valid: List[str] = []
    invalid: List[str] = []
    for addr in parse_recipients(value):
        (valid if is_valid_email(addr) else invalid).append(addr)
    return valid, invalid


# ---------------------------------------------------------------------------
# Body cleaning / quoted-reply trimming / snippets
# ---------------------------------------------------------------------------

# Markers that typically begin the quoted/original portion of a reply.
_QUOTE_MARKERS = (
    re.compile(r"^-{2,}\s*Original Message\s*-{2,}\s*$", re.IGNORECASE),
    re.compile(r"^_{5,}\s*$"),
    re.compile(r"^From:\s.+", re.IGNORECASE),
    re.compile(r"^On .+ wrote:\s*$", re.IGNORECASE),
    re.compile(r"^Sent from my \w+", re.IGNORECASE),
)


def normalize_whitespace(text: Optional[str]) -> str:
    """Collapse runs of whitespace to single spaces and strip ends."""
    if not text:
        return ""
    return " ".join(str(text).split())


def make_snippet(text: Optional[str], length: int = SNIPPET_LENGTH) -> str:
    """Return a single-line snippet of at most ``length`` characters."""
    cleaned = normalize_whitespace(text)
    if len(cleaned) <= length:
        return cleaned
    return cleaned[:length].rstrip() + "…"


def trim_quoted_reply(body: Optional[str]) -> Tuple[str, bool]:
    """Trim the quoted/original portion from a reply body.

    Returns ``(trimmed_text, was_trimmed)``. Conservative: only trims at the
    first recognised quote marker that appears after at least one line of real
    content, so a forwarded message that *starts* with ``From:`` is not gutted.
    """
    if not body:
        return "", False
    lines = body.splitlines()
    seen_content = False
    for idx, line in enumerate(lines):
        stripped = line.strip()
        if any(marker.match(stripped) for marker in _QUOTE_MARKERS):
            if seen_content and idx > 0:
                trimmed = "\n".join(lines[:idx]).rstrip()
                return trimmed, True
        elif stripped:
            seen_content = True
    return body.rstrip(), False


# ---------------------------------------------------------------------------
# Categories (Outlook's equivalent of labels)
# ---------------------------------------------------------------------------

def parse_categories(value: Optional[str]) -> List[str]:
    """Outlook stores categories as a ';'-joined string. Parse to a list."""
    if not value:
        return []
    return [c.strip() for c in value.split(";") if c.strip()]


def join_categories(categories: Iterable[str]) -> str:
    """Join a list of categories back into Outlook's storage format."""
    seen: List[str] = []
    for c in categories:
        c = c.strip()
        if c and c not in seen:
            seen.append(c)
    return "; ".join(seen)


# ---------------------------------------------------------------------------
# Search predicate (applied in Python for reliability)
# ---------------------------------------------------------------------------

def email_matches(
    email: Dict[str, Any],
    *,
    keyword: Optional[str] = None,
    sender: Optional[str] = None,
    subject: Optional[str] = None,
    recipient: Optional[str] = None,
    unread_only: bool = False,
    has_attachments: Optional[bool] = None,
    category: Optional[str] = None,
    exact_phrase: Optional[str] = None,
    exclude: Optional[Iterable[str]] = None,
) -> bool:
    """Return True if a *formatted* email dict satisfies all given filters.

    Operating on the already-normalised dict (rather than building a fragile
    DASL/SQL string) makes search reliable across Outlook versions and trivially
    unit-testable. ``keyword`` matches subject/sender/body; other params scope
    to a single field. All provided filters are AND-combined.
    """
    subject_l = (email.get("subject") or "").lower()
    sender_l = " ".join(
        str(email.get(k) or "") for k in ("sender", "sender_email")
    ).lower()
    body_l = (email.get("body") or "").lower()
    recipients_l = " ".join(email.get("recipients") or []).lower()
    categories_l = " ".join(parse_categories(email.get("categories"))).lower()

    if keyword:
        # support simple "a OR b" semantics (case-insensitive operator)
        terms = [
            t.strip().lower()
            for t in re.split(r"\s+OR\s+", keyword, flags=re.IGNORECASE)
            if t.strip()
        ]
        if terms and not any(
            t in subject_l or t in sender_l or t in body_l for t in terms
        ):
            return False

    if sender and sender.lower() not in sender_l:
        return False
    if subject and subject.lower() not in subject_l:
        return False
    if recipient and recipient.lower() not in recipients_l:
        return False
    if exact_phrase and exact_phrase.lower() not in (
        subject_l + "\n" + body_l
    ):
        return False
    if unread_only and not email.get("unread"):
        return False
    if has_attachments is not None and bool(email.get("has_attachments")) != has_attachments:
        return False
    if category and category.lower() not in categories_l:
        return False
    if exclude:
        haystack = subject_l + "\n" + sender_l + "\n" + body_l
        for term in exclude:
            if term and term.lower() in haystack:
                return False
    return True


def paginate(items: List[Any], offset: int, limit: int) -> Tuple[List[Any], Dict[str, Any]]:
    """Return ``(page, page_info)`` for the given slice of ``items``."""
    offset = max(0, int(offset))
    limit = max(1, min(int(limit), MAX_PAGE_SIZE))
    total = len(items)
    page = items[offset:offset + limit]
    return page, {
        "total_matched": total,
        "offset": offset,
        "limit": limit,
        "returned": len(page),
        "has_more": offset + limit < total,
        "next_offset": offset + limit if offset + limit < total else None,
    }


# ---------------------------------------------------------------------------
# Attachment helpers
# ---------------------------------------------------------------------------

def sanitize_filename(name: str) -> str:
    """Strip path separators and dangerous characters from an attachment name."""
    name = os.path.basename(name or "").strip()
    name = re.sub(r"[^A-Za-z0-9._ +\-()\[\]]", "_", name)
    name = name.lstrip(".") or "attachment"
    return name[:200]


def is_safe_text_attachment(filename: str) -> bool:
    """Whether a filename has an extension safe to read inline as text."""
    _, ext = os.path.splitext((filename or "").lower())
    return ext in SAFE_TEXT_EXTENSIONS


def max_attachment_bytes() -> int:
    """Configured maximum attachment size to read/save, in bytes."""
    try:
        mb = float(os.environ.get("OUTLOOK_MAX_ATTACHMENT_MB", DEFAULT_MAX_ATTACHMENT_MB))
    except (TypeError, ValueError):
        mb = DEFAULT_MAX_ATTACHMENT_MB
    return int(mb * 1024 * 1024)


# ---------------------------------------------------------------------------
# Email formatting (duck-typed against a COM MailItem or a test fake)
# ---------------------------------------------------------------------------

def _safe_getattr(obj: Any, name: str, default: Any = None) -> Any:
    try:
        return getattr(obj, name, default)
    except Exception:  # pragma: no cover - COM can raise on attribute access
        return default


def _format_dt(value: Any, fmt: str = "%Y-%m-%d %H:%M:%S") -> Optional[str]:
    if value is None:
        return None
    try:
        # COM datetimes are tz-aware; drop tz for stable, comparable strings.
        return value.replace(tzinfo=None).strftime(fmt)
    except Exception:
        try:
            return str(value)
        except Exception:
            return None


def extract_recipients(mail: Any) -> List[str]:
    """Extract ``Name <addr>`` strings from a COM-style Recipients collection."""
    recipients: List[str] = []
    coll = _safe_getattr(mail, "Recipients")
    if not coll:
        return recipients
    try:
        count = int(_safe_getattr(coll, "Count", 0) or 0)
    except Exception:
        return recipients
    for i in range(1, count + 1):
        try:
            r = coll(i)
            name = (_safe_getattr(r, "Name", "") or "").strip()
            addr = (_safe_getattr(r, "Address", "") or "").strip()
            if name and addr and name.lower() != addr.lower():
                recipients.append(f"{name} <{addr}>")
            elif addr:
                recipients.append(addr)
            elif name:
                recipients.append(name)
        except Exception:
            continue
    return recipients


def format_email_item(
    mail: Any,
    *,
    include_body: bool = True,
    trim_quotes: bool = False,
) -> Dict[str, Any]:
    """Convert a COM MailItem (or test fake) into a normalised dict.

    Never raises on a single bad attribute; missing fields degrade to ``None``
    and a ``warnings`` list records partial extraction.
    """
    warnings: List[str] = []

    body = ""
    if include_body:
        body = _safe_getattr(mail, "Body", "") or ""
        if not body:
            # Fall back to a stripped HTML body if plain text is empty.
            html = _safe_getattr(mail, "HTMLBody", "") or ""
            if html:
                body = strip_html(html)
                warnings.append("Plain-text body empty; derived from HTML.")
    was_trimmed = False
    if include_body and trim_quotes and body:
        body, was_trimmed = trim_quoted_reply(body)

    attachments_count = 0
    has_attachments = False
    att = _safe_getattr(mail, "Attachments")
    if att is not None:
        try:
            attachments_count = int(_safe_getattr(att, "Count", 0) or 0)
            has_attachments = attachments_count > 0
        except Exception:
            warnings.append("Could not read attachment count.")

    importance = _safe_getattr(mail, "Importance", 1)
    flag_status = _safe_getattr(mail, "FlagStatus", 0)

    data: Dict[str, Any] = {
        "id": _safe_getattr(mail, "EntryID"),
        "conversation_id": _safe_getattr(mail, "ConversationID"),
        "subject": _safe_getattr(mail, "Subject"),
        "sender": _safe_getattr(mail, "SenderName"),
        "sender_email": _safe_getattr(mail, "SenderEmailAddress"),
        "received_time": _format_dt(_safe_getattr(mail, "ReceivedTime")),
        "sent_time": _format_dt(_safe_getattr(mail, "SentOn")),
        "recipients": extract_recipients(mail),
        "has_attachments": has_attachments,
        "attachment_count": attachments_count,
        "unread": bool(_safe_getattr(mail, "UnRead", False)),
        "importance": IMPORTANCE_LABELS.get(importance, "Normal"),
        "flagged": FLAG_LABELS.get(flag_status, "none"),
        "categories": _safe_getattr(mail, "Categories", "") or "",
    }
    if include_body:
        data["body"] = body
        data["snippet"] = make_snippet(body)
        if was_trimmed:
            warnings.append("Quoted reply text was trimmed from the body.")
    else:
        data["snippet"] = make_snippet(_safe_getattr(mail, "Body", "") or "")
    if warnings:
        data["warnings"] = warnings
    return data


_TAG_RE = re.compile(r"<[^>]+>")
_STYLE_RE = re.compile(r"<(script|style)[^>]*>.*?</\1>", re.IGNORECASE | re.DOTALL)


def strip_html(html: str) -> str:
    """Very small HTML→text fallback (no external deps).

    Removes script/style blocks and tags, decodes a handful of common
    entities, and collapses blank lines. Good enough as a *fallback* when
    Outlook's plain-text body is unexpectedly empty.
    """
    if not html:
        return ""
    text = _STYLE_RE.sub(" ", html)
    text = re.sub(r"<\s*br\s*/?\s*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</\s*p\s*>", "\n\n", text, flags=re.IGNORECASE)
    text = _TAG_RE.sub("", text)
    for entity, repl in (
        ("&nbsp;", " "), ("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
        ("&quot;", '"'), ("&#39;", "'"), ("&apos;", "'"),
    ):
        text = text.replace(entity, repl)
    lines = [ln.rstrip() for ln in text.splitlines()]
    # collapse 3+ blank lines to a single blank line
    out: List[str] = []
    blank = 0
    for ln in lines:
        if ln.strip():
            blank = 0
            out.append(ln)
        else:
            blank += 1
            if blank <= 1:
                out.append("")
    return "\n".join(out).strip()
