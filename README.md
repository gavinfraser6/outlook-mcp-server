# Outlook MCP Server

Operate a **local Microsoft Outlook desktop profile** like a careful human email
assistant — search, read, summarise, draft, reply, forward, organise, triage,
and manage tasks — with safety rails around anything that sends or deletes.

It talks to the Outlook application installed on a Windows machine via COM
automation (`pywin32`). It does **not** use Gmail/IMAP/Graph, OAuth tokens, or
stored passwords — it drives the Outlook session the signed-in Windows user has
already authenticated. See [Authentication & Security](#authentication--security).

### Three ways to use it

| Surface | Command | For |
| --- | --- | --- |
| **MCP server** | `python outlook_mcp_server.py` | AI agents ([Model Context Protocol](https://modelcontextprotocol.io)) — Codex, Claude, etc. |
| **Web dashboard** | `python outlook_web.py` → http://127.0.0.1:8765 | A human, in the browser. Triage-first, keyboard-driven. |
| **Scheduled digest** | `python outlook_schedule.py` | Windows Task Scheduler — a daily "what needs attention" digest. |

All three share one engine, one safety model, and one deterministic triage
ranker, so behaviour is identical everywhere. Built for **busy mailboxes** (this
was validated against an inbox with 100+ unread): a persistent warm Outlook
connection, server-side filtering, early-stop scans, and explainable ranking so
you can actually *get mail done*.

---

## Table of contents

- [What this server does](#what-this-server-does)
- [Safety model](#safety-model)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Web dashboard](#web-dashboard)
- [Scheduled digest](#scheduled-digest)
- [Busy-mailbox optimisations](#busy-mailbox-optimisations)
- [Environment variables](#environment-variables)
- [Tool reference](#tool-reference)
- [Response format](#response-format)
- [Recommended agent workflow](#recommended-agent-workflow)
- [Example workflows](#example-workflows)
- [Authentication & Security](#authentication--security)
- [Error handling](#error-handling)
- [Development & testing](#development--testing)
- [Troubleshooting](#troubleshooting)
- [Known limitations](#known-limitations)

---

## What this server does

- **Search** with structured filters (sender, subject, keyword, recipient, date
  range, unread-only, has-attachments, category/label, exact phrase, exclude
  terms) plus pagination.
- **Read** full emails and whole **conversation threads** in chronological order.
- **Draft** new emails, replies, and forwards — *drafts are the default*; nothing
  is sent without an explicit, confirmed send step.
- **Send** new emails or saved drafts, gated behind `confirm=true`.
- **Organise**: move, archive, move-to-trash (recoverable), mark read/unread,
  and apply/remove Outlook **categories** (the closest thing to labels).
- **Attachments**: list metadata, save to disk (size-limited, never executed),
  and read safe text attachments inline.
- **Triage**: a deterministic, explainable "what needs attention first" ranking
  (`triage_inbox`) shared by the agent, the web dashboard and the digest.
- **Productivity**: a morning-briefing aggregator, inbox-load metrics, and
  Outlook **tasks** (create / list / complete).

All responses are **structured JSON** designed for an LLM to consume.

---

## Safety model

This server is built so an agent finds it **hard to misuse**:

| Action | Behaviour |
| --- | --- |
| Read / search | Always allowed. |
| Compose / reply / forward | **Creates a draft by default.** Returns a `draft_id`. |
| Send new email | `send_email` only — requires `confirm=true` (preview first otherwise). |
| Send a draft | `send_draft` only — requires `confirm=true`. |
| Delete | `move_to_trash` moves to **Deleted Items** (recoverable). |
| Permanent delete | **Not implemented, by design.** |
| Bulk actions | Capped (`max_results` ≤ 100) and previewed; tools act on one item per call. |
| Recipients | Validated before any draft/send; invalid addresses are rejected. |
| Secrets | None handled. Logs go to **stderr** and never include bodies unless debug is on. |

The only two tools that put mail on the wire are **`send_email`** and
**`send_draft`**, and both refuse to act until called again with `confirm=true`.

---

## Prerequisites

- **Windows** with **Microsoft Outlook** (classic desktop) installed, configured,
  and signed in. (The "new" Outlook / web Outlook does not expose COM.)
- **Python 3.10+**
- An MCP-compatible client (Claude Desktop, Codex, etc.)

---

## Installation

```bash
git clone https://github.com/gavinfraser6/outlook-mcp-server.git
cd outlook-mcp-server
pip install -r requirements.txt
```

Run it directly to verify it starts (it prints a startup banner to **stderr**):

```bash
python outlook_mcp_server.py
```

You should see `Connected to Outlook. Inbox contains N items.` on stderr. Stop
with `Ctrl+C`. Normally you let your MCP client launch it (see below).

---

## Configuration

Add the server to your MCP client. Example `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["C:\\path\\to\\outlook-mcp-server\\outlook_mcp_server.py"],
      "env": {
        "OUTLOOK_MCP_DEBUG": "false",
        "OUTLOOK_ARCHIVE_FOLDER": "Archive",
        "OUTLOOK_MAX_ATTACHMENT_MB": "25"
      }
    }
  }
}
```

Use an **absolute path** to `outlook_mcp_server.py`. The `env` block is optional.

---

## Web dashboard

A localhost, triage-first UI for getting through mail quickly — no agent needed.

```bash
python outlook_web.py            # serves http://127.0.0.1:8765
# or double-click run_web.bat on Windows
```

Then open **http://127.0.0.1:8765**. It does **not** require the MCP server to be
running — it talks to Outlook directly via a single warm connection.

**What you can do**
- **Needs attention** tab: your inbox ranked by a deterministic scorer (most
  urgent first) with explainable reason chips (unread, from manager, urgent
  language, mentions invoice/deadline, contains a question, …; newsletters and
  no-reply mail are pushed down).
- One-click **Archive / Done (mark read) / Trash / Reply** on every card.
- **Open** a card to read the whole thread.
- **Search** tab: keyword search across the mailbox.
- **Drafts** tab: review and send saved drafts.
- **Compose / reply / forward**: always creates a draft first; **Send** asks for
  a confirmation before anything leaves your outbox.

**Keyboard shortcuts** (when not typing): `j`/`k` move, `o`/`Enter` open,
`r` reply, `e` archive, `u` mark done, `#` trash, `c` compose, `/` search,
`g` refresh, `Esc` close.

**Security**
- Binds to **127.0.0.1** only by default. Don't bind to `0.0.0.0` unless you also
  set a token.
- Optional shared token: set `OUTLOOK_WEB_TOKEN`. Then open the UI with
  `http://127.0.0.1:8765/?token=YOUR_TOKEN` (the page forwards it on every API
  call; `/api/health` stays open).
- Sending still obeys the confirmation guard — the UI cannot send silently.

---

## Scheduled digest

`outlook_schedule.py` produces a deterministic "what needs my attention" digest
on a schedule. It is **read-only by default** (scans, ranks, writes files) and
prints a short summary.

```bash
python outlook_schedule.py                 # write digest + print summary
python outlook_schedule.py --days 3 --top 20
python outlook_schedule.py --auto-categorize          # opt-in: tags top unread
python outlook_schedule.py --email-to you@corp.com    # opt-in: emails the digest
```

It writes `digest.json` and `digest.html` to the state dir (`OUTLOOK_STATE_DIR`,
default `%LOCALAPPDATA%\outlook-mcp\`). The web dashboard exposes the latest one
at `/api/digest`.

**Optional, explicit side-effects** (off unless you pass the flag):
- `--auto-categorize` applies a category (default `⚑ Needs Reply`, capped at
  `--max-categorize`, only to unread inbound at/above `--min-score`) so flagged
  mail stands out inside Outlook.
- `--email-to ADDRESS` emails the HTML digest to yourself — an explicit send you
  opt into by passing the flag.

**Windows Task Scheduler setup**
1. Task Scheduler → *Create Basic Task* → set your trigger (e.g. daily 07:30).
2. Action: *Start a program*.
   - Program/script: `python` (or the full path to `python.exe`)
   - Add arguments: `outlook_schedule.py --days 2 --top 15`
   - Start in: the repo folder (or just point at `run_schedule.bat`).
3. Outlook must be running/signed in when the task fires.

A non-zero exit code is returned on failure so the scheduler can report it.

---

## Busy-mailbox optimisations

Tuned for inboxes with thousands of items and hundreds unread:

- **Warm connection.** The web UI funnels all Outlook access through one
  dedicated COM (STA) worker thread that connects once and stays connected — no
  per-request `Dispatch`, no cross-thread COM races.
- **Server-side narrowing.** Folder scans push a date floor (plus unread/subject
  where safe) into Outlook's `Restrict`, so the whole folder is never walked in
  Python. The authoritative match still runs in Python, so results stay correct.
- **Early-stop + caps.** Scans stop once enough matches for the page are found;
  results report `capped: true` when a cap is hit, so totals are an honest lower
  bound instead of a slow exact count.
- **Deterministic triage.** Ranking is a transparent rule set (no LLM latency),
  shared by the agent tool, the dashboard and the digest.

---

## Environment variables

| Variable | Default | Purpose |
| --- | --- | --- |
| `OUTLOOK_MCP_DEBUG` | `false` | When truthy (`1/true/yes/on`), enables DEBUG logging **and** allows short body previews in logs. Leave off in production. |
| `OUTLOOK_ARCHIVE_FOLDER` | `Archive` | Folder name `archive_email` moves messages to. |
| `OUTLOOK_ATTACHMENT_DIR` | system temp dir | Where `save_attachment` writes files. |
| `OUTLOOK_MAX_ATTACHMENT_MB` | `25` | Max size (MB) for reading/saving an attachment. |
| `OUTLOOK_WEB_HOST` | `127.0.0.1` | Web dashboard bind host. Keep on localhost. |
| `OUTLOOK_WEB_PORT` | `8765` | Web dashboard port. |
| `OUTLOOK_WEB_TOKEN` | _(unset)_ | If set, the web API requires this token (header `X-Outlook-Token` or `?token=`). |
| `OUTLOOK_STATE_DIR` | `%LOCALAPPDATA%\outlook-mcp` | Where the scheduled digest is written. |

---

## Tool reference

Tags: **[READ-ONLY]** safe · **[DRAFT]** writes a draft, never sends ·
**[SEND]** sends mail (needs `confirm=true`) · **[ORGANISE]** moves/labels ·
**[DESTRUCTIVE]** recoverable delete.

Action tools accept **either** `email_number` (from the most recent listing) or a
stable **`entry_id`** (returned on every result, survives across listings).

### Discovery & reading
| Tool | Tag | Summary |
| --- | --- | --- |
| `list_folders` | READ-ONLY | List mail folders / sub-folders. |
| `search_emails` | READ-ONLY | Structured search (see filters below). |
| `list_recent_emails` | READ-ONLY | Recent emails from a folder, newest first. |
| `get_unread_emails` | READ-ONLY | Unread emails only. |
| `count_unread_emails` | READ-ONLY | Unread count for a folder. |
| `get_email_by_number` | READ-ONLY | Full email: body, recipients, attachments, labels. |
| `read_thread` | READ-ONLY | Whole conversation, chronological, with participants. |
| `list_attachments` | READ-ONLY | Attachment metadata (filename, type, size, index). |
| `read_attachment` | READ-ONLY | Read a **text** attachment inline (size-limited). |
| `save_attachment` | writes file | Save an attachment to disk (never executed). |

`search_emails` filters: `keyword` (matches subject/sender/body, supports
`"a OR b"`), `sender`, `subject`, `recipient`, `days` (1–180), `unread_only`,
`has_attachments`, `category`, `exact_phrase`, `exclude` (comma-separated),
`folder_name`, `max_results` (1–100), `offset`.

### Drafting (safe)
| Tool | Tag | Summary |
| --- | --- | --- |
| `create_draft` | DRAFT | Compose a new draft (validates recipients). |
| `update_draft` | DRAFT | Edit a draft's to/cc/bcc/subject/body. |
| `list_drafts` | READ-ONLY | List saved drafts. |
| `delete_draft` | DESTRUCTIVE | Delete a draft (drafts only). |
| `reply_to_email_by_number` | DRAFT* | Reply (keeps quoted thread). Draft unless `send=true`. |
| `forward_email` | DRAFT* | Forward (keeps attachments). Draft unless `send=true`. |
| `compose_email` | LEGACY | Drafts by default; sends only if `send=true`. |

### Sending (explicit + confirmed)
| Tool | Tag | Summary |
| --- | --- | --- |
| `send_email` | SEND | Compose **and send** a new email. Needs `confirm=true`. |
| `send_draft` | SEND | Send an existing draft by `draft_id`. Needs `confirm=true`. |

### Organising
| Tool | Tag | Summary |
| --- | --- | --- |
| `move_email_by_number` | ORGANISE | Move an email to a named folder. |
| `archive_email` | ORGANISE | Move to the Archive folder. |
| `move_to_trash` | DESTRUCTIVE | Move to Deleted Items (recoverable). |
| `mark_as_read` / `mark_as_unread` | ORGANISE | Toggle read state. |
| `list_categories` | READ-ONLY | List Outlook categories (labels). |
| `apply_category` / `remove_category` | ORGANISE | Add/remove a category on an email. |

### Productivity & tasks
| Tool | Tag | Summary |
| --- | --- | --- |
| `triage_inbox` | READ-ONLY | **Ranked** "what needs attention first" with scores + reasons (deterministic). |
| `prioritize_inbox` | READ-ONLY | Raw recent inbox data for the agent to rank itself. |
| `generate_morning_briefing` | READ-ONLY | Calendar + tasks + active threads. |
| `inbox_load_estimator` | READ-ONLY | Inbox-load metrics to interpret. |
| `create_outlook_task` | writes task | Create an Outlook To-Do task. |
| `get_outlook_tasks` | READ-ONLY | List incomplete tasks by due window. |
| `mark_task_complete` | writes task | Complete a task by subject. |

---

## Response format

Every tool returns a JSON string. Success:

```json
{
  "success": true,
  "action": "search_emails",
  "page_info": { "total_matched": 3, "offset": 0, "limit": 25, "returned": 3, "has_more": false, "next_offset": null },
  "results": [
    {
      "email_number": 1,
      "entry_id": "00000000ABCD…",
      "thread_id": "C1A2…",
      "subject": "Invoice #42 due Friday",
      "from": "Acme Billing <billing@acme.com>",
      "to": ["me@example.com"],
      "date": "2026-06-30 09:12:00",
      "snippet": "Please pay invoice 42 by Friday.",
      "labels": ["Finance"],
      "unread": true,
      "has_attachments": true,
      "attachment_count": 1,
      "importance": "High",
      "flagged": "none"
    }
  ],
  "next_safe_action": "Use get_email_by_number(email_number) to read a result…"
}
```

Error:

```json
{
  "success": false,
  "error_code": "INVALID_RECIPIENT",
  "message": "One or more recipient addresses are invalid.",
  "details": "Invalid: not-an-email",
  "retryable": false,
  "invalid_recipients": ["not-an-email"]
}
```

Send tools return a **preview** with `error_code: "CONFIRMATION_REQUIRED"` until
called again with `confirm=true`.

---

## Recommended agent workflow

1. **Search before reading.** If the exact email is unknown, call `search_emails`
   (or `list_recent_emails`/`get_unread_emails`) first.
2. **Read the full thread before replying.** Use `read_thread` so you understand
   who said what.
3. **Draft before sending.** Use `create_draft` / `reply_to_email_by_number` /
   `forward_email` — these never send. Show the draft to the user.
4. **Confirm before sending or destructive/bulk actions.** Sending requires an
   explicit `send_email`/`send_draft` with `confirm=true`.
5. **Prefer archive over delete.** Use `archive_email`; `move_to_trash` is
   recoverable; there is no permanent delete.
6. **Use categories for organisation** via `apply_category` / `remove_category`.
7. **Report exactly what changed** — every action response includes the affected
   subject/ids and a count.

Pass IDs forward: `entry_id`/`thread_id`/`draft_id` from one call are the inputs
to the next.

---

## Example workflows

**"Find the email from Acme about the invoice and summarise the thread."**
```
search_emails(sender="acme", keyword="invoice", days=30)
read_thread(email_number=1)
```

**"Reply saying I'll pay Friday — draft it, don't send."**
```
reply_to_email_by_number(email_number=1, reply_text="Thanks — I'll pay by Friday.")
# → returns draft_id; show preview to the user
```

**"Okay, send it."**
```
send_draft(draft_id="…")             # returns CONFIRMATION_REQUIRED + preview
send_draft(draft_id="…", confirm=true)
```

**"Clean up: archive the newsletter and mark the rest read."**
```
search_emails(sender="newsletter", days=14)
archive_email(email_number=1)
mark_as_read(email_number=2)
```

**"What needs my attention today?"**
```
prioritize_inbox(days=1)             # you rank + explain why
```

---

## Authentication & Security

- **No OAuth, tokens, API keys, or passwords** are used or stored by this server.
  It automates the Outlook application the Windows user is already signed into, so
  there are **no secrets for this process to leak**.
- **Permission scope** equals whatever the signed-in Outlook profile can do
  (read/search/draft/send/move/categorise/tasks). Because write/send tools are
  enabled, treat this server like giving an assistant your open mailbox — run it
  only with trusted MCP clients on a trusted machine.
- **Logging** goes to **stderr only** (stdout is reserved for the MCP protocol).
  Email bodies and reply text are **never logged** unless `OUTLOOK_MCP_DEBUG` is
  enabled, and even then only a short, length-capped preview.
- **Outlook security prompts:** depending on your Outlook/antivirus configuration,
  programmatic `Send` may raise a security prompt. The server surfaces this as a
  structured `ACTION_FAILED` with guidance rather than hanging.

---

## Error handling

Errors are structured and branchable. Common `error_code`s:

| Code | Meaning |
| --- | --- |
| `OUTLOOK_NOT_AVAILABLE` | Not on Windows / pywin32 missing. |
| `OUTLOOK_CONNECTION_FAILED` | Outlook not running or not signed in (retryable). |
| `INVALID_PARAMETER` | A parameter was out of range / missing. |
| `INVALID_RECIPIENT` | One or more addresses failed validation. |
| `EMAIL_NOT_FOUND` / `FOLDER_NOT_FOUND` / `DRAFT_NOT_FOUND` | Lookup failed. |
| `NO_LISTING_CONTEXT` | Used `email_number` before any listing — pass `entry_id`. |
| `CONFIRMATION_REQUIRED` | A send needs `confirm=true` (preview returned). |
| `ATTACHMENT_TOO_LARGE` / `UNSUPPORTED_ATTACHMENT` / `ATTACHMENT_NOT_FOUND` | Attachment issues. |
| `ACTION_FAILED` | Generic failure (raw details only in debug mode). |

---

## Development & testing

The test-suite uses fakes and **does not require Windows, Outlook, or pywin32**,
so it runs anywhere (including CI):

```bash
pip install -r requirements-dev.txt
python -m pytest -q
```

What's covered (100+ tests): recipient validation, search predicates &
pagination, response envelopes, body cleaning / quote trimming, the email
formatter, draft-vs-send safety, confirmation guards, archive/trash/read-state,
labels, attachments, the deterministic **triage scorer**, the **web API**
(FastAPI `TestClient` over fakes, incl. token auth + send-confirm), and the
**scheduled digest** builder.

```bash
python -m pytest -q          # all surfaces, no Outlook needed
python outlook_web.py        # try the dashboard locally (needs Outlook)
```

Project layout:

```
outlook_mcp_server.py   # FastMCP server + COM glue + tool definitions
outlook_helpers.py      # pure, COM-free, unit-tested helpers (incl. triage scorer)
outlook_web.py          # localhost FastAPI dashboard + persistent COM worker
web/index.html          # single-page front-end (vanilla JS, no build step)
outlook_schedule.py     # Task Scheduler entry point → digest.json/.html
tests/                  # pytest suite (fakes in conftest.py)
```

---

## Troubleshooting

- **`OUTLOOK_NOT_AVAILABLE`** — you're not on Windows or `pywin32` isn't installed
  (`pip install pywin32`).
- **`OUTLOOK_CONNECTION_FAILED`** — open Outlook (classic desktop) and sign in.
- **Sends seem to hang / get blocked** — an Outlook/antivirus security prompt may
  be intercepting programmatic send; approve it or adjust trust settings.
- **`FOLDER_NOT_FOUND` on archive** — create an `Archive` folder or set
  `OUTLOOK_ARCHIVE_FOLDER` to an existing folder name.
- **Protocol/JSON errors in the client** — ensure nothing writes to stdout; this
  server logs to stderr by design. Don't add `print()` calls.

---

## Known limitations

- **Windows + classic Outlook only** (COM). No macOS/Linux, no "new Outlook".
- **No permanent delete** by design — only move-to-trash.
- `read_thread` reconstructs threads from Inbox + Sent within a day window; very
  old or cross-folder messages may be omitted.
- Categories are Outlook's per-mailbox **categories**, not Gmail-style labels.
- HTML email is read as Outlook's plain-text conversion (with an HTML fallback);
  rich formatting isn't preserved on read.
- Search scans up to a bounded number of recent items per call for latency; widen
  `days` / use `offset` to page through more.
```
