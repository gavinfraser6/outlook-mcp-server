"""Scheduled inbox digest for the Outlook MCP server.

Designed to be run on a schedule (Windows Task Scheduler) to produce a
deterministic "what needs my attention" digest without any LLM in the loop. It
reuses the very same tool functions the MCP server and web UI use, so behaviour
is consistent everywhere.

By default it is **read-only**: it scans, ranks and writes a digest
(``digest.json`` + ``digest.html``) to a state directory and prints a short
summary. Two optional, explicit side-effects can be enabled:

* ``--auto-categorize`` applies a category (default "⚑ Needs Reply") to the
  top-ranked unread inbound emails so they stand out in Outlook.
* ``--email-to ADDRESS`` emails the digest to yourself (an explicit send you
  opt into by passing the flag).

Examples (Task Scheduler "Program/script" = ``python``, "Arguments" = …)::

    outlook_schedule.py                       # write digest, print summary
    outlook_schedule.py --days 3 --top 20
    outlook_schedule.py --auto-categorize
    outlook_schedule.py --email-to me@corp.com

This is a normal CLI (not the MCP stdio server), so printing to stdout is fine.
"""

from __future__ import annotations

import argparse
import datetime
import html
import json
import os
import sys
from typing import Any, Dict, List

import outlook_helpers as H
import outlook_mcp_server as omcp

log = H.get_logger()


def _call(tool, **kwargs) -> Dict[str, Any]:
    """Call an MCP tool function and parse its JSON envelope."""
    return json.loads(tool(**kwargs))


def state_dir() -> str:
    base = (os.environ.get("OUTLOOK_STATE_DIR")
            or os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                            "outlook-mcp"))
    os.makedirs(base, exist_ok=True)
    return base


def build_digest(days: int, top: int, follow_up_days: int) -> Dict[str, Any]:
    """Gather all digest data using the shared tools."""
    unread = _call(omcp.count_unread_emails)
    triage = _call(omcp.triage_inbox, days=days, max_results=top)
    briefing = _call(omcp.generate_morning_briefing,
                     days_to_scan=min(max(days, 1), 14),
                     follow_up_days=follow_up_days)

    needs_attention: List[Dict[str, Any]] = []
    if triage.get("success"):
        for e in triage.get("results", []):
            needs_attention.append({
                "entry_id": e.get("entry_id"),
                "subject": e.get("subject"),
                "from": e.get("from"),
                "date": e.get("date"),
                "score": e.get("triage_score"),
                "reasons": e.get("triage_reasons", []),
                "unread": e.get("unread"),
            })

    awaiting_reply = []
    calendar = []
    tasks = []
    if briefing.get("success"):
        for t in briefing.get("conversation_threads", []):
            if t.get("follow_up_suggestion"):
                awaiting_reply.append({"subject": t.get("subject"),
                                       "note": t.get("follow_up_suggestion")})
        calendar = briefing.get("todays_calendar", [])
        tasks = briefing.get("todays_tasks", [])

    return {
        "generated_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "window_days": days,
        "unread_count": unread.get("unread_count") if unread.get("success") else None,
        "needs_attention": needs_attention,
        "awaiting_reply": awaiting_reply,
        "todays_calendar": calendar,
        "todays_tasks": tasks,
    }


def render_html(d: Dict[str, Any]) -> str:
    def chip(text):
        return f'<span style="background:#eef;border-radius:8px;padding:1px 6px;font-size:12px;margin-right:4px">{html.escape(str(text))}</span>'

    rows = []
    for e in d["needs_attention"]:
        reasons = " ".join(chip(r) for r in (e.get("reasons") or [])[:4])
        rows.append(
            f'<tr><td style="text-align:center;font-weight:700;color:#b30000">{e.get("score","")}</td>'
            f'<td><b>{html.escape(e.get("subject") or "(no subject)")}</b><br>'
            f'<span style="color:#666">{html.escape(e.get("from") or "")} — {html.escape(e.get("date") or "")}</span><br>'
            f'{reasons}</td></tr>')
    attention = "".join(rows) or '<tr><td colspan="2" style="color:#666">Nothing pressing 🎉</td></tr>'

    follow = "".join(
        f'<li><b>{html.escape(x["subject"] or "")}</b> — {html.escape(x["note"])}</li>'
        for x in d["awaiting_reply"]) or "<li style='color:#666'>None</li>"
    cal = "".join(
        f'<li>{html.escape(c.get("start",""))}–{html.escape(c.get("end",""))} {html.escape(c.get("subject",""))} '
        f'<span style="color:#666">{html.escape(c.get("location",""))}</span></li>'
        for c in d["todays_calendar"]) or "<li style='color:#666'>No events</li>"
    tasks = "".join(
        f'<li>{html.escape(t.get("subject",""))} <span style="color:#666">(due {html.escape(t.get("due_date",""))})</span></li>'
        for t in d["todays_tasks"]) or "<li style='color:#666'>No tasks due</li>"

    return f"""<!doctype html><html><head><meta charset="utf-8">
<title>Inbox digest</title></head>
<body style="font:14px/1.5 Segoe UI,Arial,sans-serif;color:#1a1a1a;max-width:760px;margin:0 auto;padding:16px">
<h2 style="margin:0">📧 Inbox digest</h2>
<p style="color:#666;margin:4px 0 16px">{html.escape(d["generated_at"])} • last {d["window_days"]} day(s) •
<b>{d.get("unread_count","?")}</b> unread</p>

<h3>Needs attention</h3>
<table style="border-collapse:collapse;width:100%">
<thead><tr><th style="width:40px">★</th><th style="text-align:left">Email</th></tr></thead>
<tbody>{attention}</tbody></table>

<h3>Awaiting reply (you sent, no response)</h3><ul>{follow}</ul>
<h3>Today's calendar</h3><ul>{cal}</ul>
<h3>Today's tasks</h3><ul>{tasks}</ul>
<p style="color:#999;font-size:12px">Generated by Outlook Assistant. Nothing was sent or deleted.</p>
</body></html>"""


def auto_categorize(d: Dict[str, Any], category: str, min_score: int, limit: int) -> int:
    """Apply a category to the top unread inbound emails. Returns count tagged."""
    tagged = 0
    for e in d["needs_attention"]:
        if tagged >= limit:
            break
        if e.get("unread") and (e.get("score") or 0) >= min_score and e.get("entry_id"):
            res = _call(omcp.apply_category, entry_id=e["entry_id"], category=category)
            if res.get("success"):
                tagged += 1
    return tagged


def email_digest(to: str, html_body: str, unread_count: Any) -> Dict[str, Any]:
    subject = f"📧 Inbox digest — {unread_count} unread ({datetime.date.today():%a %d %b})"
    # Explicit, user-opted send (the --email-to flag is the explicit intent).
    return _call(omcp.send_email, to=to, subject=subject,
                 body="Open in an HTML-capable client to view the digest.",
                 html_body=html_body, confirm=True)


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Outlook scheduled inbox digest")
    p.add_argument("--days", type=int, default=2, help="Look-back window (default 2)")
    p.add_argument("--top", type=int, default=15, help="Max ranked emails (default 15)")
    p.add_argument("--follow-up-days", type=int, default=2,
                   help="Silence before flagging awaiting-reply (default 2)")
    p.add_argument("--state-dir", default=None, help="Where to write the digest")
    p.add_argument("--auto-categorize", action="store_true",
                   help="Tag top unread inbound emails with a category (modifies mail)")
    p.add_argument("--category", default="⚑ Needs Reply",
                   help="Category to apply when --auto-categorize is set")
    p.add_argument("--min-score", type=int, default=6,
                   help="Min triage score for auto-categorize (default 6)")
    p.add_argument("--max-categorize", type=int, default=10,
                   help="Cap on emails auto-categorized per run (default 10)")
    p.add_argument("--email-to", default=None,
                   help="Email the digest to this address (an explicit send you opt into)")
    p.add_argument("--quiet", action="store_true", help="Suppress stdout summary")
    args = p.parse_args(argv)

    if args.state_dir:
        os.environ["OUTLOOK_STATE_DIR"] = args.state_dir
    out_dir = state_dir()

    try:
        digest = build_digest(args.days, args.top, args.follow_up_days)
    except Exception as exc:  # surface as a clean non-zero exit for the scheduler
        log.error("Digest generation failed: %s", exc, exc_info=H.debug_enabled())
        print(f"ERROR: could not generate digest: {exc}", file=sys.stderr)
        return 2

    html_body = render_html(digest)
    json_path = os.path.join(out_dir, "digest.json")
    html_path = os.path.join(out_dir, "digest.html")
    with open(json_path, "w", encoding="utf-8") as fh:
        json.dump(digest, fh, indent=2, default=str)
    with open(html_path, "w", encoding="utf-8") as fh:
        fh.write(html_body)

    actions = []
    if args.auto_categorize:
        n = auto_categorize(digest, args.category, args.min_score, args.max_categorize)
        actions.append(f"categorized {n} as '{args.category}'")
    if args.email_to:
        res = email_digest(args.email_to, html_body, digest.get("unread_count"))
        actions.append("emailed digest" if res.get("success")
                       else f"email failed: {res.get('message')}")

    if not args.quiet:
        na = digest["needs_attention"]
        print(f"Inbox digest @ {digest['generated_at']}")
        print(f"  unread: {digest.get('unread_count')}  |  needs attention: {len(na)}  "
              f"|  awaiting reply: {len(digest['awaiting_reply'])}")
        for e in na[:5]:
            print(f"   [{e.get('score'):>2}] {e.get('subject')}  — {e.get('from')}")
        if actions:
            print("  actions: " + "; ".join(actions))
        print(f"  written: {html_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
