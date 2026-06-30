"""Unit tests for the pure helpers (no Outlook required)."""

from __future__ import annotations

import datetime

import outlook_helpers as H
from outlook_helpers import ErrorCode


# --- recipient validation --------------------------------------------------

class TestRecipientValidation:
    def test_valid_addresses(self):
        for addr in ["a@b.com", "first.last@sub.domain.co", "x+tag@gmail.com"]:
            assert H.is_valid_email(addr), addr

    def test_invalid_addresses(self):
        for addr in ["", "no-at-sign", "a@b", "a@@b.com", "a b@c.com", "a@b .com"]:
            assert not H.is_valid_email(addr), addr

    def test_parse_and_validate_mixed(self):
        valid, invalid = H.validate_recipients("a@b.com, bad, c@d.co; e@f")
        assert valid == ["a@b.com", "c@d.co"]
        assert invalid == ["bad", "e@f"]

    def test_empty(self):
        assert H.validate_recipients(None) == ([], [])
        assert H.parse_recipients("") == []


# --- envelopes -------------------------------------------------------------

class TestEnvelopes:
    def test_success_shape(self):
        p = H.make_success("do_thing", count=3)
        assert p["success"] is True
        assert p["action"] == "do_thing"
        assert p["count"] == 3

    def test_error_shape(self):
        p = H.make_error(ErrorCode.INVALID_RECIPIENT, "bad", details="fix it",
                         retryable=False, invalid_recipients=["x"])
        assert p["success"] is False
        assert p["error_code"] == "INVALID_RECIPIENT"
        assert p["message"] == "bad"
        assert p["details"] == "fix it"
        assert p["retryable"] is False
        assert p["invalid_recipients"] == ["x"]

    def test_to_json_roundtrip(self):
        import json
        s = H.to_json({"a": 1, "dt": datetime.datetime(2026, 1, 1)})
        assert json.loads(s)["a"] == 1


# --- search predicate ------------------------------------------------------

class TestEmailMatches:
    base = {
        "subject": "Invoice 42 due",
        "sender": "Acme Billing",
        "sender_email": "billing@acme.com",
        "body": "Please pay the invoice.",
        "recipients": ["Me <me@x.com>"],
        "categories": "Finance; Work",
        "unread": True,
        "has_attachments": True,
    }

    def test_keyword_hits_subject_or_body(self):
        assert H.email_matches(self.base, keyword="invoice")
        assert H.email_matches(self.base, keyword="pay")
        assert not H.email_matches(self.base, keyword="zzz")

    def test_keyword_or_alternation(self):
        assert H.email_matches(self.base, keyword="statement OR invoice")
        assert not H.email_matches(self.base, keyword="statement OR receipt")

    def test_sender_subject_recipient(self):
        assert H.email_matches(self.base, sender="acme.com")
        assert H.email_matches(self.base, subject="invoice")
        assert H.email_matches(self.base, recipient="me@x.com")
        assert not H.email_matches(self.base, sender="other.com")

    def test_unread_and_attachments(self):
        assert H.email_matches(self.base, unread_only=True)
        assert H.email_matches(self.base, has_attachments=True)
        assert not H.email_matches(self.base, has_attachments=False)

    def test_category_and_exclude(self):
        assert H.email_matches(self.base, category="finance")
        assert not H.email_matches(self.base, exclude=["invoice"])
        assert H.email_matches(self.base, exclude=["nothere"])

    def test_exact_phrase(self):
        assert H.email_matches(self.base, exact_phrase="please pay")
        assert not H.email_matches(self.base, exact_phrase="please refund")

    def test_filters_are_anded(self):
        assert H.email_matches(self.base, keyword="invoice", sender="acme")
        assert not H.email_matches(self.base, keyword="invoice", sender="other")


# --- pagination ------------------------------------------------------------

class TestPagination:
    def test_basic_slice(self):
        page, info = H.paginate(list(range(10)), offset=0, limit=3)
        assert page == [0, 1, 2]
        assert info["total_matched"] == 10
        assert info["has_more"] is True
        assert info["next_offset"] == 3

    def test_last_page(self):
        page, info = H.paginate(list(range(5)), offset=3, limit=10)
        assert page == [3, 4]
        assert info["has_more"] is False
        assert info["next_offset"] is None

    def test_limit_capped(self):
        _, info = H.paginate(list(range(500)), offset=0, limit=99999)
        assert info["limit"] == H.MAX_PAGE_SIZE


# --- body cleaning ---------------------------------------------------------

class TestBodyCleaning:
    def test_snippet_truncates(self):
        s = H.make_snippet("word " * 100, length=20)
        assert len(s) <= 21 and s.endswith("…")

    def test_snippet_collapses_whitespace(self):
        assert H.make_snippet("a\n\n   b\t c") == "a b c"

    def test_trim_quoted_reply(self):
        body = "My reply here.\nThanks!\n-----Original Message-----\nFrom: x@y.com"
        trimmed, was = H.trim_quoted_reply(body)
        assert was is True
        assert "Original Message" not in trimmed
        assert "My reply here." in trimmed

    def test_trim_does_not_gut_forward_starting_with_from(self):
        body = "From: someone\nSubject: hi\nbody"
        trimmed, was = H.trim_quoted_reply(body)
        assert was is False
        assert trimmed.startswith("From:")

    def test_strip_html(self):
        assert H.strip_html("<p>Hi <b>there</b></p><script>bad()</script>") == "Hi there"


# --- categories ------------------------------------------------------------

class TestCategories:
    def test_parse_and_join(self):
        assert H.parse_categories("Work; Personal ;; Urgent") == ["Work", "Personal", "Urgent"]
        assert H.join_categories(["A", "A", "B"]) == "A; B"


# --- attachment helpers ----------------------------------------------------

class TestAttachments:
    def test_sanitize_filename(self):
        assert H.sanitize_filename("../../etc/passwd") == "passwd"
        assert H.sanitize_filename("weird/na;me*.txt") == "na_me_.txt"
        assert H.sanitize_filename("") == "attachment"

    def test_is_safe_text(self):
        assert H.is_safe_text_attachment("notes.txt")
        assert H.is_safe_text_attachment("data.CSV")
        assert not H.is_safe_text_attachment("malware.exe")
        assert not H.is_safe_text_attachment("photo.png")


# --- formatter -------------------------------------------------------------

class TestFormatter:
    def test_format_minimal_object(self):
        from conftest import FakeMail, FakeAttachment
        m = FakeMail(subject="Hi", body="Body text",
                     attachments=[FakeAttachment("a.pdf", 10)],
                     categories="Work; Urgent", importance=2)
        d = H.format_email_item(m)
        assert d["subject"] == "Hi"
        assert d["sender_email"] == "alice@example.com"
        assert d["has_attachments"] is True
        assert d["attachment_count"] == 1
        assert d["importance"] == "High"
        assert d["snippet"] == "Body text"
        assert "Bob" in d["recipients"][0]

    def test_format_html_fallback(self):
        from conftest import FakeMail
        m = FakeMail(body="", html_body="<p>From <b>HTML</b></p>")
        d = H.format_email_item(m)
        assert "HTML" in d["body"]
        assert any("HTML" in w for w in d.get("warnings", []))


# --- redaction -------------------------------------------------------------

class TestRedaction:
    def test_redacts_by_default(self, monkeypatch):
        monkeypatch.delenv("OUTLOOK_MCP_DEBUG", raising=False)
        out = H.redact("secret body content")
        assert "secret" not in out
        assert "redacted" in out

    def test_preview_in_debug(self, monkeypatch):
        monkeypatch.setenv("OUTLOOK_MCP_DEBUG", "1")
        out = H.redact("hello world")
        assert "hello world" in out
