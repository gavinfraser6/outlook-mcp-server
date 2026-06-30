"""Server-level tests using monkeypatched Outlook fakes (no live mailbox)."""

from __future__ import annotations

import datetime
import json

import pytest

import outlook_helpers as H
from outlook_helpers import ErrorCode


def _load(result):
    """Tools return JSON strings; parse to a dict for assertions."""
    return json.loads(result)


# --- decorator / error envelope -------------------------------------------

class TestErrorEnvelope:
    def test_outlook_error_serialized(self):
        import outlook_mcp_server as s

        @s.email_tool
        def boom():
            raise s.OutlookError(ErrorCode.FOLDER_NOT_FOUND, "nope", details="hint")

        out = _load(boom())
        assert out["success"] is False
        assert out["error_code"] == "FOLDER_NOT_FOUND"
        assert out["details"] == "hint"

    def test_generic_error_sanitized(self, monkeypatch):
        import outlook_mcp_server as s
        monkeypatch.delenv("OUTLOOK_MCP_DEBUG", raising=False)

        @s.email_tool
        def boom():
            raise ValueError("super secret internal detail")

        out = _load(boom())
        assert out["success"] is False
        assert out["error_code"] == "ACTION_FAILED"
        # raw exception text must not leak unless debug mode is on
        assert "super secret" not in json.dumps(out)


# --- search & listing ------------------------------------------------------

class TestSearch:
    def test_search_keyword(self, server):
        s, ns, outlook = server
        out = _load(s.search_emails(keyword="invoice", days=30))
        assert out["success"] is True
        assert out["page_info"]["total_matched"] == 1
        assert out["results"][0]["subject"].startswith("Invoice")
        assert out["results"][0]["entry_id"] == "A"
        assert out["results"][0]["email_number"] == 1

    def test_search_unread_only(self, server):
        s, ns, outlook = server
        out = _load(s.search_emails(days=30, unread_only=True))
        subjects = [r["subject"] for r in out["results"]]
        assert "Lunch tomorrow?" not in subjects  # that one is read
        assert len(out["results"]) == 2

    def test_search_has_attachments(self, server):
        s, ns, outlook = server
        out = _load(s.search_emails(days=30, has_attachments=True))
        assert len(out["results"]) == 1
        assert out["results"][0]["entry_id"] == "A"

    def test_search_exclude(self, server):
        s, ns, outlook = server
        out = _load(s.search_emails(keyword="notes", days=30, exclude="spam-word"))
        assert out["page_info"]["total_matched"] == 0

    def test_search_pagination(self, server):
        s, ns, outlook = server
        out = _load(s.search_emails(days=30, max_results=2, offset=0))
        assert out["page_info"]["returned"] == 2
        assert out["page_info"]["has_more"] is True
        assert out["page_info"]["next_offset"] == 2

    def test_empty_result_is_success(self, server):
        s, ns, outlook = server
        out = _load(s.search_emails(keyword="nonexistent-term", days=30))
        assert out["success"] is True
        assert out["results"] == []

    def test_invalid_days(self, server):
        s, ns, outlook = server
        out = _load(s.search_emails(keyword="x", days=9999))
        assert out["success"] is False
        assert out["error_code"] == "INVALID_PARAMETER"

    def test_list_and_count_unread(self, server):
        s, ns, outlook = server
        out = _load(s.count_unread_emails())
        assert out["unread_count"] == 2


# --- conversation insights -------------------------------------------------

class TestConversationInsights:
    def test_groups_messages_into_conversations(self, server):
        from conftest import FakeMail

        s, ns, outlook = server
        now = datetime.datetime.now()
        follow_on = FakeMail(
            entry_id="A2",
            subject="Re: Invoice #42 due Friday",
            sender_name="Acme Billing",
            sender_email="billing@acme.com",
            body="Can you confirm payment today?",
            received=now - datetime.timedelta(minutes=20),
            unread=True,
            conversation_id="C1",
        )
        ns.GetDefaultFolder(H.OL_FOLDER_INBOX).Items.append(follow_on)
        ns.register(follow_on)

        out = _load(s.conversation_insights(days=30, max_results=10))

        assert out["success"] is True
        conv = next(c for c in out["conversations"] if c["conversation_id"] == "C1")
        assert conv["message_count"] == 2
        assert conv["subject"] == "Invoice #42 due Friday"
        assert conv["unread_count"] == 2
        assert conv["follow_up"]["state"] == "reply_owed"

    def test_follow_up_hints_are_bidirectional(self, server, monkeypatch):
        from conftest import FakeMail

        s, ns, outlook = server
        monkeypatch.setattr(s, "_get_my_email", lambda namespace: "me@example.com")
        now = datetime.datetime.now()
        inbox = ns.GetDefaultFolder(H.OL_FOLDER_INBOX).Items
        sent = ns.GetDefaultFolder(H.OL_FOLDER_SENT).Items

        older_inbound = FakeMail(
            entry_id="F1", subject="Client proposal",
            sender_name="Client", sender_email="client@example.com",
            body="The proposal looks possible.",
            received=now - datetime.timedelta(days=5),
            unread=False, conversation_id="CFOLLOW",
        )
        my_request = FakeMail(
            entry_id="S1", subject="Re: Client proposal",
            sender_name="Me", sender_email="me@example.com",
            body="Can you send feedback on the proposal?",
            received=now - datetime.timedelta(days=4),
            sent=now - datetime.timedelta(days=4),
            unread=False, conversation_id="CFOLLOW",
        )
        my_update = FakeMail(
            entry_id="S2", subject="Draft approval",
            sender_name="Me", sender_email="me@example.com",
            body="Here is the draft.",
            received=now - datetime.timedelta(days=3),
            sent=now - datetime.timedelta(days=3),
            unread=False, conversation_id="CREPLY",
        )
        their_request = FakeMail(
            entry_id="F2", subject="Re: Draft approval",
            sender_name="Partner", sender_email="partner@example.com",
            body="Can you approve this by Friday?",
            received=now - datetime.timedelta(days=2),
            unread=True, conversation_id="CREPLY",
        )
        for mail in (older_inbound, their_request):
            inbox.append(mail)
            ns.register(mail)
        for mail in (my_request, my_update):
            sent.append(mail)
            ns.register(mail)

        out = _load(s.conversation_insights(days=30, max_results=20, follow_up_days=2))
        by_id = {c["conversation_id"]: c for c in out["conversations"]}

        assert by_id["CFOLLOW"]["follow_up"]["state"] == "waiting_on_them"
        assert by_id["CFOLLOW"]["follow_up"]["due"] is True
        assert by_id["CREPLY"]["follow_up"]["state"] == "reply_owed"
        assert by_id["CREPLY"]["follow_up"]["due"] is True
        assert out["mailbox_insights"]["waiting_on_them"] >= 1
        assert out["mailbox_insights"]["reply_owed"] >= 1


# --- reading ---------------------------------------------------------------

class TestRead:
    def test_get_email_by_number_after_search(self, server):
        s, ns, outlook = server
        s.search_emails(keyword="invoice", days=30)
        out = _load(s.get_email_by_number(email_number=1))
        assert out["success"] is True
        assert out["subject"].startswith("Invoice")
        assert out["body"]
        assert out["attachment_count"] == 1

    def test_get_email_by_entry_id_without_listing(self, server):
        s, ns, outlook = server
        out = _load(s.get_email_by_number(entry_id="B"))
        assert out["subject"] == "Lunch tomorrow?"

    def test_get_email_no_context(self, server):
        s, ns, outlook = server
        out = _load(s.get_email_by_number(email_number=1))
        assert out["success"] is False
        assert out["error_code"] == "NO_LISTING_CONTEXT"

    def test_bad_entry_id(self, server):
        s, ns, outlook = server
        out = _load(s.get_email_by_number(entry_id="DOESNOTEXIST"))
        assert out["error_code"] == "EMAIL_NOT_FOUND"


# --- drafts & sending safety ----------------------------------------------

class TestDraftsAndSending:
    def test_create_draft_saves_not_sends(self, server):
        s, ns, outlook = server
        out = _load(s.create_draft(to="x@y.com", subject="Hi", body="Body"))
        assert out["success"] is True
        assert out["status"] == "draft_saved"
        assert out["draft_id"]
        created = outlook.created[-1]
        assert created.saved is True
        assert created.sent_flag is False

    def test_create_draft_rejects_bad_recipient(self, server):
        s, ns, outlook = server
        out = _load(s.create_draft(to="not-an-email", subject="Hi", body="B"))
        assert out["success"] is False
        assert out["error_code"] == "INVALID_RECIPIENT"
        assert "not-an-email" in out["invalid_recipients"]

    def test_send_email_requires_confirm(self, server):
        s, ns, outlook = server
        out = _load(s.send_email(to="x@y.com", subject="Hi", body="B"))
        assert out["success"] is False
        assert out["error_code"] == "CONFIRMATION_REQUIRED"
        assert out["preview"]["to"] == ["x@y.com"]
        # nothing should have been created/sent
        assert all(not m.sent_flag for m in outlook.created)

    def test_send_email_with_confirm_sends(self, server):
        s, ns, outlook = server
        out = _load(s.send_email(to="x@y.com", subject="Hi", body="B", confirm=True))
        assert out["success"] is True
        assert out["status"] == "sent"
        assert outlook.created[-1].sent_flag is True

    def test_send_email_invalid_recipient_blocks_before_confirm(self, server):
        s, ns, outlook = server
        out = _load(s.send_email(to="bad", subject="Hi", body="B", confirm=True))
        assert out["error_code"] == "INVALID_RECIPIENT"
        assert not outlook.created  # never created the item

    def test_compose_email_defaults_to_draft(self, server):
        s, ns, outlook = server
        out = _load(s.compose_email(recipient_email="x@y.com", subject="Hi", body="B"))
        assert out["status"] == "draft_saved"
        assert outlook.created[-1].sent_flag is False

    def test_send_draft_preview_then_send(self, server):
        s, ns, outlook = server
        draft = _load(s.create_draft(to="x@y.com", subject="Hi", body="B"))
        did = draft["draft_id"]
        # preview (no confirm)
        prev = _load(s.send_draft(draft_id=did))
        assert prev["error_code"] == "CONFIRMATION_REQUIRED"
        # confirmed
        sent = _load(s.send_draft(draft_id=did, confirm=True))
        assert sent["status"] == "sent"
        assert ns.GetItemFromID(did).sent_flag is True

    def test_send_draft_unknown_id(self, server):
        s, ns, outlook = server
        out = _load(s.send_draft(draft_id="NOPE", confirm=True))
        assert out["error_code"] == "DRAFT_NOT_FOUND"


# --- reply / forward (draft-first) ----------------------------------------

class TestReplyForward:
    def test_reply_creates_draft_by_default(self, server):
        s, ns, outlook = server
        out = _load(s.reply_to_email_by_number(entry_id="A", reply_text="Thanks!"))
        assert out["status"] == "draft_saved"
        assert "Thanks!" in out["body_preview"]

    def test_reply_requires_text(self, server):
        s, ns, outlook = server
        out = _load(s.reply_to_email_by_number(entry_id="A", reply_text="  "))
        assert out["error_code"] == "INVALID_PARAMETER"

    def test_reply_send_true_sends(self, server):
        s, ns, outlook = server
        out = _load(s.reply_to_email_by_number(entry_id="A", reply_text="Yes", send=True))
        assert out["status"] == "sent"

    def test_forward_validates_recipient(self, server):
        s, ns, outlook = server
        out = _load(s.forward_email(to="bad-addr", entry_id="A"))
        assert out["error_code"] == "INVALID_RECIPIENT"

    def test_forward_creates_draft(self, server):
        s, ns, outlook = server
        out = _load(s.forward_email(to="x@y.com", entry_id="A", comment="FYI"))
        assert out["status"] == "draft_saved"


# --- organise: archive / trash / read state / labels ----------------------

class TestOrganise:
    def test_move_to_trash(self, server):
        s, ns, outlook = server
        out = _load(s.move_to_trash(entry_id="B"))
        assert out["status"] == "moved_to_deleted_items"
        assert ns.GetItemFromID("B").moved_to.Name == "Deleted Items"

    def test_archive_email(self, server):
        s, ns, outlook = server
        out = _load(s.archive_email(entry_id="B"))
        assert out["status"] == "archived"
        assert ns.GetItemFromID("B").moved_to.Name == "Archive"

    def test_mark_read_unread(self, server):
        s, ns, outlook = server
        _load(s.mark_as_read(entry_id="A"))
        assert ns.GetItemFromID("A").UnRead is False
        _load(s.mark_as_unread(entry_id="A"))
        assert ns.GetItemFromID("A").UnRead is True

    def test_apply_and_remove_category(self, server):
        s, ns, outlook = server
        out = _load(s.apply_category(category="Important", entry_id="B"))
        assert "Important" in out["labels"]
        assert "Important" in ns.GetItemFromID("B").Categories
        out2 = _load(s.remove_category(category="Important", entry_id="B"))
        assert "Important" not in out2["labels"]

    def test_list_categories(self, server):
        s, ns, outlook = server
        out = _load(s.list_categories())
        names = [c["name"] for c in out["categories"]]
        assert "Work" in names

    def test_move_requires_destination(self, server):
        s, ns, outlook = server
        out = _load(s.move_email_by_number(destination_folder_name="", entry_id="A"))
        assert out["error_code"] == "INVALID_PARAMETER"

    def test_archive_missing_folder(self, server, monkeypatch):
        s, ns, outlook = server
        # remove the Archive folder so lookup fails
        ns.folders_by_index[6].Folders = []
        ns.Folders = [ns.folders_by_index[6]]
        out = _load(s.archive_email(entry_id="A"))
        assert out["error_code"] == "FOLDER_NOT_FOUND"


# --- attachments -----------------------------------------------------------

class TestAttachments:
    def test_list_attachments(self, server):
        s, ns, outlook = server
        out = _load(s.list_attachments(entry_id="A"))
        assert out["attachment_count"] == 1
        assert out["attachments"][0]["filename"] == "invoice.pdf"

    def test_read_attachment_rejects_binary(self, server):
        s, ns, outlook = server
        out = _load(s.read_attachment(attachment_index=1, entry_id="A"))
        assert out["error_code"] == "UNSUPPORTED_ATTACHMENT"

    def test_read_text_attachment(self, server, tmp_path, monkeypatch):
        from conftest import FakeMail, FakeAttachment
        s, ns, outlook = server
        monkeypatch.setenv("OUTLOOK_ATTACHMENT_DIR", str(tmp_path))
        m = FakeMail(entry_id="T", subject="text",
                     attachments=[FakeAttachment("notes.txt", 20, content="hello file")])
        ns.register(m)
        out = _load(s.read_attachment(attachment_index=1, entry_id="T"))
        assert out["success"] is True
        assert out["content"] == "hello file"

    def test_save_attachment_too_large(self, server, monkeypatch):
        from conftest import FakeMail, FakeAttachment
        s, ns, outlook = server
        monkeypatch.setenv("OUTLOOK_MAX_ATTACHMENT_MB", "0.0001")
        m = FakeMail(entry_id="BIG", attachments=[FakeAttachment("big.bin", 10_000_000)])
        ns.register(m)
        out = _load(s.save_attachment(attachment_index=1, entry_id="BIG"))
        assert out["error_code"] == "ATTACHMENT_TOO_LARGE"

    def test_attachment_index_out_of_range(self, server):
        s, ns, outlook = server
        out = _load(s.list_attachments(entry_id="B"))  # B has no attachments
        assert out["attachment_count"] == 0
        out2 = _load(s.save_attachment(attachment_index=1, entry_id="B"))
        assert out2["error_code"] == "ATTACHMENT_NOT_FOUND"


# --- outlook-not-available path -------------------------------------------

class TestNotAvailable:
    def test_reports_unavailable(self, monkeypatch):
        import outlook_mcp_server as s
        monkeypatch.setattr(s, "_WIN32_AVAILABLE", False)
        # real _connect should raise OUTLOOK_NOT_AVAILABLE
        out = _load(s.list_folders())
        assert out["success"] is False
        assert out["error_code"] == "OUTLOOK_NOT_AVAILABLE"
