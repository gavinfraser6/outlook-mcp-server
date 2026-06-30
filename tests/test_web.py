"""Web API tests using FastAPI's TestClient over the fake Outlook (no COM)."""

from __future__ import annotations

import json

import pytest
from fastapi.testclient import TestClient

import outlook_web


@pytest.fixture
def client(server, monkeypatch):
    """A TestClient whose Outlook calls hit the in-memory fakes."""
    s, ns, outlook = server  # monkeypatches omcp._connect/_namespace to fakes
    monkeypatch.delenv("OUTLOOK_WEB_TOKEN", raising=False)

    async def call(tool, **kwargs):
        # Tools are synchronous and operate on the fakes; just parse the envelope.
        return json.loads(tool(**kwargs))

    app = outlook_web.create_app(call)
    return TestClient(app), ns, outlook


class TestReadOnly:
    def test_health(self, client):
        c, ns, outlook = client
        r = c.get("/api/health")
        assert r.status_code == 200 and r.json()["success"] is True

    def test_index_served(self, client):
        c, ns, outlook = client
        r = c.get("/")
        assert r.status_code == 200
        assert "Outlook Assistant" in r.text

    def test_unread_count(self, client):
        c, ns, outlook = client
        r = c.get("/api/unread_count")
        assert r.json()["unread_count"] == 2

    def test_triage_ranks_invoice_first(self, client):
        c, ns, outlook = client
        r = c.get("/api/triage?days=30&limit=10")
        body = r.json()
        assert body["success"] is True
        assert body["results"][0]["subject"].startswith("Invoice")
        assert "triage_score" in body["results"][0]
        assert "triage_reasons" in body["results"][0]

    def test_search(self, client):
        c, ns, outlook = client
        r = c.get("/api/search?keyword=invoice&days=30")
        assert r.json()["page_info"]["total_matched"] == 1

    def test_email_and_thread(self, client):
        c, ns, outlook = client
        assert c.get("/api/email?entry_id=A").json()["subject"].startswith("Invoice")
        t = c.get("/api/thread?entry_id=A").json()
        assert t["success"] is True


class TestActions:
    def test_archive(self, client):
        c, ns, outlook = client
        r = c.post("/api/archive", json={"entry_id": "B"})
        assert r.json()["status"] == "archived"
        assert ns.GetItemFromID("B").moved_to.Name == "Archive"

    def test_trash(self, client):
        c, ns, outlook = client
        r = c.post("/api/trash", json={"entry_id": "B"})
        assert r.json()["status"] == "moved_to_deleted_items"

    def test_mark_read(self, client):
        c, ns, outlook = client
        c.post("/api/mark", json={"entry_id": "A", "read": True})
        assert ns.GetItemFromID("A").UnRead is False

    def test_category(self, client):
        c, ns, outlook = client
        r = c.post("/api/category", json={"entry_id": "B", "category": "Later"})
        assert "Later" in r.json()["labels"]

    def test_missing_field_is_400(self, client):
        c, ns, outlook = client
        r = c.post("/api/archive", json={})
        assert r.status_code == 400
        assert r.json()["error_code"] == "INVALID_PARAMETER"


class TestDraftsAndSendSafety:
    def test_create_draft(self, client):
        c, ns, outlook = client
        r = c.post("/api/draft", json={"to": "x@y.com", "subject": "Hi", "body": "B"})
        assert r.json()["status"] == "draft_saved"

    def test_send_email_requires_confirm(self, client):
        c, ns, outlook = client
        r = c.post("/api/send_email", json={"to": "x@y.com", "subject": "Hi", "body": "B"})
        assert r.json()["error_code"] == "CONFIRMATION_REQUIRED"
        assert all(not m.sent_flag for m in outlook.created)

    def test_send_email_confirmed(self, client):
        c, ns, outlook = client
        r = c.post("/api/send_email", json={"to": "x@y.com", "subject": "Hi",
                                            "body": "B", "confirm": True})
        assert r.json()["status"] == "sent"

    def test_reply_creates_draft(self, client):
        c, ns, outlook = client
        r = c.post("/api/reply", json={"entry_id": "A", "reply_text": "Thanks"})
        assert r.json()["status"] == "draft_saved"

    def test_bad_recipient_rejected(self, client):
        c, ns, outlook = client
        r = c.post("/api/draft", json={"to": "not-email", "subject": "x", "body": "y"})
        assert r.json()["error_code"] == "INVALID_RECIPIENT"


class TestAuth:
    def test_token_required_when_set(self, server, monkeypatch):
        s, ns, outlook = server
        monkeypatch.setenv("OUTLOOK_WEB_TOKEN", "secret")

        async def call(tool, **kwargs):
            return json.loads(tool(**kwargs))

        c = TestClient(outlook_web.create_app(call))
        # no token -> 401
        assert c.get("/api/triage?days=7").status_code == 401
        # health stays open
        assert c.get("/api/health").status_code == 200
        # with token -> ok
        assert c.get("/api/triage?days=7", headers={"X-Outlook-Token": "secret"}).status_code == 200
