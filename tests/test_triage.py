"""Tests for the deterministic triage scorer."""

from __future__ import annotations

import datetime

import outlook_helpers as H


def _email(**over):
    base = {
        "subject": "Hello", "body": "Just saying hi.",
        "sender": "Sam Person", "sender_email": "sam@corp.com",
        "unread": True, "importance": "Normal", "flagged": "none",
        "has_attachments": False,
        "received_time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    base.update(over)
    return base


class TestTriageScore:
    def test_urgent_invoice_scores_high(self):
        e = _email(subject="URGENT: Invoice overdue, payment required by Friday",
                   body="Please pay today. Can you confirm?", importance="High",
                   has_attachments=True)
        score, reasons = H.triage_score(e)
        assert score >= 10
        joined = " ".join(reasons)
        assert "urgent language" in joined
        assert "mentions invoice/payment" in joined
        assert "high importance" in joined

    def test_automated_is_damped(self):
        e = _email(subject="Weekly newsletter", body="unsubscribe at the bottom",
                   sender_email="newsletter@news.com")
        score, reasons = H.triage_score(e)
        assert "looks automated/bulk" in reasons
        assert score <= 2

    def test_manager_bonus(self):
        e = _email(sender="Dana Boss")
        score, reasons = H.triage_score(e, manager_name="Dana Boss")
        assert "from your manager" in reasons
        assert score >= 6  # unread(2)+manager(4)+recency

    def test_recency_bumps(self):
        recent = _email()
        old = _email(received_time=(datetime.datetime.now() - datetime.timedelta(days=20)).strftime("%Y-%m-%d %H:%M:%S"))
        s_recent, _ = H.triage_score(recent)
        s_old, _ = H.triage_score(old)
        assert s_recent > s_old

    def test_question_adds_reason(self):
        e = _email(body="Are you available tomorrow?")
        _, reasons = H.triage_score(e)
        assert "contains a question" in reasons

    def test_score_never_negative(self):
        e = _email(subject="newsletter", body="unsubscribe",
                   sender_email="noreply@x.com", unread=False,
                   received_time=(datetime.datetime.now() - datetime.timedelta(days=40)).strftime("%Y-%m-%d %H:%M:%S"))
        score, _ = H.triage_score(e)
        assert score >= 0


class TestRankForTriage:
    def test_orders_most_urgent_first(self):
        a = _email(subject="Lunch?", body="grab lunch?")
        b = _email(subject="URGENT payment overdue invoice", importance="High")
        c = _email(subject="FYI newsletter", body="unsubscribe",
                   sender_email="newsletter@x.com")
        ranked = H.rank_for_triage([a, b, c])
        assert ranked[0]["subject"].startswith("URGENT")
        assert all("triage_score" in r and "triage_reasons" in r for r in ranked)
        scores = [r["triage_score"] for r in ranked]
        assert scores == sorted(scores, reverse=True)


class TestDaslBuilder:
    def test_escapes_quotes(self):
        out = H.escape_dasl_literal("O'Brien %weird% \"x\"")
        assert "''" in out
        assert "%" not in out and '"' not in out

    def test_build_restriction_clauses(self):
        now = datetime.datetime(2026, 1, 1, 9, 0, 0)
        r = H.build_inbox_restriction(threshold=now, unread_only=True, subject="report")
        assert "[ReceivedTime] >=" in r
        assert "[UnRead] = True" in r
        assert "[Subject] Like '%report%'" in r

    def test_build_restriction_none_when_empty(self):
        assert H.build_inbox_restriction() is None
