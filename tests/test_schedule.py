"""Tests for the scheduled digest builder (no COM, uses fakes)."""

from __future__ import annotations

import outlook_schedule as sched


class TestBuildDigest:
    def test_digest_shape_and_ranking(self, server):
        s, ns, outlook = server
        d = sched.build_digest(days=30, top=10, follow_up_days=2)
        assert d["unread_count"] == 2
        assert d["needs_attention"], "expected ranked emails"
        # invoice should be the most urgent
        assert d["needs_attention"][0]["subject"].startswith("Invoice")
        assert "generated_at" in d

    def test_render_html_contains_subject(self, server):
        s, ns, outlook = server
        d = sched.build_digest(days=30, top=10, follow_up_days=2)
        html = sched.render_html(d)
        assert "Inbox digest" in html
        assert "Invoice" in html

    def test_build_digest_is_read_only(self, server):
        s, ns, outlook = server
        sched.build_digest(days=30, top=10, follow_up_days=2)
        # nothing should have been sent or moved
        assert not outlook.created
        assert all(m.moved_to is None for m in ns._by_id.values())

    def test_auto_categorize_tags_unread(self, server):
        s, ns, outlook = server
        d = sched.build_digest(days=30, top=10, follow_up_days=2)
        n = sched.auto_categorize(d, "⚑ Needs Reply", min_score=1, limit=10)
        assert n >= 1
        # the invoice (A) is unread + top-ranked -> should be tagged
        assert "Needs Reply" in ns.GetItemFromID("A").Categories
