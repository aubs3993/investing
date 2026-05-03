"""Smoke tests for the audit_tickers utility.

Builds a tiny in-memory-style DB, runs the audit, and asserts the report
captures the obvious cases (cashtag-only, common English word, etc.).
"""

from __future__ import annotations

import io
import sqlite3
import sys

import pytest

from social.wsb_momentum import audit_tickers, db


@pytest.fixture
def populated_conn(tmp_path):
    c = db.init_db(tmp_path / "wsb.db")
    c.row_factory = sqlite3.Row

    # Helper: insert a post with one snapshot, then map to tickers.
    def add(post_id, title, body, tickers, score):
        db.upsert_post(
            c, post_id=post_id, title=title, author="u", body=body,
            created_utc=1, first_seen_utc=1, source_listing="hot",
            permalink=None, flair=None,
        )
        db.add_post_tickers(c, post_id, tickers)
        db.record_upvote_snapshot(
            c, post_id=post_id, snapshot_utc=1000,
            score=score, num_comments=0, upvote_ratio=0.9,
        )

    # NOTE: post.first_seen_utc is set to 1 above. The audit window is 7 days
    # via SQL strftime('%s','now','-7 days'), which yields a large positive
    # cutoff — so first_seen_utc=1 falls OUTSIDE the window. To exercise the
    # report path we need recent timestamps. Re-insert with a current ts.
    import time as _t
    now = int(_t.time())
    c.execute("DELETE FROM upvote_snapshots")
    c.execute("DELETE FROM post_tickers")
    c.execute("DELETE FROM posts")

    def add_recent(post_id, title, body, tickers, score):
        db.upsert_post(
            c, post_id=post_id, title=title, author="u", body=body,
            created_utc=now, first_seen_utc=now, source_listing="hot",
            permalink=None, flair=None,
        )
        db.add_post_tickers(c, post_id, tickers)
        db.record_upvote_snapshot(
            c, post_id=post_id, snapshot_utc=now,
            score=score, num_comments=0, upvote_ratio=0.9,
        )

    # A real, confidently-discussed ticker — should NOT appear suspicious.
    add_recent(
        "p1", "TSLA earnings preview",
        "Long $TSLA into Q3, expecting beat on auto margins. Detailed DD follows..."
        + ("padding " * 100),
        ["TSLA"], 800,
    )
    # A common English word in disguise — should be flagged.
    add_recent(
        "p2", "ALL signs point up",
        "ALL of my positions are green today. Keep going.",
        ["ALL"], 25,
    )
    # Single-letter ticker, no cashtag — should score high on suspicion.
    add_recent(
        "p3", "Discussing ratios",
        "P/E and P/S ratios across the cohort. Everything trades at premium P now.",
        ["P"], 50,
    )

    yield c
    c.close()


def test_default_report_runs_and_flags_obvious_noise(populated_conn, monkeypatch, capsys):
    # Empty the blacklist for this test so the audit tool's "is this a
    # candidate?" logic can be verified in isolation, independent of which
    # noise words are already permanently blocked in production config.
    monkeypatch.setattr(audit_tickers, "TICKER_BLACKLIST", frozenset())
    monkeypatch.setattr(audit_tickers, "_connect", lambda _p: populated_conn)
    audit_tickers.main([])
    out = capsys.readouterr().out

    # Header rendered.
    assert "WSB ticker audit" in out
    assert "Suggested blacklist additions" in out

    # The obvious noise tickers must appear in the suggested-additions block.
    suggestions_section = out.split("Suggested blacklist additions")[1]
    assert '"ALL"' in suggestions_section, out
    assert '"P"' in suggestions_section, out

    # The clean ticker should NOT be in the suggestions section.
    assert '"TSLA"' not in suggestions_section, out


def test_single_ticker_deep_dive_runs(populated_conn, monkeypatch, capsys):
    monkeypatch.setattr(audit_tickers, "_connect", lambda _p: populated_conn)
    audit_tickers.main(["--ticker", "P"])
    out = capsys.readouterr().out

    assert "Deep-dive: ticker 'P'" in out
    assert "1 post(s) mention P" in out
    # Mention contexts surface — the bare matches in the body should be shown.
    assert "P/E" in out or "P/S" in out


def test_single_ticker_no_mentions(populated_conn, monkeypatch, capsys):
    monkeypatch.setattr(audit_tickers, "_connect", lambda _p: populated_conn)
    audit_tickers.main(["--ticker", "ZZZZX"])
    out = capsys.readouterr().out
    assert "No posts mention ZZZZX" in out


def test_blacklist_bug_bonus_only_fires_on_recent_extractions(tmp_path, monkeypatch, capsys):
    """Refinement: a blacklisted ticker that only appears in OLD post_tickers
    rows must NOT trigger the +10 'still extracted (bug?)' bonus. Only recent
    extractions (within BLACKLIST_BUG_CUTOFF_HOURS) count."""
    import time as _t
    c = db.init_db(tmp_path / "wsb.db")
    import sqlite3 as _sq
    c.row_factory = _sq.Row

    now = int(_t.time())
    old = now - 5 * 86400          # 5 days ago — outside the 24h cutoff
    fresh = now - 3600             # 1 hour ago — inside the 24h cutoff

    # post 1: OLD, mentions blacklisted "ZZZ"
    db.upsert_post(
        c, post_id="old1", title="Talking about ZZZ stocks",
        author="u", body="ZZZ ZZZ ZZZ great pick" + ("padding " * 80),
        created_utc=old, first_seen_utc=old, source_listing="hot",
        permalink=None, flair=None,
    )
    db.add_post_tickers(c, "old1", ["ZZZ"])
    db.record_upvote_snapshot(c, post_id="old1", snapshot_utc=old, score=100, num_comments=10, upvote_ratio=0.9)

    # post 2: ALSO OLD, mentions blacklisted "QQQQ"
    db.upsert_post(
        c, post_id="old2", title="QQQQ all over",
        author="u", body="QQQQ QQQQ QQQQ" + ("padding " * 80),
        created_utc=old, first_seen_utc=old, source_listing="hot",
        permalink=None, flair=None,
    )
    db.add_post_tickers(c, "old2", ["QQQQ"])
    db.record_upvote_snapshot(c, post_id="old2", snapshot_utc=old, score=100, num_comments=10, upvote_ratio=0.9)

    # post 3: FRESH, mentions blacklisted "WWW" (this one IS a real bug — extractor missed the blacklist)
    db.upsert_post(
        c, post_id="fresh1", title="WWW going wild",
        author="u", body="WWW WWW WWW" + ("padding " * 80),
        created_utc=fresh, first_seen_utc=fresh, source_listing="hot",
        permalink=None, flair=None,
    )
    db.add_post_tickers(c, "fresh1", ["WWW"])
    db.record_upvote_snapshot(c, post_id="fresh1", snapshot_utc=fresh, score=100, num_comments=10, upvote_ratio=0.9)

    # Pretend all three of these tickers are blacklisted in production config.
    monkeypatch.setattr(audit_tickers, "TICKER_BLACKLIST", frozenset({"ZZZ", "QQQQ", "WWW"}))
    monkeypatch.setattr(audit_tickers, "_connect", lambda _p: c)

    audit_tickers.main([])
    out = capsys.readouterr().out

    # Old extractions: must NOT carry the "still extracted (bug?)" bonus —
    # they're historical leftovers from before the blacklist was tightened.
    assert "ZZZ" in out
    assert "QQQQ" in out
    # The bonus reason text must not appear next to ZZZ or QQQQ.
    zzz_lines = [ln for ln in out.splitlines() if ln.strip().startswith("ZZZ")]
    qqqq_lines = [ln for ln in out.splitlines() if ln.strip().startswith("QQQQ")]
    for ln in zzz_lines + qqqq_lines:
        assert "still extracted" not in ln and "(bug?)" not in ln, ln

    # Fresh extraction of a blacklisted ticker IS a bug — bonus must fire.
    www_lines = [ln for ln in out.splitlines() if ln.strip().startswith("WWW")]
    assert any("(bug?)" in ln for ln in www_lines), out
    c.close()
