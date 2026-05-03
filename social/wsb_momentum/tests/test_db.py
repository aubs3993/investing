"""Tests for db.py.

Run from repo root:
    pytest social/wsb_momentum/tests/test_db.py
"""

from __future__ import annotations

import sqlite3
import time

import pytest

from social.wsb_momentum import db


@pytest.fixture
def conn(tmp_path):
    c = db.init_db(tmp_path / "wsb.db")
    yield c
    c.close()


def test_schema_creates_all_tables(conn):
    names = {r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table'"
    ).fetchall()}
    assert {"posts", "post_tickers", "upvote_snapshots", "price_snapshots"} <= names


def test_init_db_is_idempotent(tmp_path):
    p = tmp_path / "wsb.db"
    db.init_db(p).close()
    # Second call must not error (CREATE IF NOT EXISTS for tables AND indexes).
    db.init_db(p).close()


def test_upsert_post_inserts_then_ignores(conn):
    args = dict(
        post_id="t3_abc",
        title="Test",
        author="alice",
        body="body",
        created_utc=100,
        first_seen_utc=200,
        source_listing="hot",
        permalink="/r/wsb/comments/abc",
        flair=None,
    )
    assert db.upsert_post(conn, **args) is True
    # Second call ignored — and source_listing must NOT be overwritten even if
    # we pass a different value (the spec's edge case).
    args2 = {**args, "source_listing": "top"}
    assert db.upsert_post(conn, **args2) is False
    row = conn.execute("SELECT source_listing FROM posts WHERE id=?", ("t3_abc",)).fetchone()
    assert row[0] == "hot"


def test_add_post_tickers_dedupes(conn):
    db.upsert_post(
        conn, post_id="t3_abc", title="t", author=None, body=None,
        created_utc=1, first_seen_utc=1, source_listing="hot",
        permalink=None, flair=None,
    )
    inserted_first = db.add_post_tickers(conn, "t3_abc", ["TSLA", "GME", "AAPL"])
    assert inserted_first == 3
    # Re-insert same set + one new — only the new one should be added.
    inserted_second = db.add_post_tickers(conn, "t3_abc", ["TSLA", "GME", "AAPL", "NVDA"])
    assert inserted_second == 1
    rows = conn.execute(
        "SELECT ticker FROM post_tickers WHERE post_id=? ORDER BY ticker",
        ("t3_abc",),
    ).fetchall()
    assert [r[0] for r in rows] == ["AAPL", "GME", "NVDA", "TSLA"]


def test_add_post_tickers_empty(conn):
    db.upsert_post(
        conn, post_id="t3_abc", title="t", author=None, body=None,
        created_utc=1, first_seen_utc=1, source_listing="hot",
        permalink=None, flair=None,
    )
    assert db.add_post_tickers(conn, "t3_abc", []) == 0


def test_upvote_snapshot_replace_on_collision(conn):
    db.upsert_post(
        conn, post_id="t3_abc", title="t", author=None, body=None,
        created_utc=1, first_seen_utc=1, source_listing="hot",
        permalink=None, flair=None,
    )
    db.record_upvote_snapshot(
        conn, post_id="t3_abc", snapshot_utc=1000,
        score=100, num_comments=10, upvote_ratio=0.9,
    )
    db.record_upvote_snapshot(
        conn, post_id="t3_abc", snapshot_utc=1000,
        score=150, num_comments=20, upvote_ratio=0.95,
    )
    rows = conn.execute(
        "SELECT score, num_comments, upvote_ratio FROM upvote_snapshots WHERE post_id=? AND snapshot_utc=?",
        ("t3_abc", 1000),
    ).fetchall()
    # Replace (not duplicate): exactly one row, with the second value.
    assert len(rows) == 1
    assert rows[0] == (150, 20, 0.95)


def test_upvote_snapshots_accumulate_across_runs(conn):
    """Per spec clarification A: each run records a new snapshot at its own ts."""
    db.upsert_post(
        conn, post_id="t3_abc", title="t", author=None, body=None,
        created_utc=1, first_seen_utc=1000, source_listing="hot",
        permalink=None, flair=None,
    )
    for ts, score in [(1000, 10), (3000, 25), (5000, 40)]:
        db.record_upvote_snapshot(
            conn, post_id="t3_abc", snapshot_utc=ts,
            score=score, num_comments=score // 2, upvote_ratio=0.8,
        )
    history = db.get_post_history(conn, "t3_abc")
    assert [(h[0], h[1]) for h in history] == [(1000, 10), (3000, 25), (5000, 40)]


def test_price_snapshot_replace_on_collision(conn):
    db.record_price_snapshot(
        conn, ticker="TSLA", snapshot_utc=1000,
        open_=100.0, high=110.0, low=99.0, close=105.0, volume=12345,
    )
    db.record_price_snapshot(
        conn, ticker="TSLA", snapshot_utc=1000,
        open_=101.0, high=111.0, low=100.0, close=106.0, volume=22222,
    )
    rows = conn.execute(
        "SELECT close, volume, source FROM price_snapshots WHERE ticker=? AND snapshot_utc=?",
        ("TSLA", 1000),
    ).fetchall()
    assert rows == [(106.0, 22222, "yfinance")]


def test_get_active_tickers_respects_window(conn):
    now = int(time.time())
    # post_a: seen 2 days ago
    db.upsert_post(
        conn, post_id="t3_a", title="t", author=None, body=None,
        created_utc=now - 86400 * 2, first_seen_utc=now - 86400 * 2,
        source_listing="hot", permalink=None, flair=None,
    )
    db.add_post_tickers(conn, "t3_a", ["TSLA"])
    # post_b: seen 30 days ago — outside default 7d window
    db.upsert_post(
        conn, post_id="t3_b", title="t", author=None, body=None,
        created_utc=now - 86400 * 30, first_seen_utc=now - 86400 * 30,
        source_listing="hot", permalink=None, flair=None,
    )
    db.add_post_tickers(conn, "t3_b", ["GME"])

    assert db.get_active_tickers(conn, since_days=7) == ["TSLA"]
    assert sorted(db.get_active_tickers(conn, since_days=60)) == ["GME", "TSLA"]


def test_multi_post_per_ticker(conn):
    """Per spec clarification B: two posts mentioning the same ticker => two
    independent rows in posts, both mapped via post_tickers, each with its own
    upvote time series. No dedup at ingest."""
    for pid in ("t3_p1", "t3_p2"):
        db.upsert_post(
            conn, post_id=pid, title=f"about GME {pid}", author=None, body=None,
            created_utc=1, first_seen_utc=1, source_listing="hot",
            permalink=None, flair=None,
        )
        db.add_post_tickers(conn, pid, ["GME"])
        db.record_upvote_snapshot(
            conn, post_id=pid, snapshot_utc=1000,
            score=50, num_comments=5, upvote_ratio=0.9,
        )
    rows = conn.execute(
        "SELECT post_id FROM post_tickers WHERE ticker='GME' ORDER BY post_id"
    ).fetchall()
    assert [r[0] for r in rows] == ["t3_p1", "t3_p2"]
    snap_count = conn.execute(
        "SELECT COUNT(*) FROM upvote_snapshots WHERE snapshot_utc=1000"
    ).fetchone()[0]
    assert snap_count == 2


def test_fundamentals_table_created(conn):
    cols = {r[1] for r in conn.execute("PRAGMA table_info(ticker_fundamentals)").fetchall()}
    assert {
        "ticker", "snapshot_date", "shares_outstanding", "shares_short",
        "float_shares", "short_ratio", "short_pct_float",
        "float_pct_outstanding", "held_pct_institutions",
        "held_pct_insiders", "held_pct_retail", "source",
    } == cols


def test_fundamentals_ownership_round_trip(conn):
    db.record_fundamentals(conn, "TSLA", 1700000000, {
        "shares_outstanding": 3_180_000_000,
        "held_pct_institutions": 0.45,
        "held_pct_insiders": 0.13,
        "held_pct_retail": 0.42,
    })
    row = conn.execute(
        "SELECT held_pct_institutions, held_pct_insiders, held_pct_retail "
        "FROM ticker_fundamentals WHERE ticker='TSLA'"
    ).fetchone()
    assert row == (0.45, 0.13, 0.42)


def test_init_db_backfills_columns_on_legacy_table(tmp_path):
    """Simulate a pre-existing DB that has the old fundamentals schema and confirm
    init_db ALTERs in the new ownership columns without losing data."""
    p = tmp_path / "legacy.db"
    legacy = sqlite3.connect(str(p))
    legacy.executescript("""
        CREATE TABLE ticker_fundamentals (
          ticker                TEXT NOT NULL,
          snapshot_date         INTEGER NOT NULL,
          shares_outstanding    INTEGER,
          shares_short          INTEGER,
          float_shares          INTEGER,
          short_ratio           REAL,
          short_pct_float       REAL,
          float_pct_outstanding REAL,
          source                TEXT NOT NULL DEFAULT 'yfinance',
          PRIMARY KEY (ticker, snapshot_date)
        );
    """)
    legacy.execute(
        "INSERT INTO ticker_fundamentals (ticker, snapshot_date, shares_outstanding) VALUES (?,?,?)",
        ("OLD", 1700000000, 12345),
    )
    legacy.commit()
    legacy.close()

    # init_db should add the missing columns and preserve the existing row.
    c = db.init_db(p)
    try:
        cols = {r[1] for r in c.execute("PRAGMA table_info(ticker_fundamentals)").fetchall()}
        assert "held_pct_institutions" in cols
        assert "held_pct_insiders" in cols
        assert "held_pct_retail" in cols
        # Pre-existing row still there with original data, new columns NULL.
        row = c.execute(
            "SELECT shares_outstanding, held_pct_institutions, held_pct_insiders, held_pct_retail "
            "FROM ticker_fundamentals WHERE ticker='OLD'"
        ).fetchone()
        assert row == (12345, None, None, None)
        # Re-running init_db on a now-current DB must not error.
        c.close()
        db.init_db(p).close()
    finally:
        try:
            c.close()
        except Exception:
            pass


def test_record_fundamentals_full_row(conn):
    db.record_fundamentals(conn, "TSLA", 1700000000, {
        "shares_outstanding": 3_180_000_000,
        "shares_short": 90_000_000,
        "float_shares": 2_700_000_000,
        "short_ratio": 1.4,
        "short_pct_float": 0.0333,
        "float_pct_outstanding": 0.849,
    })
    row = conn.execute(
        "SELECT shares_outstanding, shares_short, float_shares, short_ratio, "
        "short_pct_float, float_pct_outstanding, source "
        "FROM ticker_fundamentals WHERE ticker='TSLA'"
    ).fetchone()
    assert row == (3_180_000_000, 90_000_000, 2_700_000_000, 1.4, 0.0333, 0.849, "yfinance")


def test_record_fundamentals_partial_nulls(conn):
    db.record_fundamentals(conn, "FOO", 1700000000, {
        "shares_outstanding": 100_000_000,
        # everything else missing
    })
    row = conn.execute(
        "SELECT shares_outstanding, shares_short, float_shares, short_ratio, "
        "short_pct_float, float_pct_outstanding "
        "FROM ticker_fundamentals WHERE ticker='FOO'"
    ).fetchone()
    assert row == (100_000_000, None, None, None, None, None)


def test_record_fundamentals_replace_same_day(conn):
    db.record_fundamentals(conn, "GME", 1700000000, {"shares_outstanding": 380_000_000})
    db.record_fundamentals(conn, "GME", 1700000000, {"shares_outstanding": 381_000_000})
    n = conn.execute("SELECT COUNT(*) FROM ticker_fundamentals WHERE ticker='GME'").fetchone()[0]
    assert n == 1
    val = conn.execute(
        "SELECT shares_outstanding FROM ticker_fundamentals WHERE ticker='GME'"
    ).fetchone()[0]
    assert val == 381_000_000


def test_record_fundamentals_new_day_appends(conn):
    db.record_fundamentals(conn, "GME", 1700000000, {"shares_outstanding": 380_000_000})
    db.record_fundamentals(conn, "GME", 1700086400, {"shares_outstanding": 380_500_000})
    rows = conn.execute(
        "SELECT snapshot_date, shares_outstanding FROM ticker_fundamentals "
        "WHERE ticker='GME' ORDER BY snapshot_date"
    ).fetchall()
    assert rows == [(1700000000, 380_000_000), (1700086400, 380_500_000)]


def test_get_recent_post_ids_window(conn):
    now = int(time.time())
    cases = [
        ("t3_now", now - 3600),         # 1h old
        ("t3_24h", now - 86400),        # 1d old
        ("t3_60h", now - 60 * 3600),    # 60h old
        ("t3_old", now - 5 * 86400),    # 5d old
    ]
    for pid, created in cases:
        db.upsert_post(
            conn, post_id=pid, title="t", author=None, body=None,
            created_utc=created, first_seen_utc=created,
            source_listing="hot", permalink=None, flair=None,
        )
    # 72h window includes everything except t3_old
    assert sorted(db.get_recent_post_ids(conn, since_hours=72)) == ["t3_24h", "t3_60h", "t3_now"]
    # 6h window includes only the freshest
    assert db.get_recent_post_ids(conn, since_hours=6) == ["t3_now"]


def test_full_pass_is_idempotent(tmp_path):
    """Running the same insert sequence twice produces the same row counts."""
    p = tmp_path / "wsb.db"

    def one_pass():
        c = db.init_db(p)
        try:
            db.upsert_post(
                c, post_id="t3_abc", title="t", author=None, body=None,
                created_utc=1, first_seen_utc=1000, source_listing="hot",
                permalink=None, flair=None,
            )
            db.add_post_tickers(c, "t3_abc", ["TSLA", "GME"])
            db.record_upvote_snapshot(
                c, post_id="t3_abc", snapshot_utc=2000,
                score=100, num_comments=10, upvote_ratio=0.9,
            )
            c.commit()
        finally:
            c.close()

    one_pass()
    one_pass()  # second run must not duplicate

    c = sqlite3.connect(str(p))
    try:
        assert c.execute("SELECT COUNT(*) FROM posts").fetchone()[0] == 1
        assert c.execute("SELECT COUNT(*) FROM post_tickers").fetchone()[0] == 2
        assert c.execute("SELECT COUNT(*) FROM upvote_snapshots").fetchone()[0] == 1
    finally:
        c.close()
