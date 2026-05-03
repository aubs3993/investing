"""Tests for the ticker family mapping + the auto-generated SQL view."""

from __future__ import annotations

import sqlite3

import pytest

from social.wsb_momentum import db, ticker_families


def test_dict_shape_value_is_pair_of_str_int():
    for raw, value in ticker_families.TICKER_FAMILY.items():
        assert isinstance(raw, str) and raw.isupper()
        assert isinstance(value, tuple) and len(value) == 2
        underlying, leverage = value
        assert isinstance(underlying, str) and underlying.isupper()
        assert isinstance(leverage, int)
        assert leverage != 0  # 0x leverage is meaningless


def test_known_mappings_present():
    fam = ticker_families.TICKER_FAMILY
    assert fam["AMDL"] == ("AMD", 2)
    assert fam["NVDL"] == ("NVDA", 2)
    assert fam["TSLL"] == ("TSLA", 2)
    assert fam["TQQQ"] == ("QQQ", 3)
    assert fam["SQQQ"] == ("QQQ", -3)
    assert fam["SOXL"] == ("SOXX", 3)
    assert fam["SOXS"] == ("SOXX", -3)
    assert fam["UPRO"] == ("SPY", 3)


def test_view_sql_includes_each_mapping():
    sql = ticker_families.build_unified_view_sql()
    for raw, (underlying, leverage) in ticker_families.TICKER_FAMILY.items():
        assert f"WHEN '{raw}' THEN '{underlying}'" in sql
        assert f"WHEN '{raw}' THEN {leverage}" in sql
    assert "ELSE pt.ticker" in sql
    assert "ELSE 1" in sql
    assert "FROM post_tickers pt" in sql


def test_view_aggregates_amdl_into_amd(tmp_path):
    """End-to-end: populate post_tickers with AMD + AMDL + non-mapped ticker;
    confirm the view rolls AMDL up to AMD with leverage=2 and leaves AMD as-is."""
    c = db.init_db(tmp_path / "wsb.db")
    c.row_factory = sqlite3.Row
    try:
        for pid in ("p1", "p2", "p3"):
            db.upsert_post(
                c, post_id=pid, title="t", author=None, body=None,
                created_utc=1, first_seen_utc=1, source_listing="hot",
                permalink=None, flair=None,
            )
        db.add_post_tickers(c, "p1", ["AMD"])     # underlying long
        db.add_post_tickers(c, "p2", ["AMDL"])    # 2x leveraged AMD
        db.add_post_tickers(c, "p3", ["TSLA"])    # untouched fall-through

        rows = {r["raw_ticker"]: dict(r) for r in c.execute(
            "SELECT raw_ticker, underlying_ticker, leverage FROM ticker_attention_unified"
        ).fetchall()}
        assert rows["AMD"]  == {"raw_ticker": "AMD",  "underlying_ticker": "AMD",  "leverage": 1}
        assert rows["AMDL"] == {"raw_ticker": "AMDL", "underlying_ticker": "AMDL".replace("AMDL", "AMD"), "leverage": 2}
        assert rows["TSLA"] == {"raw_ticker": "TSLA", "underlying_ticker": "TSLA", "leverage": 1}

        # Aggregation by underlying_ticker should collapse AMD + AMDL into one row.
        agg = c.execute("""
            SELECT underlying_ticker, COUNT(*) AS n
            FROM ticker_attention_unified
            WHERE underlying_ticker = 'AMD'
        """).fetchone()
        assert agg["n"] == 2
    finally:
        c.close()


def test_view_inverse_etf_carries_negative_leverage(tmp_path):
    c = db.init_db(tmp_path / "wsb.db")
    c.row_factory = sqlite3.Row
    try:
        db.upsert_post(
            c, post_id="bear1", title="t", author=None, body=None,
            created_utc=1, first_seen_utc=1, source_listing="hot",
            permalink=None, flair=None,
        )
        db.add_post_tickers(c, "bear1", ["SQQQ"])

        row = c.execute(
            "SELECT raw_ticker, underlying_ticker, leverage "
            "FROM ticker_attention_unified WHERE raw_ticker='SQQQ'"
        ).fetchone()
        assert row["underlying_ticker"] == "QQQ"
        assert row["leverage"] == -3

        # Bullish-vs-bearish split query should classify this as bearish.
        row2 = c.execute("""
            SELECT
              CASE WHEN leverage > 0 THEN 'bull' ELSE 'bear' END AS side,
              underlying_ticker
            FROM ticker_attention_unified WHERE raw_ticker='SQQQ'
        """).fetchone()
        assert row2["side"] == "bear"
    finally:
        c.close()


def test_view_recreated_on_each_init_db_picks_up_dict_changes(tmp_path, monkeypatch):
    """If TICKER_FAMILY gains a new entry, the next init_db() must reflect it."""
    p = tmp_path / "wsb.db"

    # First init with a baseline mapping.
    monkeypatch.setattr(ticker_families, "TICKER_FAMILY", {"AAA": ("BBB", 2)})
    c = db.init_db(p)
    c.row_factory = sqlite3.Row
    db.upsert_post(c, post_id="x", title="t", author=None, body=None,
                   created_utc=1, first_seen_utc=1, source_listing="hot",
                   permalink=None, flair=None)
    db.add_post_tickers(c, "x", ["AAA", "CCC"])
    c.commit()
    rows = {r["raw_ticker"]: r["underlying_ticker"] for r in c.execute(
        "SELECT raw_ticker, underlying_ticker FROM ticker_attention_unified"
    ).fetchall()}
    assert rows == {"AAA": "BBB", "CCC": "CCC"}
    c.close()

    # Add a new mapping; re-init must pick it up.
    monkeypatch.setattr(ticker_families, "TICKER_FAMILY", {"AAA": ("BBB", 2), "CCC": ("DDD", -3)})
    c = db.init_db(p)
    c.row_factory = sqlite3.Row
    rows = {r["raw_ticker"]: (r["underlying_ticker"], r["leverage"]) for r in c.execute(
        "SELECT raw_ticker, underlying_ticker, leverage FROM ticker_attention_unified"
    ).fetchall()}
    assert rows == {"AAA": ("BBB", 2), "CCC": ("DDD", -3)}
    c.close()
