"""SQLite schema + helper functions for the WSB momentum collector.

No ORM. All functions take an explicit connection (or a path that we open
ourselves) so callers control transaction scope. The collector is meant to
be re-run every 30 minutes; helpers are idempotent.
"""

from __future__ import annotations

import sqlite3
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Iterable

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS posts (
  id              TEXT PRIMARY KEY,
  title           TEXT NOT NULL,
  author          TEXT,
  body            TEXT,
  created_utc     INTEGER NOT NULL,
  first_seen_utc  INTEGER NOT NULL,
  source_listing  TEXT NOT NULL,
  permalink       TEXT,
  flair           TEXT
);

CREATE TABLE IF NOT EXISTS post_tickers (
  post_id  TEXT NOT NULL,
  ticker   TEXT NOT NULL,
  PRIMARY KEY (post_id, ticker),
  FOREIGN KEY (post_id) REFERENCES posts(id)
);

CREATE TABLE IF NOT EXISTS upvote_snapshots (
  post_id        TEXT NOT NULL,
  snapshot_utc   INTEGER NOT NULL,
  score          INTEGER NOT NULL,
  num_comments   INTEGER NOT NULL,
  upvote_ratio   REAL,
  PRIMARY KEY (post_id, snapshot_utc),
  FOREIGN KEY (post_id) REFERENCES posts(id)
);

CREATE TABLE IF NOT EXISTS price_snapshots (
  ticker         TEXT NOT NULL,
  snapshot_utc   INTEGER NOT NULL,
  open           REAL,
  high           REAL,
  low            REAL,
  close          REAL,
  volume         INTEGER,
  source         TEXT NOT NULL DEFAULT 'yfinance',
  PRIMARY KEY (ticker, snapshot_utc)
);

CREATE TABLE IF NOT EXISTS ticker_fundamentals (
  ticker                 TEXT NOT NULL,
  snapshot_date          INTEGER NOT NULL,   -- unix midnight UTC of the date pulled
  shares_outstanding     INTEGER,
  shares_short           INTEGER,
  float_shares           INTEGER,
  short_ratio            REAL,               -- days to cover, from yfinance
  short_pct_float        REAL,               -- computed: shares_short / float_shares
  float_pct_outstanding  REAL,               -- computed: float_shares / shares_outstanding
  held_pct_institutions  REAL,               -- from yfinance: heldPercentInstitutions
  held_pct_insiders      REAL,               -- from yfinance: heldPercentInsiders
  held_pct_retail        REAL,               -- computed: 1 - institutions - insiders (NULL if would be <0)
  source                 TEXT NOT NULL DEFAULT 'yfinance',
  PRIMARY KEY (ticker, snapshot_date)
);

CREATE INDEX IF NOT EXISTS idx_upvote_post ON upvote_snapshots(post_id);
CREATE INDEX IF NOT EXISTS idx_upvote_time ON upvote_snapshots(snapshot_utc);
CREATE INDEX IF NOT EXISTS idx_price_ticker_time ON price_snapshots(ticker, snapshot_utc);
CREATE INDEX IF NOT EXISTS idx_post_tickers_ticker ON post_tickers(ticker);
CREATE INDEX IF NOT EXISTS idx_fundamentals_ticker_date ON ticker_fundamentals(ticker, snapshot_date);
"""


# Columns we may need to add to pre-existing tables. CREATE TABLE IF NOT EXISTS
# only handles whole tables; for column additions on existing DBs we ALTER.
# Keep this list aligned with the SCHEMA_SQL above for any column added after
# the initial release.
_BACKFILL_COLUMNS = {
    "ticker_fundamentals": (
        ("held_pct_institutions", "REAL"),
        ("held_pct_insiders",     "REAL"),
        ("held_pct_retail",       "REAL"),
    ),
}


def _backfill_columns(conn: sqlite3.Connection) -> None:
    """ALTER TABLE ADD COLUMN for any column listed in _BACKFILL_COLUMNS that
    is missing from the live schema. Idempotent: skips columns already present.
    """
    for table, cols in _BACKFILL_COLUMNS.items():
        existing = {r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()}
        if not existing:
            # Table doesn't exist yet — SCHEMA_SQL above will have created it
            # before we got here, so this branch shouldn't fire. Defensive.
            continue
        for name, sql_type in cols:
            if name not in existing:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {sql_type}")


def _refresh_views(conn: sqlite3.Connection) -> None:
    """Drop + recreate analytical SQL views.

    Views are rebuilt from in-Python source (e.g. ticker_families.py) on every
    init_db() call so dict edits are picked up automatically without a separate
    migration step.
    """
    # Imported here to avoid a top-level cycle: db.py is imported by every
    # entry point, and ticker_families.py is purely data with no deps.
    from .ticker_families import build_unified_view_sql

    conn.execute("DROP VIEW IF EXISTS ticker_attention_unified")
    conn.execute(build_unified_view_sql("ticker_attention_unified"))


def init_db(path: str | Path) -> sqlite3.Connection:
    """Open (creating if needed) the DB at `path` and ensure the schema exists.

    Safe to call on every collector run. Also adds any new columns missing
    from a pre-existing DB (see _BACKFILL_COLUMNS) and rebuilds analytical
    views (see _refresh_views).
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript(SCHEMA_SQL)
    _backfill_columns(conn)
    _refresh_views(conn)
    conn.commit()
    return conn


@contextmanager
def connect(path: str | Path):
    """Context manager wrapper around init_db."""
    conn = init_db(path)
    try:
        yield conn
    finally:
        conn.close()


def upsert_post(
    conn: sqlite3.Connection,
    *,
    post_id: str,
    title: str,
    author: str | None,
    body: str | None,
    created_utc: int,
    first_seen_utc: int,
    source_listing: str,
    permalink: str | None,
    flair: str | None,
) -> bool:
    """Insert a post if new. Returns True if inserted, False if it already existed.

    Existing posts are NOT modified — `source_listing` is write-once and locked
    to whichever listing first surfaced the post on its first collector run.
    Subsequent runs that re-encounter the post in a different listing do not
    overwrite it; they only record a new upvote snapshot.
    """
    cur = conn.execute(
        """
        INSERT OR IGNORE INTO posts
            (id, title, author, body, created_utc, first_seen_utc,
             source_listing, permalink, flair)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (post_id, title, author, body, created_utc, first_seen_utc,
         source_listing, permalink, flair),
    )
    return cur.rowcount == 1


def add_post_tickers(conn: sqlite3.Connection, post_id: str, tickers: Iterable[str]) -> int:
    """Insert (post_id, ticker) pairs, ignoring duplicates. Returns rows inserted."""
    rows = [(post_id, t) for t in tickers]
    if not rows:
        return 0
    cur = conn.executemany(
        "INSERT OR IGNORE INTO post_tickers (post_id, ticker) VALUES (?, ?)",
        rows,
    )
    return cur.rowcount or 0


def record_upvote_snapshot(
    conn: sqlite3.Connection,
    *,
    post_id: str,
    snapshot_utc: int,
    score: int,
    num_comments: int,
    upvote_ratio: float | None,
) -> None:
    """Write an upvote snapshot. Replaces on (post_id, snapshot_utc) collision."""
    conn.execute(
        """
        INSERT OR REPLACE INTO upvote_snapshots
            (post_id, snapshot_utc, score, num_comments, upvote_ratio)
        VALUES (?, ?, ?, ?, ?)
        """,
        (post_id, snapshot_utc, score, num_comments, upvote_ratio),
    )


def record_price_snapshot(
    conn: sqlite3.Connection,
    *,
    ticker: str,
    snapshot_utc: int,
    open_: float | None,
    high: float | None,
    low: float | None,
    close: float | None,
    volume: int | None,
    source: str = "yfinance",
) -> None:
    """Write a price bar. Replaces on (ticker, snapshot_utc) collision."""
    conn.execute(
        """
        INSERT OR REPLACE INTO price_snapshots
            (ticker, snapshot_utc, open, high, low, close, volume, source)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (ticker, snapshot_utc, open_, high, low, close, volume, source),
    )


FUNDAMENTALS_FIELDS = (
    "shares_outstanding",
    "shares_short",
    "float_shares",
    "short_ratio",
    "short_pct_float",
    "float_pct_outstanding",
    "held_pct_institutions",
    "held_pct_insiders",
    "held_pct_retail",
)


def record_fundamentals(
    conn: sqlite3.Connection,
    ticker: str,
    snapshot_date: int,
    fields: dict,
    *,
    source: str = "yfinance",
) -> None:
    """Write one daily fundamentals row. INSERT OR REPLACE on (ticker, date).

    `fields` may carry any subset of FUNDAMENTALS_FIELDS; absent keys are
    stored as NULL. Extra keys are ignored.
    """
    values = [ticker, snapshot_date]
    values.extend(fields.get(f) for f in FUNDAMENTALS_FIELDS)
    values.append(source)
    cols = "ticker, snapshot_date, " + ", ".join(FUNDAMENTALS_FIELDS) + ", source"
    placeholders = ", ".join(["?"] * len(values))
    conn.execute(
        f"INSERT OR REPLACE INTO ticker_fundamentals ({cols}) VALUES ({placeholders})",
        values,
    )


def get_active_tickers(conn: sqlite3.Connection, since_days: int = 7) -> list[str]:
    """Distinct tickers from posts whose first_seen_utc is within `since_days`."""
    cutoff = int(time.time()) - since_days * 86400
    rows = conn.execute(
        """
        SELECT DISTINCT pt.ticker
        FROM post_tickers pt
        JOIN posts p ON p.id = pt.post_id
        WHERE p.first_seen_utc >= ?
        ORDER BY pt.ticker
        """,
        (cutoff,),
    ).fetchall()
    return [r[0] for r in rows]


def get_recent_post_ids(conn: sqlite3.Connection, since_hours: int) -> list[str]:
    """Post ids whose `created_utc` is within the last `since_hours` hours.

    Used by the collector's refresh pass to re-snapshot posts that may no
    longer surface in any listing but whose scores are still moving.
    """
    cutoff = int(time.time()) - since_hours * 3600
    rows = conn.execute(
        "SELECT id FROM posts WHERE created_utc >= ? ORDER BY created_utc",
        (cutoff,),
    ).fetchall()
    return [r[0] for r in rows]


def get_post_history(conn: sqlite3.Connection, post_id: str) -> list[tuple]:
    """Return the upvote-snapshot time series for a given post, oldest first."""
    return conn.execute(
        """
        SELECT snapshot_utc, score, num_comments, upvote_ratio
        FROM upvote_snapshots
        WHERE post_id = ?
        ORDER BY snapshot_utc
        """,
        (post_id,),
    ).fetchall()
