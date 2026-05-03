"""Fetch yfinance 15-minute bars for tickers seen in the last N days.

Run as:
    python -m social.wsb_momentum.price_fetcher

Schedule on the same 30-minute cadence as the collector. Idempotent
(price_snapshots replaces on (ticker, snapshot_utc) collision).
"""

from __future__ import annotations

import time
import traceback
from datetime import datetime, timezone

import yfinance as yf

from . import db
from .config import DB_PATH

ACTIVE_WINDOW_DAYS = 7
YF_PERIOD = "2d"
YF_INTERVAL = "15m"


def _bar_to_kwargs(ticker: str, ts, row) -> dict:
    """Convert one yfinance bar (pandas Timestamp + row Series) to record_price_snapshot kwargs."""
    snapshot_utc = int(ts.timestamp())
    return dict(
        ticker=ticker,
        snapshot_utc=snapshot_utc,
        open_=_float(row.get("Open")),
        high=_float(row.get("High")),
        low=_float(row.get("Low")),
        close=_float(row.get("Close")),
        volume=_int(row.get("Volume")),
    )


def _float(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if f != f:  # NaN
        return None
    return f


def _int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def fetch_for_ticker(conn, ticker: str) -> int:
    """Fetch bars for one ticker and write them. Returns rows written, or -1 on failure."""
    try:
        hist = yf.Ticker(ticker).history(period=YF_PERIOD, interval=YF_INTERVAL, auto_adjust=False)
    except Exception as e:
        print(f"  {ticker}: yfinance error: {e!r}")
        return -1

    if hist is None or hist.empty:
        print(f"  {ticker}: no data returned")
        return 0

    written = 0
    for ts, row in hist.iterrows():
        kwargs = _bar_to_kwargs(ticker, ts, row)
        if kwargs["close"] is None and kwargs["open_"] is None:
            continue
        try:
            db.record_price_snapshot(conn, **kwargs)
            written += 1
        except Exception as e:
            print(f"  {ticker} @ {ts}: write failed: {e!r}")
    return written


def run_once() -> None:
    snapshot_utc = int(time.time())
    iso = datetime.fromtimestamp(snapshot_utc, tz=timezone.utc).isoformat()
    print(f"[price-fetch @ {iso}] starting")

    conn = db.init_db(DB_PATH)
    try:
        tickers = db.get_active_tickers(conn, since_days=ACTIVE_WINDOW_DAYS)
        print(f"  active tickers (last {ACTIVE_WINDOW_DAYS}d): {len(tickers)}")

        total_written = 0
        ok = 0
        empty = 0
        failed = 0

        for t in tickers:
            n = fetch_for_ticker(conn, t)
            if n > 0:
                ok += 1
                total_written += n
            elif n == 0:
                empty += 1
            else:
                failed += 1
            conn.commit()

        print(
            f"[price-fetch @ {iso}] done: {ok} ok, {empty} empty, {failed} failed; "
            f"{total_written} bar rows written"
        )
    except Exception:
        traceback.print_exc()
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    run_once()
