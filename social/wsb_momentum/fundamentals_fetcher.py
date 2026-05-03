"""Fetch daily fundamentals (shares outstanding, short interest, float) per active ticker.

Run as:
    python -m social.wsb_momentum.fundamentals_fetcher

Schedule once a day (the cadence-mismatch with the 30-min collector is by
design — short interest only updates twice a month with a ~2 week lag, but
float and shares outstanding can move on corporate actions, so daily is the
natural cadence).

Pulls from `yfinance.Ticker(t).info`. Idempotent: same-day re-runs replace;
new-day runs append a new row per ticker.
"""

from __future__ import annotations

import time
from datetime import datetime, timezone

import yfinance as yf

from . import db
from .config import (
    DB_PATH,
    FUNDAMENTALS_ACTIVE_WINDOW_DAYS,
    SLEEP_BETWEEN_FUNDAMENTALS_FETCHES_S,
)


# yfinance .info -> our column names
INFO_KEY_MAP = {
    "shares_outstanding":    "sharesOutstanding",
    "shares_short":          "sharesShort",
    "float_shares":          "floatShares",
    "short_ratio":           "shortRatio",
    "held_pct_institutions": "heldPercentInstitutions",
    "held_pct_insiders":     "heldPercentInsiders",
}

# Fields used in the "full vs partial vs skipped" classification. Derived
# fields (short_pct_float, float_pct_outstanding, held_pct_retail) are not
# included — they're computed, not pulled.
RAW_FIELDS_FOR_CLASSIFY = (
    "shares_outstanding",
    "shares_short",
    "float_shares",
    "short_ratio",
    "held_pct_institutions",
    "held_pct_insiders",
)


def _today_midnight_utc() -> int:
    """Today's unix midnight in UTC."""
    now = datetime.now(timezone.utc)
    midnight = datetime(now.year, now.month, now.day, tzinfo=timezone.utc)
    return int(midnight.timestamp())


def _coerce_int(v):
    if v is None:
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if f != f:  # NaN
        return None
    return int(f)


def _coerce_float(v):
    if v is None:
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if f != f:
        return None
    return f


def fetch_for_ticker(ticker: str) -> dict | None:
    """Pull yfinance .info and return a fields dict (with derived ratios), or None on failure."""
    try:
        info = yf.Ticker(ticker).info
    except Exception as e:
        print(f"  {ticker}: yfinance .info error: {e!r}")
        return None

    if not info or not isinstance(info, dict):
        return None

    fields: dict = {
        "shares_outstanding":    _coerce_int(info.get(INFO_KEY_MAP["shares_outstanding"])),
        "shares_short":          _coerce_int(info.get(INFO_KEY_MAP["shares_short"])),
        "float_shares":          _coerce_int(info.get(INFO_KEY_MAP["float_shares"])),
        "short_ratio":           _coerce_float(info.get(INFO_KEY_MAP["short_ratio"])),
        "held_pct_institutions": _coerce_float(info.get(INFO_KEY_MAP["held_pct_institutions"])),
        "held_pct_insiders":     _coerce_float(info.get(INFO_KEY_MAP["held_pct_insiders"])),
    }

    # Derived ratios — null-safe.
    sf = fields["shares_short"]
    fl = fields["float_shares"]
    so = fields["shares_outstanding"]
    fields["short_pct_float"] = (sf / fl) if (sf and fl and fl > 0) else None
    fields["float_pct_outstanding"] = (fl / so) if (fl and so and so > 0) else None

    inst = fields["held_pct_institutions"]
    ins = fields["held_pct_insiders"]
    if inst is not None and ins is not None:
        retail = 1.0 - inst - ins
        # Tolerance guards against float noise (e.g. 1 - 0.8 - 0.2 == -1.1e-16).
        # Only treat *meaningfully* negative results as a data inconsistency.
        if retail < -1e-6:
            # Yahoo's data can be inconsistent (esp. dual-listed / foreign issues
            # where institutions+insiders > 100%). Clamp to NULL rather than
            # storing a negative ownership %.
            print(f"  {ticker}: computed held_pct_retail={retail:.4f} < 0 "
                  f"(inst={inst:.4f}, ins={ins:.4f}) — clamping to NULL")
            fields["held_pct_retail"] = None
        else:
            fields["held_pct_retail"] = max(retail, 0.0)
    else:
        fields["held_pct_retail"] = None

    return fields


def _classify(fields: dict) -> str:
    """Categorize how complete the row is for the summary log."""
    if fields is None:
        return "skipped"
    present = sum(1 for k in RAW_FIELDS_FOR_CLASSIFY if fields.get(k) is not None)
    if present == 0:
        return "skipped"
    if present == len(RAW_FIELDS_FOR_CLASSIFY):
        return "full"
    return "partial"


def run_once() -> None:
    snapshot_date = _today_midnight_utc()
    iso = datetime.fromtimestamp(snapshot_date, tz=timezone.utc).date().isoformat()
    print(f"[fundamentals @ {iso}] starting")

    conn = db.init_db(DB_PATH)
    try:
        tickers = db.get_active_tickers(conn, since_days=FUNDAMENTALS_ACTIVE_WINDOW_DAYS)
        print(f"  active tickers (last {FUNDAMENTALS_ACTIVE_WINDOW_DAYS}d): {len(tickers)}")

        n_full = n_partial = n_skip = 0
        for t in tickers:
            fields = fetch_for_ticker(t)
            kind = _classify(fields)
            if kind == "skipped":
                if fields is None:
                    print(f"  {t}: no .info returned, skipping")
                else:
                    print(f"  {t}: all key fields missing, skipping")
                n_skip += 1
            else:
                try:
                    db.record_fundamentals(conn, t, snapshot_date, fields)
                    if kind == "full":
                        n_full += 1
                    else:
                        n_partial += 1
                        missing = [k for k in RAW_FIELDS_FOR_CLASSIFY if fields.get(k) is None]
                        print(f"  {t}: partial (missing: {', '.join(missing)})")
                except Exception as e:
                    print(f"  {t}: write failed: {e!r}")
                    n_skip += 1
            conn.commit()
            time.sleep(SLEEP_BETWEEN_FUNDAMENTALS_FETCHES_S)

        n_updated = n_full + n_partial
        print(
            f"[fundamentals @ {iso}] updated {n_updated} tickers, "
            f"{n_full} with full data, {n_partial} with partial data, "
            f"{n_skip} skipped (no data)"
        )
    finally:
        conn.close()


if __name__ == "__main__":
    run_once()
