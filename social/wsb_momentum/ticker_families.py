"""Mapping of leveraged / inverse stock ETFs to their underlying.

Used by analysis-layer queries — NOT by the ingest path. The raw
`post_tickers` table continues to capture tickers exactly as written
(AMDL, TQQQ, SOXL stay distinct rows). The `ticker_attention_unified`
SQL view, which db.init_db() rebuilds on every run, joins this dict
into a fall-through CASE expression so analysis queries can either:

    - aggregate by `underlying_ticker` to roll AMDL up to AMD, or
    - filter on `leverage` to split bullish (>0) vs bearish (<0)
      sentiment, or
    - drill back via `raw_ticker` to AMDL specifically.

Crypto ETFs (BITO, IBIT, FBTC, ETHA, etc.) are intentionally absent —
they don't roll up to a stock and the project is about stock momentum.
They simply fall through ELSE in the view (raw_ticker = underlying_ticker,
leverage = 1).

Edit freely. After editing, re-run any script that calls db.init_db()
(collector, price_fetcher, fundamentals_fetcher) to refresh the view.
"""

from __future__ import annotations

# raw ticker -> (underlying ticker, leverage multiplier)
# Convention: positive = long the underlying, negative = inverse.
# Unleveraged (1x or -1x) inverses are included where they're WSB-relevant.
TICKER_FAMILY: dict[str, tuple[str, int]] = {
    # ---- Single-stock leveraged longs ----
    "AMDL": ("AMD",   2),
    "NVDL": ("NVDA",  2),
    "TSLL": ("TSLA",  2),
    "MSFU": ("MSFT",  2),
    "METU": ("META",  2),
    "AMZU": ("AMZN",  2),
    "GOOL": ("GOOGL", 2),
    "MSTU": ("MSTR",  2),

    # ---- Single-stock inverse ----
    "TSLZ": ("TSLA", -2),
    "NVD":  ("NVDA", -2),
    "MSTZ": ("MSTR", -2),

    # ---- Broad index leveraged ----
    "TQQQ": ("QQQ",   3),
    "SQQQ": ("QQQ",  -3),
    "QLD":  ("QQQ",   2),
    "PSQ":  ("QQQ",  -1),
    "SPXL": ("SPY",   3),
    "SPXS": ("SPY",  -3),
    "UPRO": ("SPY",   3),
    "SH":   ("SPY",  -1),
    "SSO":  ("SPY",   2),
    "SDS":  ("SPY",  -2),

    # ---- Sector leveraged ----
    "SOXL": ("SOXX",  3),
    "SOXS": ("SOXX", -3),
    "FAS":  ("XLF",   3),
    "FAZ":  ("XLF",  -3),
    "TNA":  ("IWM",   3),
    "TZA":  ("IWM",  -3),
}


def build_unified_view_sql(view_name: str = "ticker_attention_unified") -> str:
    """Generate `CREATE VIEW IF NOT EXISTS …` SQL for the unified attention view.

    The view exposes three columns per row in post_tickers:
      - post_id          (passthrough)
      - raw_ticker       (passthrough — what was actually extracted)
      - underlying_ticker (mapped if raw_ticker is in TICKER_FAMILY, else raw_ticker)
      - leverage         (multiplier from TICKER_FAMILY, else 1)
    """
    if not TICKER_FAMILY:
        # Pathological case — degenerate to a passthrough view.
        case_underlying = "pt.ticker"
        case_leverage = "1"
    else:
        # Build the CASE expressions deterministically (sorted by raw ticker)
        # so the generated SQL is stable across runs.
        items = sorted(TICKER_FAMILY.items())
        underlying_when = "\n".join(
            f"        WHEN '{raw}' THEN '{underlying}'"
            for raw, (underlying, _lev) in items
        )
        leverage_when = "\n".join(
            f"        WHEN '{raw}' THEN {lev}"
            for raw, (_underlying, lev) in items
        )
        case_underlying = (
            "CASE pt.ticker\n"
            f"{underlying_when}\n"
            "        ELSE pt.ticker\n"
            "    END"
        )
        case_leverage = (
            "CASE pt.ticker\n"
            f"{leverage_when}\n"
            "        ELSE 1\n"
            "    END"
        )

    return (
        f"CREATE VIEW IF NOT EXISTS {view_name} AS\n"
        f"SELECT\n"
        f"    pt.post_id,\n"
        f"    pt.ticker AS raw_ticker,\n"
        f"    {case_underlying} AS underlying_ticker,\n"
        f"    {case_leverage} AS leverage\n"
        f"FROM post_tickers pt;"
    )
