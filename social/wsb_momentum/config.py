"""Constants for the WSB momentum collector."""

from __future__ import annotations

import os
from pathlib import Path

PACKAGE_DIR = Path(__file__).resolve().parent
DATA_DIR = PACKAGE_DIR / "data"
OUTPUT_DIR = PACKAGE_DIR / "output"

DB_PATH = OUTPUT_DIR / "wsb.db"
TICKERS_CSV = DATA_DIR / "tickers_nyse_nasdaq.csv"

SUBREDDIT = "wallstreetbets"
LISTINGS = ("hot", "new", "rising", "top")
TOP_TIMEFRAME = "day"

# Reddit asks for a unique, descriptive UA. Override via env var if you want.
USER_AGENT = os.environ.get(
    "WSB_MOMENTUM_UA",
    "wsb-momentum-research/0.1 (by u/aubs3993)",
)

REDDIT_FETCH_LIMIT = 100
SLEEP_BETWEEN_LISTINGS_S = 1.0

# Refresh-pass: re-snapshot every post created within this many hours of the
# current run, even if it no longer surfaces in any listing. Reddit scores are
# effectively frozen after ~3 days, so 72h is the natural cutoff.
REFRESH_WINDOW_HOURS = 72
SLEEP_BETWEEN_POST_FETCHES_S = 0.75

# Fundamentals fetcher: how far back to consider a ticker "active" and how
# long to sleep between yfinance .info calls (Yahoo dislikes rapid bursts).
FUNDAMENTALS_ACTIVE_WINDOW_DAYS = 7
SLEEP_BETWEEN_FUNDAMENTALS_FETCHES_S = 0.75

# Words that look like tickers but aren't. Extend freely.
TICKER_BLACKLIST = frozenset({
    # WSB jargon
    "DD", "YOLO", "FOMO", "FUD", "FYI", "IMO", "IMHO", "TLDR", "TLDR",
    "ATH", "ATL", "EOD", "EOW", "EOM", "EOY", "WSB", "OP",
    # Options / market shorthand
    "IV", "OTM", "ITM", "ATM", "PR", "PE", "PT", "EPS", "EBIT",
    "EBITDA", "PEG", "ROE", "ROIC", "ROA", "FCF", "OPEX", "CAPEX",
    "BUY", "SELL", "HOLD", "LONG", "SHORT", "CALL", "CALLS", "PUT", "PUTS",
    # Acronyms / orgs
    "CEO", "CFO", "CTO", "COO", "CIO", "USA", "USD", "EU", "UK", "GDP",
    "ETF", "IPO", "SPAC", "SEC", "FDA", "FOMC", "FED", "FBI", "CIA", "DOJ",
    "DOD", "DOT", "DOE", "EPA", "IRS", "FTC", "FCC", "NYSE", "AMEX",
    # Crypto / tech (often false positives in WSB threads)
    "BTC", "ETH", "AI", "GPU", "CPU", "RAM", "SSD", "API", "URL", "CSV",
    "PDF", "HTML", "JSON", "ML", "LLM",
    # Generic English / pronouns / common verbs
    "A", "I", "AM", "AN", "AND", "AS", "AT", "BE", "BUT", "BY", "DO",
    "FOR", "GO", "HE", "IF", "IN", "IS", "IT", "ITS", "ME", "MY", "NO",
    "NOT", "OF", "ON", "OR", "OUR", "OUT", "PER", "SHE", "SO", "THE",
    "THIS", "TO", "UP", "US", "WE", "WHO", "WHY", "YES", "YOU", "YOUR",
    # Single-letter false positives observed in production: P from P/E, P/S,
    # P/FFO, S&P, P&L; S from S&P; E from P/E. Cashtags ($P, $S, $E) still
    # match because the cashtag pass bypasses the blacklist.
    "P", "S", "E",
    # Audit tool flagged (suspicion >= 4.0): common English words.
    "ARE", "ALL", "MAN", "B", "U",
    # Audit tool tier-2 (single-letter, no cashtag use). Excludes T (AT&T) and
    # V (Visa) — both have legitimate bare-ticker discussion on WSB.
    "Q", "D", "G", "L", "R",
    # Other common false positives observed on WSB
    "OK", "LOL", "LMAO", "ROFL", "WTF", "WTH", "OMG", "SMH", "TBH",
    "AF", "BS", "MF", "NSFW", "NSFL", "RIP", "GG", "EZ", "GL", "HF",
    "MOON", "BULL", "BEAR", "BAG", "BAGS", "TENDIES", "RETARD", "APE",
    "HODL", "FOMO", "DCA",
    "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT",
    "NOV", "DEC",
    "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN",
})
