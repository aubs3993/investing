"""Extract stock tickers from WSB post text.

Two regex passes over the input:
  - cashtag: $TSLA  -> bypasses the blacklist (explicit author intent)
  - bare:    TSLA   -> must be in the NYSE/NASDAQ ticker set AND not blacklisted

Mixed-case tokens (e.g. "Tsla", "tsla") are intentionally ignored — the
all-caps convention is a strong signal vs. random capitalization.
"""

from __future__ import annotations

import csv
import re
from functools import lru_cache
from pathlib import Path

from .config import TICKER_BLACKLIST, TICKERS_CSV

CASHTAG_RE = re.compile(r"\$([A-Z]{1,5})\b")
BARE_RE = re.compile(r"(?<![A-Za-z0-9$])([A-Z]{1,5})(?![A-Za-z0-9])")


@lru_cache(maxsize=1)
def load_ticker_set(csv_path: Path = TICKERS_CSV) -> frozenset[str]:
    """Read the reference CSV and return the set of valid tickers (uppercase)."""
    with Path(csv_path).open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        return frozenset(row["ticker"].strip().upper() for row in reader if row.get("ticker"))


def extract_tickers(text: str) -> set[str]:
    """Return the set of valid tickers found in `text`.

    Cashtags ($TSLA) bypass the blacklist; bare matches (TSLA) do not.
    Both must appear in the loaded NYSE/NASDAQ ticker set.
    """
    if not text:
        return set()

    valid = load_ticker_set()
    found: set[str] = set()

    for sym in CASHTAG_RE.findall(text):
        if sym in valid:
            found.add(sym)

    for sym in BARE_RE.findall(text):
        if sym in valid and sym not in TICKER_BLACKLIST:
            found.add(sym)

    return found
