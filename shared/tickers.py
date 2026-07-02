"""Ticker validation and filesystem-safe naming, shared by all per-ticker scripts.

The raw ticker (e.g. 700:HK) is what CapIQ formulas and `inp_ticker` expect.
Windows forbids ':' in file and directory names (a colon silently addresses an
NTFS alternate data stream), so every path derived from a ticker must go
through `fs_ticker()` instead of using the raw string.
"""
from __future__ import annotations

import re

# Leading digit allowed for exchange-qualified international tickers (700:HK).
TICKER_RE = re.compile(r"^[A-Z0-9][A-Z0-9.\-:]{0,14}$")


def validate_ticker(raw: str) -> str:
    """Normalize a CLI-supplied ticker; raise SystemExit if it's malformed."""
    t = (raw or "").strip().upper()
    if not TICKER_RE.match(t):
        raise SystemExit(f"Invalid ticker: {raw!r}. Expected something like AAPL, BRK.B, 700:HK.")
    return t


def fs_ticker(ticker: str) -> str:
    """Filesystem-safe form of a ticker for directory and file names.

    ':' (exchange-qualified tickers like 700:HK) is invalid in Windows paths,
    so it becomes '_'. Keep the raw ticker for CapIQ formula inputs.
    """
    return ticker.replace(":", "_")
