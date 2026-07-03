"""Helpers shared by the extract_* brief scripts.

companies/scripts/extract_historicals.py and extract_broker_estimates.py read
the hidden data tabs of a per-ticker model and emit a JSON/markdown brief.
Their cell access, value formatting, and CLI plumbing are identical, so they
live here (per the repo rule: refactor into shared/ as soon as code is
duplicated across two scripts).
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from shared.tickers import validate_ticker


def safe_div(num, den):
    try:
        if den in (None, 0) or num is None:
            return None
        return num / den
    except (TypeError, ZeroDivisionError):
        return None


def get_cell(ws, row, col):
    return ws.cell(row, col).value


def fmt_pct(v): return f"{v*100:.1f}%" if isinstance(v, (int, float)) else "—"
def fmt_num(v): return f"{v:,.0f}" if isinstance(v, (int, float)) else "—"
def fmt_money(v): return f"${v:,.2f}" if isinstance(v, (int, float)) else "—"


def run_extract_cli(extract_fn, to_markdown_fn, description, argv=None):
    """Shared CLI body for the extract scripts: parse args, dispatch, write."""
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("ticker")
    parser.add_argument("--format", choices=["json", "markdown"], default="json")
    parser.add_argument("--output", default=None,
                        help="Write to file instead of stdout.")
    args = parser.parse_args(argv)
    ticker = validate_ticker(args.ticker)
    data = extract_fn(ticker)
    if args.format == "json":
        out = json.dumps(data, indent=2, default=str)
    else:
        out = to_markdown_fn(data)
    if args.output:
        Path(args.output).write_text(out, encoding="utf-8")
        print(f"Wrote {args.output}")
    else:
        sys.stdout.write(out + "\n")
