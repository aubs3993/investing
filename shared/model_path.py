"""Resolve which company_model.xlsx a fetch should write to.

The default target is the per-ticker copy at
companies/output/<TICKER>/<TICKER>_model.xlsx. The committed master template
(templates/company_model.xlsx) must never be written by a fetch — if the
per-ticker copy doesn't exist yet, abort and point at new_ticker instead of
silently clobbering the master.
"""
from __future__ import annotations

from pathlib import Path

from shared.tickers import fs_ticker

REPO_ROOT = Path(__file__).resolve().parent.parent


def resolve_model_path(ticker: str, override: str | None = None) -> Path:
    """Return the model file the fetch should write to.

    Priority:
      1. --model-path override (if provided; must exist)
      2. companies/output/<TICKER>/<TICKER>_model.xlsx (per-ticker copy)

    Raises SystemExit if the per-ticker copy is missing — writing to the
    master template requires an explicit --model-path.
    """
    if override:
        p = Path(override)
        if not p.exists():
            raise FileNotFoundError(f"--model-path {p} does not exist")
        return p

    fs = fs_ticker(ticker)
    per_ticker = REPO_ROOT / "companies" / "output" / fs / f"{fs}_model.xlsx"
    if per_ticker.exists():
        return per_ticker
    raise SystemExit(
        f"Missing {per_ticker}. Bootstrap with "
        f"`python -m companies.scripts.new_ticker {ticker}` first. "
        f"(Writing to any other file requires an explicit --model-path.)"
    )
