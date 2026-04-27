"""Resolve which company_model.xlsx to write to.

Per-ticker copy at companies/output/<TICKER>/<TICKER>_model.xlsx wins over
the master template at templates/company_model.xlsx. Without this, fetching
for a real ticker would clobber the master template.
"""
from __future__ import annotations

from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
MASTER_TEMPLATE = REPO_ROOT / "templates" / "company_model.xlsx"


def resolve_model_path(ticker: str, override: str | None = None) -> Path:
    """Return the model file the fetch should write to.

    Priority:
      1. --model-path override (if provided)
      2. companies/output/<TICKER>/<TICKER>_model.xlsx (per-ticker copy)
      3. templates/company_model.xlsx (master)
    """
    if override:
        p = Path(override)
        if not p.exists():
            raise FileNotFoundError(f"--model-path {p} does not exist")
        return p

    per_ticker = REPO_ROOT / "companies" / "output" / ticker / f"{ticker}_model.xlsx"
    if per_ticker.exists():
        return per_ticker
    return MASTER_TEMPLATE
