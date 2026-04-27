"""Bootstrap per-ticker scaffolding from the master template.

Usage:
    python -m companies.scripts.new_ticker <TICKER>

Creates:
    companies/output/<TICKER>/<TICKER>_model.xlsx        (copy of templates/company_model.xlsx)
    companies/configs/<TICKER>.yaml                       (skeleton)

After bootstrap, run the fetch + extract + research workflow.
"""
from __future__ import annotations

import argparse
import re
import shutil
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]
MASTER_TEMPLATE = REPO_ROOT / "templates" / "company_model.xlsx"
TICKER_RE = re.compile(r"^[A-Z][A-Z0-9.\-:]{0,14}$")


SKELETON_YAML = """\
ticker: {ticker}
company_name: ""
sector: ""
research_date: {today}
analyst: Aubrey

# drivers, single_drivers, consensus_comparison populated by driver research playbook
"""


def _validate_ticker(raw: str) -> str:
    t = (raw or "").strip().upper()
    if not TICKER_RE.match(t):
        raise SystemExit(f"Invalid ticker: {raw!r}. Expected something like AAPL, BRK.B, 700:HK.")
    return t


def bootstrap(ticker: str) -> None:
    if not MASTER_TEMPLATE.exists():
        raise SystemExit(
            f"Missing {MASTER_TEMPLATE}. Run `python -m shared.scaffold_template` first."
        )

    output_dir = REPO_ROOT / "companies" / "output" / ticker
    configs_dir = REPO_ROOT / "companies" / "configs"
    output_dir.mkdir(parents=True, exist_ok=True)
    configs_dir.mkdir(parents=True, exist_ok=True)

    model_path = output_dir / f"{ticker}_model.xlsx"
    if model_path.exists():
        raise SystemExit(
            f"{model_path} already exists. Delete it explicitly if you want to start fresh."
        )

    shutil.copy2(MASTER_TEMPLATE, model_path)

    # Set inp_ticker named range to the actual ticker.
    wb = load_workbook(model_path)
    if "inp_ticker" not in wb.defined_names:
        raise SystemExit("Master template missing inp_ticker named range; regenerate it.")
    defn = wb.defined_names["inp_ticker"]
    for sheet_name, cell_range in defn.destinations:
        wb[sheet_name][cell_range] = ticker
    wb.save(model_path)

    config_path = configs_dir / f"{ticker}.yaml"
    if not config_path.exists():
        config_path.write_text(
            SKELETON_YAML.format(ticker=ticker, today=date.today().isoformat()),
            encoding="utf-8",
        )

    print(f"Created scaffolding for {ticker}.")
    print()
    print("Next steps:")
    print(f"  1. python -m shared.fetch_capiq {ticker}                          # historicals")
    print(f"  2. python -m shared.fetch_broker_estimates {ticker}               # broker forecasts")
    print(f"  3. python -m companies.scripts.extract_historicals {ticker}       # review brief")
    print(f"  4. python -m companies.scripts.extract_broker_estimates {ticker}  # review consensus")
    print(f"  5. Open Claude Code, ask: \"research drivers for {ticker} using the playbook")
    print(f"     at companies/scripts/driver_research_playbook.md\"")
    print(f"  6. Review companies/configs/{ticker}.yaml and companies/output/{ticker}/drivers_rationale.md")
    print(f"  7. python -m companies.scripts.populate_drivers {ticker}          # write to model")
    print(f"  8. Open the model in Excel for final review")


def main(argv=None):
    parser = argparse.ArgumentParser(description="Bootstrap per-ticker model + config.")
    parser.add_argument("ticker")
    args = parser.parse_args(argv)
    ticker = _validate_ticker(args.ticker)
    bootstrap(ticker)


if __name__ == "__main__":
    main()
