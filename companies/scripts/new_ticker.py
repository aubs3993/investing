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
import shutil
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from shared.tickers import fs_ticker, validate_ticker

REPO_ROOT = Path(__file__).resolve().parents[2]
MASTER_TEMPLATE = REPO_ROOT / "templates" / "company_model.xlsx"


SKELETON_YAML = """\
ticker: {ticker}
company_name: ""
sector: ""
research_date: {today}
analyst: Aubrey

# drivers, single_drivers, consensus_comparison populated by driver research playbook
"""


def bootstrap(ticker: str) -> None:
    if not MASTER_TEMPLATE.exists():
        raise SystemExit(
            f"Missing {MASTER_TEMPLATE}. Run `python -m shared.scaffold_template` first."
        )

    # Paths use the filesystem-safe ticker (':' is invalid in Windows paths);
    # the raw ticker still goes into inp_ticker and the YAML.
    fs = fs_ticker(ticker)
    output_dir = REPO_ROOT / "companies" / "output" / fs
    configs_dir = REPO_ROOT / "companies" / "configs"
    output_dir.mkdir(parents=True, exist_ok=True)
    configs_dir.mkdir(parents=True, exist_ok=True)

    model_path = output_dir / f"{fs}_model.xlsx"
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

    config_path = configs_dir / f"{fs}.yaml"
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
    print(f"  6. Review companies/configs/{fs}.yaml and companies/output/{fs}/drivers_rationale.md")
    print(f"  7. python -m companies.scripts.populate_drivers {ticker}          # write to model")
    print(f"  8. Open the model in Excel for final review")


def main(argv=None):
    parser = argparse.ArgumentParser(description="Bootstrap per-ticker model + config.")
    parser.add_argument("ticker")
    args = parser.parse_args(argv)
    ticker = validate_ticker(args.ticker)
    bootstrap(ticker)


if __name__ == "__main__":
    main()
