"""Write driver values from a ticker's YAML config to its model.

Usage:
    python -m companies.scripts.populate_drivers <TICKER>

Reads:
    companies/configs/<TICKER>.yaml
Writes to:
    companies/output/<TICKER>/<TICKER>_model.xlsx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]

# Map YAML driver key -> (named range in workbook, expected length, number format)
DRIVER_MAP = [
    ("revenue_growth", "drv_revenue_growth", 10, "0.0%"),
    ("gross_margin",   "drv_gross_margin",   10, "0.0%"),
    ("opex_pct_rev",   "drv_opex_pct_rev",   10, "0.0%"),
    ("capex_pct_rev",  "drv_capex_pct_rev",  10, "0.0%"),
    ("da_pct_capex",   "drv_da_pct_capex",   10, '0.00"x"'),
    ("exit_multiple",  "drv_exit_multiple",  10, '0.0"x"'),
]

# YAML single_drivers key -> (named range, format)
SINGLE_MAP = [
    ("dps_growth",       "inp_dps_growth",       "0.0%"),
    ("cash_sweep_pct",   "inp_cash_sweep_pct",   "0.0%"),
    ("min_cash_balance", "inp_min_cash",         "#,##0;(#,##0)"),
]


def _config_path_for(ticker: str) -> Path:
    return REPO_ROOT / "companies" / "configs" / f"{ticker}.yaml"


def _model_path_for(ticker: str) -> Path:
    return REPO_ROOT / "companies" / "output" / ticker / f"{ticker}_model.xlsx"


def _validate_config(config: dict) -> None:
    if "drivers" not in config:
        raise SystemExit("YAML missing 'drivers' section.")
    if "single_drivers" not in config:
        raise SystemExit("YAML missing 'single_drivers' section.")
    for key, _name, expected_len, _fmt in DRIVER_MAP:
        if key not in config["drivers"]:
            raise SystemExit(f"YAML drivers missing '{key}'.")
        values = config["drivers"][key].get("values")
        if not isinstance(values, list) or len(values) != expected_len:
            raise SystemExit(
                f"drivers.{key}.values must be a list of {expected_len} numbers; got {values!r}"
            )
        for v in values:
            if not isinstance(v, (int, float)):
                raise SystemExit(f"drivers.{key}.values contains non-numeric: {v!r}")
    for key, _name, _fmt in SINGLE_MAP:
        if key not in config["single_drivers"]:
            raise SystemExit(f"YAML single_drivers missing '{key}'.")
        v = config["single_drivers"][key]
        if not isinstance(v, (int, float)):
            raise SystemExit(f"single_drivers.{key} must be numeric; got {v!r}")


def _resolve_named_range(wb, name):
    """Yield (sheet, list_of_cells) for each destination of a named range."""
    if name not in wb.defined_names:
        raise SystemExit(f"Named range {name!r} missing from workbook.")
    defn = wb.defined_names[name]
    for sheet_name, cell_range in defn.destinations:
        sheet = wb[sheet_name]
        cells = list(sheet[cell_range])
        flat = [c for row in cells for c in row]
        yield sheet, flat


def write_to_named_range(wb, range_name, values, num_format=None):
    written = 0
    for _sheet, cells in _resolve_named_range(wb, range_name):
        if len(cells) != len(values):
            raise SystemExit(
                f"{range_name}: expected {len(cells)} cells, got {len(values)} values"
            )
        for cell, value in zip(cells, values):
            cell.value = value
            if num_format and not cell.number_format or cell.number_format == "General":
                cell.number_format = num_format
            elif num_format and cell.number_format in ("0.0%", "0.0\"x\"", "0.00\"x\"", "#,##0;(#,##0)"):
                cell.number_format = num_format
        written += len(cells)
    return written


def write_single(wb, range_name, value, num_format=None):
    for _sheet, cells in _resolve_named_range(wb, range_name):
        if len(cells) != 1:
            raise SystemExit(f"{range_name}: expected 1 cell, got {len(cells)}")
        cell = cells[0]
        cell.value = value
        if num_format and (not cell.number_format or cell.number_format == "General"):
            cell.number_format = num_format
    return 1


def populate(ticker: str) -> None:
    cfg_path = _config_path_for(ticker)
    model_path = _model_path_for(ticker)
    if not cfg_path.exists():
        raise SystemExit(f"Missing {cfg_path}. Bootstrap with `python -m companies.scripts.new_ticker {ticker}`.")
    if not model_path.exists():
        raise SystemExit(f"Missing {model_path}. Bootstrap with `python -m companies.scripts.new_ticker {ticker}`.")

    try:
        import yaml
    except ImportError:
        raise SystemExit("PyYAML required. Install with `pip install pyyaml` (or `pip install -r requirements.txt`).")

    config = yaml.safe_load(cfg_path.read_text(encoding="utf-8"))
    _validate_config(config)

    wb = load_workbook(model_path)

    print(f"Populating drivers for {ticker} -> {model_path}")
    print()
    print("Drivers:")
    print(f"  {'Named range':<25} {'Values'}")
    for key, name, _expected_len, fmt in DRIVER_MAP:
        values = config["drivers"][key]["values"]
        write_to_named_range(wb, name, values, num_format=fmt)
        rendered = ", ".join(f"{v:.3f}" if isinstance(v, float) else str(v) for v in values)
        print(f"  {name:<25} [{rendered}]")

    print()
    print("Single drivers:")
    for key, name, fmt in SINGLE_MAP:
        v = config["single_drivers"][key]
        write_single(wb, name, v, num_format=fmt)
        print(f"  {name:<25} {v}")

    wb.save(model_path)
    print()
    print(f"Saved {model_path}")


def main(argv=None):
    parser = argparse.ArgumentParser(description="Populate driver values from YAML into a ticker's model.")
    parser.add_argument("ticker")
    args = parser.parse_args(argv)
    ticker = args.ticker.strip().upper()
    populate(ticker)


if __name__ == "__main__":
    main()
