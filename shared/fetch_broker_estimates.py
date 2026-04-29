"""Drive a live broker-estimates refresh and push values into the model.

Usage:
    python -m shared.fetch_broker_estimates <TICKER>
    python -m shared.fetch_broker_estimates <TICKER> --headless
    python -m shared.fetch_broker_estimates <TICKER> --model-path PATH

Workflow:
    1. Open templates/broker_fetcher.xlsx (visible by default).
    2. Set the `broker_fetcher_ticker` named range to the requested ticker.
    3. Force calculation, then wait for CapIQ async queries to resolve.
    4. Validate that Fetcher and _Broker_Data share the same column-A field
       labels.
    5. Read fetched cells (P&L estimates rows 10-17 cols B-G, sentiment B28/B29/B31,
       FY year labels B4-B6) and write to _Broker_Data verbatim. Skip rows
       20-25 and B30/B32 — those are formulas in the main template.
    6. Stamp last-fetch timestamp + ticker on _Broker_Data.
    7. Save target model file; close fetcher without saving.
"""
from __future__ import annotations

import argparse
import re
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

from shared import broker_layout
from shared.excel_session import AppPrefs, get_or_create_app, workbook_already_open
from shared.model_path import resolve_model_path

REPO_ROOT = Path(__file__).resolve().parent.parent
FETCHER_PATH = REPO_ROOT / "templates" / "broker_fetcher.xlsx"
ASYNC_BUFFER_SECS = 3
TICKER_RE = re.compile(r"^[A-Z][A-Z0-9.\-:]{0,14}$")


def _validate_ticker(raw: str) -> str:
    t = (raw or "").strip().upper()
    if not TICKER_RE.match(t):
        raise SystemExit(f"Invalid ticker: {raw!r}. Expected something like AAPL, BRK.B, 700:HK.")
    return t


def _expected_field_labels():
    return broker_layout.all_field_rows()


def _read_field_labels(sheet) -> list[tuple[int, str]]:
    out = []
    for r, _ in _expected_field_labels():
        out.append((r, sheet.range((r, 2)).value))  # column B = field label
    return out


def _validate_layout_match(fetcher_sheet, broker_sheet) -> None:
    expected = _expected_field_labels()
    fet = _read_field_labels(fetcher_sheet)
    brk = _read_field_labels(broker_sheet)
    misaligned = []
    for (r, exp), (_, fv), (_, bv) in zip(expected, fet, brk):
        if (fv or "") != exp or (bv or "") != exp:
            misaligned.append((r, exp, fv, bv))
    if misaligned:
        msg = ["Layout mismatch between Fetcher and _Broker_Data."]
        for r, exp, fv, bv in misaligned:
            msg.append(f"  row {r}: expected {exp!r} | Fetcher: {fv!r} | _Broker_Data: {bv!r}")
        msg.append(
            "Fix: edit shared/broker_layout.py and rerun both scaffolders, "
            "OR manually align the rows in both workbooks."
        )
        raise SystemExit("\n".join(msg))


def _check_capiq_loaded(fetcher_sheet) -> None:
    """C{first PNL row} = IQ_EST_REV(...). #NAME? indicates CapIQ plugin missing."""
    first_pnl_row = broker_layout.PNL[0][0]
    val = fetcher_sheet.range((first_pnl_row, 3)).value
    if isinstance(val, str) and val.strip().startswith("#NAME"):
        raise SystemExit(
            f"CapIQ plugin not loaded (C{first_pnl_row} returned #NAME?). Open Excel "
            f"manually, sign in to S&P Capital IQ, then retry."
        )


def _format_money(v):
    if isinstance(v, (int, float)):
        return f"${v:,.0f}M" if abs(v) >= 1_000_000 else f"${v:,.2f}"
    return repr(v)


# Cells the fetcher provides values for (everything except formula rows).
# New layout: column A is a spacer, so PNL data lives in cols C–H.
PNL_ROWS = [r for r, *_ in broker_layout.PNL]
PNL_COLS = (3, 8)  # C..H inclusive
FY_LABEL_ROWS = [broker_layout.ROW_FY1_YEAR, broker_layout.ROW_FY2_YEAR, broker_layout.ROW_FY3_YEAR]
SENTIMENT_FETCH_ROWS = [r for r, _, fmt in broker_layout.SENTIMENT if fmt is not None]


def fetch(ticker: str, headless: bool = False, model_path_override: str | None = None) -> None:
    model_path = resolve_model_path(ticker, model_path_override)
    if not model_path.exists():
        raise SystemExit(
            f"Missing {model_path}. Run `python -m shared.scaffold_template` first."
        )
    if not FETCHER_PATH.exists():
        raise SystemExit(
            f"Missing {FETCHER_PATH}. Run `python -m shared.scaffold_broker_fetcher` first."
        )
    print(f"Writing broker estimate values to: {model_path}")

    # Attach to a running Excel if one exists (CapIQ auth lives in that
    # session). Otherwise spawn one. xlwings is lazy-imported via
    # excel_session so --help works without it.
    app, owns_app = get_or_create_app(headless=headless)
    if not owns_app:
        for path, label in [(FETCHER_PATH, "broker_fetcher.xlsx"),
                            (model_path, model_path.name)]:
            if workbook_already_open(app, path):
                raise SystemExit(
                    f"{label} is already open in your Excel. "
                    f"Close it before running the fetch script."
                )

    fetcher_wb = None
    model_wb = None
    try:
        prefs = AppPrefs(app)
        prefs.__enter__()
        fetcher_wb = app.books.open(str(FETCHER_PATH), update_links=False)
        fetcher_sheet = fetcher_wb.sheets["Fetcher"]

        try:
            fetcher_wb.names["broker_fetcher_ticker"].refers_to_range.value = ticker
        except Exception:
            fetcher_sheet.range((broker_layout.ROW_TICKER, 3)).value = ticker

        app.calculate()
        try:
            app.api.CalculateUntilAsyncQueriesDone()
        except Exception:
            time.sleep(5)
        time.sleep(ASYNC_BUFFER_SECS)

        _check_capiq_loaded(fetcher_sheet)

        model_wb = app.books.open(str(model_path), update_links=False)
        if "_Broker_Data" not in [s.name for s in model_wb.sheets]:
            raise SystemExit(
                "Template hasn't been regenerated with broker layer. "
                "Run `python -m shared.scaffold_template` first."
            )
        broker_sheet = model_wb.sheets["_Broker_Data"]

        _validate_layout_match(fetcher_sheet, broker_sheet)

        # Stamp metadata first (timestamp + ticker) — values in column C.
        broker_sheet.range((broker_layout.ROW_LAST_FETCH, 3)).value = (
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        )
        broker_sheet.range((broker_layout.ROW_TICKER, 3)).value = ticker

        # FY year labels (col C in new layout)
        fy_values = []
        for r in FY_LABEL_ROWS:
            v = fetcher_sheet.range((r, 3)).value
            broker_sheet.range((r, 3)).value = v
            fy_values.append(v)

        # P&L grid (rows 10-17, cols B-G)
        cells_written = 0
        errors = 0
        err_samples = []
        for r in PNL_ROWS:
            row_values = fetcher_sheet.range((r, PNL_COLS[0]), (r, PNL_COLS[1])).value
            broker_sheet.range((r, PNL_COLS[0]), (r, PNL_COLS[1])).value = row_values
            for v in (row_values or []):
                if v not in (None, ""):
                    cells_written += 1
                if isinstance(v, str) and v.startswith("#"):
                    errors += 1
                    if len(err_samples) < 5:
                        err_samples.append(v)

        # Sentiment fetched rows (column C in new layout). Skip implied-upside
        # and recommendation-distribution rows — those are formulas / blank.
        for r in SENTIMENT_FETCH_ROWS:
            v = fetcher_sheet.range((r, 3)).value
            broker_sheet.range((r, 3)).value = v
            if v not in (None, ""):
                cells_written += 1
            if isinstance(v, str) and v.startswith("#"):
                errors += 1
                if len(err_samples) < 5:
                    err_samples.append(v)

        model_wb.save()

        print(f"Broker estimate fetch complete: {ticker}")
        print(f"  Last fetch: {datetime.now():%Y-%m-%d %H:%M:%S}")
        print(f"  FY labels: FY1={fy_values[0]!r} FY2={fy_values[1]!r} FY3={fy_values[2]!r}")
        print(f"  Cells written: {cells_written}")
        print(f"  Errors: {errors}" + (f"  e.g. {err_samples}" if err_samples else ""))

        # Sample value reads (new layout positions): Revenue FY1 mean = C13,
        # EBITDA FY1 mean = C15, # analysts = C31, avg price target = C32.
        rev_fy1    = broker_sheet.range((13, 3)).value
        ebitda_fy1 = broker_sheet.range((15, 3)).value
        n_analysts = broker_sheet.range((31, 3)).value
        avg_target = broker_sheet.range((32, 3)).value
        print()
        print("  Sample values:")
        print(f"    Revenue FY1 (mean):   {_format_money(rev_fy1)}")
        print(f"    EBITDA FY1 (mean):    {_format_money(ebitda_fy1)}")
        print(f"    # Analysts (rev FY1): {n_analysts!r}")
        print(f"    Avg price target:     {_format_money(avg_target)}")

        if errors:
            print(
                f"\n  WARNING: {errors} cells came back as Excel errors. "
                "If many are #N/A, the IQ_EST_* function names may need adjusting "
                "in broker_fetcher.xlsx (variants like _HIGH/_LOW/_NUM_* are inconsistent)."
            )

    except SystemExit:
        raise
    except Exception:
        print("Unhandled error during fetch:", file=sys.stderr)
        traceback.print_exc()
        raise SystemExit(1)
    finally:
        # Only close workbooks the script itself opened.
        try:
            if model_wb is not None:
                model_wb.close()
        except Exception:
            pass
        try:
            if fetcher_wb is not None:
                fetcher_wb.close()
        except Exception:
            pass
        try:
            prefs.__exit__(None, None, None)
        except Exception:
            pass
        # Only quit the app if we spawned it.
        if owns_app:
            try:
                app.quit()
            except Exception:
                pass


def main(argv=None):
    parser = argparse.ArgumentParser(description="Refresh _Broker_Data in company_model.xlsx for a ticker.")
    parser.add_argument("ticker", help="Ticker to fetch (e.g. AAPL, BRK.B, 700:HK)")
    parser.add_argument("--headless", action="store_true",
                        help="Run Excel hidden. Default is visible so CapIQ auth issues are easy to debug.")
    parser.add_argument("--model-path", default=None,
                        help="Override the model file. Default: per-ticker copy if it exists, else master template.")
    args = parser.parse_args(argv)
    ticker = _validate_ticker(args.ticker)
    fetch(ticker, headless=args.headless, model_path_override=args.model_path)


if __name__ == "__main__":
    main()
