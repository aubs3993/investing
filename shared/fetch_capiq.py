"""Drive a live CapIQ refresh and push values into the main template.

Usage:
    python -m shared.fetch_capiq <TICKER>
    python -m shared.fetch_capiq <TICKER> --headless

Workflow:
    1. Open templates/capiq_fetcher.xlsx (visible by default).
    2. Set the `fetcher_ticker` named range to the requested ticker.
    3. Force calculation, then wait for CapIQ async queries to resolve.
    4. Validate that Fetcher and _CapIQ_Data share the same column-A field
       labels (both come from shared.capiq_layout, but the user may have
       hand-edited one without the other).
    5. Read Fetcher's used range as values; write to _CapIQ_Data verbatim.
    6. Stamp last-fetch timestamp + ticker on _CapIQ_Data.
    7. Save company_model.xlsx; close fetcher without saving.
"""
from __future__ import annotations

import argparse
import re
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

from shared import capiq_layout
from shared.excel_session import AppPrefs, get_or_create_app, workbook_already_open
from shared.model_path import resolve_model_path

REPO_ROOT = Path(__file__).resolve().parent.parent
FETCHER_PATH = REPO_ROOT / "templates" / "capiq_fetcher.xlsx"
ASYNC_BUFFER_SECS = 3
TICKER_RE = re.compile(r"^[A-Z][A-Z0-9.\-:]{0,14}$")


def _validate_ticker(raw: str) -> str:
    t = (raw or "").strip().upper()
    if not TICKER_RE.match(t):
        raise SystemExit(f"Invalid ticker: {raw!r}. Expected something like AAPL, BRK.B, 700:HK.")
    return t


def _expected_field_labels():
    """Ordered list of (row, label) sourced from shared.capiq_layout."""
    return capiq_layout.all_field_rows()


def _read_field_labels(sheet) -> list[tuple[int, str]]:
    """Read column A on a sheet at the rows we expect data on (skip blanks)."""
    labels = []
    for r, _ in _expected_field_labels():
        v = sheet.range((r, 1)).value
        labels.append((r, v))
    return labels


def _validate_layout_match(fetcher_sheet, capiq_sheet) -> None:
    expected = _expected_field_labels()
    fet_labels = _read_field_labels(fetcher_sheet)
    cap_labels = _read_field_labels(capiq_sheet)
    misaligned = []
    for (r, exp_label), (_, fet_label), (_, cap_label) in zip(expected, fet_labels, cap_labels):
        if (fet_label or "") != exp_label or (cap_label or "") != exp_label:
            misaligned.append((r, exp_label, fet_label, cap_label))
    if misaligned:
        msg = ["Layout mismatch between Fetcher and _CapIQ_Data."]
        for r, exp, fet, cap in misaligned:
            msg.append(f"  row {r}: expected {exp!r} | Fetcher: {fet!r} | _CapIQ_Data: {cap!r}")
        msg.append(
            "Fix: edit shared/capiq_layout.py and rerun both scaffolders, OR "
            "manually align the rows in both workbooks."
        )
        raise SystemExit("\n".join(msg))


def _check_capiq_loaded(fetcher_sheet) -> None:
    """E7 holds IQ_COMPANY_NAME; if CapIQ plugin is missing, it returns #NAME?."""
    val = fetcher_sheet.range("E7").value
    if isinstance(val, str) and val.strip().startswith("#NAME"):
        raise SystemExit(
            "CapIQ plugin not loaded (E7 returned #NAME?). Open Excel manually, "
            "sign in to the S&P Capital IQ plugin, then retry."
        )


def _count_errors(values_2d) -> tuple[int, list[str]]:
    err_count = 0
    samples = []
    for row in values_2d or []:
        for v in row:
            if isinstance(v, str) and v.startswith("#"):
                err_count += 1
                if len(samples) < 5:
                    samples.append(v)
    return err_count, samples


def _format_money(v):
    if isinstance(v, (int, float)):
        return f"${v:,.0f}M" if abs(v) >= 1_000_000 else f"${v:,.2f}"
    return repr(v)


def fetch(ticker: str, headless: bool = False, model_path_override: str | None = None) -> None:
    model_path = resolve_model_path(ticker, model_path_override)
    if not model_path.exists():
        raise SystemExit(
            f"Missing {model_path}. Run `python -m shared.scaffold_template` first."
        )
    if not FETCHER_PATH.exists():
        raise SystemExit(
            f"Missing {FETCHER_PATH}. Run `python -m shared.scaffold_capiq_fetcher` first."
        )
    print(f"Writing CapIQ values to: {model_path}")

    # Attach to a running Excel if one exists (CapIQ auth lives in that
    # session). Otherwise spawn one. Lazy-imported via excel_session so that
    # --help works without xlwings/Excel installed.
    app, owns_app = get_or_create_app(headless=headless)
    if not owns_app:
        # Don't risk writing to the wrong instance: refuse if either workbook
        # is already open in the user's Excel.
        for path, label in [(FETCHER_PATH, "capiq_fetcher.xlsx"),
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

        # Drive the ticker via the named range so layout drift in B3 doesn't break us.
        try:
            fetcher_wb.names["fetcher_ticker"].refers_to_range.value = ticker
        except Exception:
            fetcher_sheet.range((capiq_layout.ROW_TICKER, 2)).value = ticker

        # Trigger calc, then wait for async CapIQ queries.
        app.calculate()
        try:
            app.api.CalculateUntilAsyncQueriesDone()
        except Exception:
            # Some Excel versions expose this differently; fall back to a sleep.
            time.sleep(5)
        time.sleep(ASYNC_BUFFER_SECS)

        _check_capiq_loaded(fetcher_sheet)

        model_wb = app.books.open(str(model_path), update_links=False)
        if "_CapIQ_Data" not in [s.name for s in model_wb.sheets]:
            raise SystemExit(
                "Template hasn't been regenerated with CapIQ layer. "
                "Run `python -m shared.scaffold_template` first."
            )
        capiq_sheet = model_wb.sheets["_CapIQ_Data"]

        _validate_layout_match(fetcher_sheet, capiq_sheet)

        # Copy used range Fetcher -> _CapIQ_Data, values only.
        used = fetcher_sheet.used_range
        last_row = used.last_cell.row
        last_col = used.last_cell.column
        rng_addr = (1, 1, last_row, last_col)
        values = fetcher_sheet.range((1, 1), (last_row, last_col)).value

        capiq_sheet.range((1, 1), (last_row, last_col)).value = values

        # Stamp metadata after copy (otherwise the verbatim copy would clobber).
        capiq_sheet.range((capiq_layout.ROW_LAST_FETCH, 2)).value = (
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        )
        capiq_sheet.range((capiq_layout.ROW_TICKER, 2)).value = ticker

        model_wb.save()
        err_count, samples = _count_errors(values)
        cells_refreshed = sum(
            1 for row in (values or []) for v in row if v not in (None, "")
        )

        print(f"CapIQ fetch complete: {ticker}")
        print(f"  Last fetch: {datetime.now():%Y-%m-%d %H:%M:%S}")
        print(f"  Cells refreshed: {cells_refreshed}")
        print(f"  Errors: {err_count}" + (f"  e.g. {samples}" if samples else ""))

        # Sanity-check pulls
        def _val(addr):
            try:
                return capiq_sheet.range(addr).value
            except Exception:
                return None
        rev_fy1 = _val((25, 4))   # D25
        ebitda_fy1 = _val((32, 4))  # D32
        cash = _val((16, 5))      # E16
        debt = _val((17, 5))      # E17
        print()
        print("  Sample values:")
        print(f"    Revenue (FY-1):  {_format_money(rev_fy1)}")
        print(f"    EBITDA (FY-1):   {_format_money(ebitda_fy1)}")
        print(f"    Cash:            {_format_money(cash)}")
        print(f"    Total Debt:      {_format_money(debt)}")

        if err_count:
            print(
                f"\n  WARNING: {err_count} cells came back as Excel errors. "
                f"If this looks like a bad ticker, re-run with the right one. "
                f"If many fields are #N/A, the CapIQ function names may need adjusting "
                f"in capiq_fetcher.xlsx."
            )

    except SystemExit:
        raise
    except Exception:
        print("Unhandled error during fetch:", file=sys.stderr)
        traceback.print_exc()
        raise SystemExit(1)
    finally:
        # Only close workbooks the script itself opened. Never touch books
        # that were already open in an attached Excel session.
        try:
            if model_wb is not None:
                model_wb.close()
        except Exception:
            pass
        try:
            if fetcher_wb is not None:
                fetcher_wb.close()  # don't save — keep fetcher state fresh next run
        except Exception:
            pass
        # Restore screen_updating / display_alerts to their prior values.
        try:
            prefs.__exit__(None, None, None)
        except Exception:
            pass
        # Only quit the app if we spawned it. If we attached to a running
        # Excel (CapIQ auth lives there), leave it alone.
        if owns_app:
            try:
                app.quit()
            except Exception:
                pass


def main(argv=None):
    parser = argparse.ArgumentParser(description="Refresh _CapIQ_Data in company_model.xlsx for a ticker.")
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
