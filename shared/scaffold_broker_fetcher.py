"""Scaffold templates/broker_fetcher.xlsx — live broker-estimate workbook.

Idempotent. Same conventions as scaffold_capiq_fetcher.py.

Layout MUST mirror _Broker_Data on the main template. Both read from
shared.broker_layout to keep them in lockstep.

Run:
    python -m shared.scaffold_broker_fetcher
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName

from shared import broker_layout

REPO_ROOT = Path(__file__).resolve().parent.parent
FETCHER_PATH = REPO_ROOT / "templates" / "broker_fetcher.xlsx"
TICKER_REF = "broker_fetcher_ticker"

BLUE = "0000FF"
WHITE = "FFFFFF"
YELLOW = "FFFF00"
HEADER_HEX = "1F4E78"
DOTTED_BLUE = Side(border_style="dotted", color=BLUE)

INPUT_FONT = Font(color=BLUE, name="Calibri", size=11)
INPUT_FILL = PatternFill("solid", fgColor=YELLOW)
INPUT_BORDER = Border(left=DOTTED_BLUE, right=DOTTED_BLUE, top=DOTTED_BLUE, bottom=DOTTED_BLUE)
HEADER_FONT = Font(color=WHITE, bold=True, name="Calibri", size=11)
HEADER_FILL = PatternFill("solid", fgColor=HEADER_HEX)
BANNER_FONT = Font(italic=True, color="808080", bold=True, name="Calibri", size=10)
LABEL_BOLD = Font(bold=True, name="Calibri", size=11)
FORMULA_FONT = Font(color="000000", name="Calibri", size=11)


def style_input(cell):
    cell.font = INPUT_FONT
    cell.fill = INPUT_FILL
    cell.border = INPUT_BORDER


def style_header(cell):
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="center")


def style_formula(cell):
    cell.font = FORMULA_FONT


def _est_args(ticker_ref, period, omit_currency):
    if omit_currency:
        return f"{ticker_ref},{period}"
    return f"{ticker_ref},{period},IQ_USD"


def build_fetcher(ws):
    ws.cell(broker_layout.ROW_BANNER, 1,
            "Broker Estimates Fetcher — formulas only. Open in Excel with CapIQ plugin loaded to refresh."
            ).font = BANNER_FONT
    ws.cell(broker_layout.ROW_LAST_FETCH, 1, "Run via:").font = LABEL_BOLD
    ws.cell(broker_layout.ROW_LAST_FETCH, 2, "python -m shared.fetch_broker_estimates <TICKER>")

    ws.cell(broker_layout.ROW_TICKER, 1, "Ticker:").font = LABEL_BOLD
    ticker_cell = ws.cell(broker_layout.ROW_TICKER, 2, "AAPL")
    style_input(ticker_cell)
    ticker_cell.comment = Comment(
        "Set this to the ticker you want broker estimates for. CapIQ formulas "
        "refresh automatically when you change it (and when fetch_broker_estimates.py "
        "drives it programmatically).",
        "scaffold_broker_fetcher",
    )

    # FY year labels — best-effort; user may need to adjust function names per their CapIQ build.
    ws.cell(broker_layout.ROW_FY1_YEAR, 1, "FY1 fiscal year:").font = LABEL_BOLD
    ws.cell(broker_layout.ROW_FY2_YEAR, 1, "FY2 fiscal year:").font = LABEL_BOLD
    ws.cell(broker_layout.ROW_FY3_YEAR, 1, "FY3 fiscal year:").font = LABEL_BOLD
    style_formula(ws.cell(broker_layout.ROW_FY1_YEAR, 2,
                          f"=YEAR(IQ_FY1_FYE_DATE({TICKER_REF}))"))
    style_formula(ws.cell(broker_layout.ROW_FY2_YEAR, 2,
                          f"=YEAR(IQ_FY2_FYE_DATE({TICKER_REF}))"))
    style_formula(ws.cell(broker_layout.ROW_FY3_YEAR, 2,
                          f"=YEAR(IQ_FY3_FYE_DATE({TICKER_REF}))"))

    for j, h in enumerate(broker_layout.COL_HEADERS):
        style_header(ws.cell(broker_layout.ROW_COL_HEADERS, j + 1, h))

    # Section A: P&L estimates
    for row, label, mean_fn, high_fn, low_fn, count_fn, omit_ccy in broker_layout.PNL:
        ws.cell(row, 1, label)
        for col_letter, period in zip(broker_layout.MEAN_COLS, broker_layout.PERIODS):
            style_formula(ws[f"{col_letter}{row}"])
            ws[f"{col_letter}{row}"] = f"={mean_fn}({_est_args(TICKER_REF, period, omit_ccy)})"
            style_formula(ws[f"{col_letter}{row}"])
        ws[f"{broker_layout.HIGH_COL}{row}"] = (
            f"={high_fn}({_est_args(TICKER_REF, 'IQ_FY1', omit_ccy)})"
        )
        style_formula(ws[f"{broker_layout.HIGH_COL}{row}"])
        ws[f"{broker_layout.LOW_COL}{row}"] = (
            f"={low_fn}({_est_args(TICKER_REF, 'IQ_FY1', omit_ccy)})"
        )
        style_formula(ws[f"{broker_layout.LOW_COL}{row}"])
        # Count never takes a currency arg.
        ws[f"{broker_layout.COUNT_COL}{row}"] = f"={count_fn}({TICKER_REF},IQ_FY1)"
        style_formula(ws[f"{broker_layout.COUNT_COL}{row}"])

    # Section B: implied calcs — left blank in fetcher; main template has the formulas.
    for r, label, _ in broker_layout.IMPLIED:
        ws.cell(r, 1, label)

    # Section C: sentiment
    for r, label, formula_template in broker_layout.SENTIMENT:
        ws.cell(r, 1, label)
        if formula_template is None:
            continue
        formula = formula_template.format(t=TICKER_REF)
        style_formula(ws.cell(r, 2, formula))

    ws.column_dimensions["A"].width = 38
    for letter in ["B", "C", "D", "E", "F", "G"]:
        ws.column_dimensions[letter].width = 14


def build():
    wb = Workbook()
    wb.remove(wb.active)
    fetcher = wb.create_sheet("Fetcher")
    build_fetcher(fetcher)

    wb.defined_names[TICKER_REF] = DefinedName(
        name=TICKER_REF,
        attr_text=f"Fetcher!$B${broker_layout.ROW_TICKER}",
    )

    FETCHER_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(FETCHER_PATH)
    return FETCHER_PATH


if __name__ == "__main__":
    out = build()
    print(f"Wrote {out}")
