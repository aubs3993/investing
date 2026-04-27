"""Scaffold templates/company_model.xlsx — base corporate financial model.

Idempotent: re-running overwrites the file. This is the master template only —
runtime scripts copy this file to companies/output/<TICKER>/ before filling.

Run from repo root:
    python -m shared.scaffold_template
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from shared import broker_layout, capiq_layout


REPO_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = REPO_ROOT / "templates" / "company_model.xlsx"


# --- Styles ---
BLUE = "0000FF"
GREEN = "008000"
WHITE = "FFFFFF"
YELLOW = "FFFF00"
HEADER_HEX = "1F4E78"

DOTTED_BLUE = Side(border_style="dotted", color=BLUE)
THIN_BLACK = Side(border_style="thin", color="000000")

INPUT_FONT = Font(color=BLUE, name="Calibri", size=11)
INPUT_FILL = PatternFill("solid", fgColor=YELLOW)
INPUT_BORDER = Border(left=DOTTED_BLUE, right=DOTTED_BLUE, top=DOTTED_BLUE, bottom=DOTTED_BLUE)
LINK_FONT = Font(color=GREEN, name="Calibri", size=11)
FORMULA_FONT = Font(color="000000", name="Calibri", size=11)
SUBTOTAL_FONT = Font(color="000000", bold=True, name="Calibri", size=11)
SUBTOTAL_BORDER = Border(top=THIN_BLACK)
HEADER_FONT = Font(color=WHITE, bold=True, name="Calibri", size=11)
HEADER_FILL = PatternFill("solid", fgColor=HEADER_HEX)
SECTION_FONT = Font(bold=True, size=12, name="Calibri")
TITLE_FONT = Font(bold=True, size=14, name="Calibri")
LABEL_BOLD = Font(bold=True, name="Calibri", size=11)
BANNER_FONT = Font(italic=True, color="808080", name="Calibri", size=10)


# --- Number formats ---
NUM = "#,##0;(#,##0)"
NUM_DEC = "#,##0.00;(#,##0.00)"
PCT = "0.0%"
MULT = '0.0"x"'
DATE = "mm/dd/yyyy"
PRICE = "$#,##0.00;($#,##0.00)"
DPS = "$0.0000"
TEXT = "@"


def style_input(cell, num_format=NUM):
    cell.font = INPUT_FONT
    cell.fill = INPUT_FILL
    cell.border = INPUT_BORDER
    cell.number_format = num_format


def style_formula(cell, num_format=NUM):
    cell.font = FORMULA_FONT
    cell.number_format = num_format


def style_link(cell, num_format=NUM):
    cell.font = LINK_FONT
    cell.number_format = num_format


def style_subtotal(cell, num_format=NUM):
    cell.font = SUBTOTAL_FONT
    cell.border = SUBTOTAL_BORDER
    cell.number_format = num_format


def style_header(cell):
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="center")


def add_named_range(wb, name, sheet, ref):
    wb.defined_names[name] = DefinedName(name=name, attr_text=f"{sheet}!{ref}")


def make_year_columns():
    cy = datetime.now().year
    return (
        [(cy + o, f"{cy + o}A") for o in (-3, -2, -1)]
        + [(cy + o, f"{cy + o}E") for o in range(0, 10)]
    )


def col(i):
    """0-indexed year position -> Excel column letter (B..N)."""
    return get_column_letter(i + 2)


# --- Inputs tab layout (fixed row positions other tabs reference) ---
# Tuple: (label, named_range, num_format, default, capiq_link)
# capiq_link: formula string -> green link to _CapIQ_Data; None -> blue input cell
INPUTS_FIRST_ROW = 4
INPUT_ROWS = [
    ("Ticker",                     "inp_ticker",              TEXT,  "TICKR",        None),
    ("Company Name",               "inp_company_name",        TEXT,  "Company Name", "=_CapIQ_Data!E7"),
    ("Sector",                     "inp_sector",              TEXT,  "Sector",       "=_CapIQ_Data!E8"),
    ("Currency",                   "inp_currency",            TEXT,  "USD",          "=_CapIQ_Data!E9"),
    ("Fiscal Year End Month",      "inp_fye_month",           "0",   12,             None),
    ("Run Date",                   "inp_run_date",            DATE,  None,           None),
    ("Current Price",              "inp_current_price",       PRICE, 0,              "=_CapIQ_Data!E13"),
    ("Diluted Shares Outstanding", "inp_diluted_shares",      NUM,   0,              "=_CapIQ_Data!E14"),
    ("Quarterly DPS",              "inp_quarterly_dps",       DPS,   0,              "=_CapIQ_Data!E15"),
    ("DPS Annual Growth %",        "inp_dps_growth",          PCT,   0,              None),
    ("Tax Rate",                   "inp_tax_rate",            PCT,   0.21,           "=_CapIQ_Data!E20"),
    ("Cash & Equivalents",         "inp_cash",                NUM,   0,              "=_CapIQ_Data!E16"),
    ("Total Debt",                 "inp_debt",                NUM,   0,              "=_CapIQ_Data!E17"),
    ("Minority Interest (NCI)",    "inp_minority_interest",   NUM,   0,              "=_CapIQ_Data!E18"),
    ("Equity Investments",         "inp_equity_investments",  NUM,   0,              "=_CapIQ_Data!E19"),
    ("Cash Sweep %",               "inp_cash_sweep_pct",      PCT,   0,              None),
    ("Minimum Cash Balance",       "inp_min_cash",            NUM,   0,              None),
]
SECTION_B_HEADER_ROW = 22
CALC_ROWS = [
    ("Annual DPS (current)",                "=inp_quarterly_dps*4",                                                        "annual_dps",      DPS),
    ("Total Annual Dividends Paid (current)", "=annual_dps*inp_diluted_shares",                                            "calc_annual_div", NUM),
    ("Current Market Cap",                  "=inp_current_price*inp_diluted_shares",                                       "mkt_cap",         NUM),
    ("Net Debt",                            "=inp_debt-inp_cash",                                                          None,              NUM),
    ("Current Enterprise Value",            "=mkt_cap+inp_debt-inp_cash+inp_minority_interest-inp_equity_investments",     None,              NUM),
]
SECTION_C_HEADER_ROW = 29
DRIVER_HEADER_ROW = 30
DRIVER_FIRST_ROW = 31
DRIVERS = [
    ("Revenue Growth %",         PCT,  0.05),
    ("Gross Margin %",           PCT,  0.40),
    ("Total OpEx % of Revenue",  PCT,  0.20),
    ("CapEx % of Revenue",       PCT,  0.05),
    ("D&A % of CapEx",           PCT,  1.0),
    ("Exit EBITDA Multiple",     MULT, 10),
]
DRV_REV = 31
DRV_GM = 32
DRV_OPEX = 33
DRV_CAPEX = 34
DRV_DA = 35
DRV_EXIT = 36


def build_inputs(ws, years):
    ws["A1"] = "Inputs & Assumptions"
    ws["A1"].font = TITLE_FONT

    ws["A3"] = "Inputs (CapIQ-driven where available; manual otherwise)"
    ws["A3"].font = SECTION_FONT
    for i, (label, _name, fmt, default, link) in enumerate(INPUT_ROWS):
        r = INPUTS_FIRST_ROW + i
        ws.cell(r, 1, label)
        if link is not None:
            c = ws.cell(r, 2, link)
            style_link(c, num_format=fmt)
        else:
            val = default if default is not None else datetime.now().date()
            c = ws.cell(r, 2, val)
            style_input(c, num_format=fmt)

    ws.cell(SECTION_B_HEADER_ROW, 1, "Calculated").font = SECTION_FONT
    for i, (label, formula, _name, fmt) in enumerate(CALC_ROWS):
        r = SECTION_B_HEADER_ROW + 1 + i
        ws.cell(r, 1, label)
        c = ws.cell(r, 2, formula)
        style_formula(c, num_format=fmt)

    ws.cell(SECTION_C_HEADER_ROW, 1, "Driver Assumptions (manual fill)").font = SECTION_FONT
    for i, (_yr, lbl) in enumerate(years):
        if lbl.endswith("E"):
            style_header(ws.cell(DRIVER_HEADER_ROW, i + 2, lbl))
    for i, (label, fmt, default) in enumerate(DRIVERS):
        r = DRIVER_FIRST_ROW + i
        ws.cell(r, 1, label)
        for j, (_yr, lbl) in enumerate(years):
            if lbl.endswith("E"):
                c = ws.cell(r, j + 2, default)
                style_input(c, num_format=fmt)

    ws.column_dimensions["A"].width = 35
    for j in range(len(years)):
        ws.column_dimensions[col(j)].width = 14


def register_inputs_named_ranges(wb):
    for i, (_label, name, _fmt, _default, _link) in enumerate(INPUT_ROWS):
        r = INPUTS_FIRST_ROW + i
        add_named_range(wb, name, "Inputs", f"$B${r}")
    for i, (_label, _formula, name, _fmt) in enumerate(CALC_ROWS):
        if name is None:
            continue
        r = SECTION_B_HEADER_ROW + 1 + i
        add_named_range(wb, name, "Inputs", f"$B${r}")


def build_cover(ws, _years):
    ws["A1"] = "Cover"
    ws["A1"].font = TITLE_FONT
    rows = [
        ("Company Name",  "=inp_company_name",  TEXT),
        ("Ticker",        "=inp_ticker",        TEXT),
        ("Sector",        "=inp_sector",        TEXT),
        ("Currency",      "=inp_currency",      TEXT),
        ("Run Date",      "=inp_run_date",      DATE),
        ("Current Price", "=inp_current_price", PRICE),
        ("Market Cap",    "=mkt_cap",           NUM),
    ]
    for i, (label, formula, fmt) in enumerate(rows):
        r = 3 + i
        ws.cell(r, 1, label).font = LABEL_BOLD
        c = ws.cell(r, 2, formula)
        style_link(c, num_format=fmt)
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 28


def _is_proj(years, i):
    return years[i][1].endswith("E")


def _set_year_headers(ws, years, projection_only=False):
    for i, (_yr, lbl) in enumerate(years):
        if projection_only and not lbl.endswith("E"):
            continue
        style_header(ws.cell(1, i + 2, lbl))


def _set_widths(ws, n):
    ws.column_dimensions["A"].width = 35
    for i in range(n):
        ws.column_dimensions[col(i)].width = 14


def build_is(ws, years):
    n = len(years)
    _set_year_headers(ws, years)

    ws.cell(2, 1, "Revenue")
    for i in range(n):
        cur = col(i)
        if not _is_proj(years, i):
            style_link(ws.cell(2, i + 2, f"=_CapIQ_Data!{cur}25"))
        else:
            prev = col(i - 1)
            style_formula(ws.cell(2, i + 2, f"={prev}2*(1+Inputs!{cur}${DRV_REV})"))

    ws.cell(3, 1, "Revenue Growth %")
    for i in range(n):
        cur = col(i)
        if i == 0:
            ws.cell(3, i + 2, None)
        else:
            prev = col(i - 1)
            style_formula(ws.cell(3, i + 2, f"={cur}2/{prev}2-1"), num_format=PCT)

    ws.cell(4, 1, "COGS")
    for i in range(n):
        cur = col(i)
        if not _is_proj(years, i):
            style_link(ws.cell(4, i + 2, f"=_CapIQ_Data!{cur}26"))
        else:
            style_formula(ws.cell(4, i + 2, f"={cur}2*(1-Inputs!{cur}${DRV_GM})"))

    ws.cell(5, 1, "Gross Profit")
    for i in range(n):
        cur = col(i)
        style_subtotal(ws.cell(5, i + 2, f"={cur}2-{cur}4"))

    ws.cell(6, 1, "Gross Margin %")
    for i in range(n):
        cur = col(i)
        style_formula(ws.cell(6, i + 2, f"=IFERROR({cur}5/{cur}2,0)"), num_format=PCT)

    ws.cell(7, 1, "Total OpEx")
    for i in range(n):
        cur = col(i)
        if not _is_proj(years, i):
            style_link(ws.cell(7, i + 2, f"=_CapIQ_Data!{cur}30"))
        else:
            style_formula(ws.cell(7, i + 2, f"={cur}2*Inputs!{cur}${DRV_OPEX}"))

    ws.cell(8, 1, "EBITDA")
    for i in range(n):
        cur = col(i)
        style_subtotal(ws.cell(8, i + 2, f"={cur}5-{cur}7"))

    ws.cell(9, 1, "EBITDA Margin %")
    for i in range(n):
        cur = col(i)
        style_formula(ws.cell(9, i + 2, f"=IFERROR({cur}8/{cur}2,0)"), num_format=PCT)

    ws.cell(10, 1, "D&A")
    for i in range(n):
        cur = col(i)
        if not _is_proj(years, i):
            style_link(ws.cell(10, i + 2, f"=_CapIQ_Data!{cur}31"))
        else:
            # CapEx on CF is negative; negate, then multiply by D&A%
            style_link(ws.cell(10, i + 2, f"=-CF!{cur}5*Inputs!{cur}${DRV_DA}"))

    ws.cell(11, 1, "EBIT")
    for i in range(n):
        cur = col(i)
        style_subtotal(ws.cell(11, i + 2, f"={cur}8-{cur}10"))

    ws.cell(12, 1, "EBIT Margin %")
    for i in range(n):
        cur = col(i)
        style_formula(ws.cell(12, i + 2, f"=IFERROR({cur}11/{cur}2,0)"), num_format=PCT)

    ws.cell(13, 1, "Interest Expense")
    for i in range(n):
        cur = col(i)
        if not _is_proj(years, i):
            style_link(ws.cell(13, i + 2, f"=_CapIQ_Data!{cur}34"))
        else:
            style_link(ws.cell(13, i + 2, f"=Debt!{cur}30"))

    ws.cell(14, 1, "Interest Income")
    first_proj = next(j for j, (_y, l) in enumerate(years) if l.endswith("E"))
    for i in range(n):
        cur = col(i)
        if not _is_proj(years, i):
            style_link(ws.cell(14, i + 2, f"=_CapIQ_Data!{cur}35"))
        elif i == first_proj:
            style_input(ws.cell(14, i + 2, 0))
        else:
            prev = col(i - 1)
            style_formula(ws.cell(14, i + 2, f"={prev}14"))

    ws.cell(15, 1, "Pre-tax Income")
    for i in range(n):
        cur = col(i)
        style_subtotal(ws.cell(15, i + 2, f"={cur}11-{cur}13+{cur}14"))

    ws.cell(16, 1, "Taxes")
    for i in range(n):
        cur = col(i)
        style_formula(ws.cell(16, i + 2, f"={cur}15*inp_tax_rate"))

    ws.cell(17, 1, "Effective Tax Rate %")
    for i in range(n):
        cur = col(i)
        style_formula(ws.cell(17, i + 2, f"=IFERROR({cur}16/{cur}15,0)"), num_format=PCT)

    ws.cell(18, 1, "Net Income")
    for i in range(n):
        cur = col(i)
        style_subtotal(ws.cell(18, i + 2, f"={cur}15-{cur}16"))

    ws.cell(19, 1, "Diluted Shares Outstanding")
    for i in range(n):
        cur = col(i)
        if not _is_proj(years, i):
            style_link(ws.cell(19, i + 2, f"=_CapIQ_Data!{cur}39"))
        else:
            style_link(ws.cell(19, i + 2, "=inp_diluted_shares"))

    ws.cell(20, 1, "Diluted EPS")
    for i in range(n):
        cur = col(i)
        style_formula(ws.cell(20, i + 2, f"=IFERROR({cur}18/{cur}19,0)"), num_format=NUM_DEC)

    ws.freeze_panes = "B2"
    _set_widths(ws, n)


def build_cf(ws, years):
    n = len(years)
    _set_year_headers(ws, years)

    ws.cell(2, 1, "EBITDA")
    for i in range(n):
        cur = col(i)
        style_link(ws.cell(2, i + 2, f"=IS!{cur}8"))

    ws.cell(3, 1, "(Less) Cash Taxes")
    for i in range(n):
        cur = col(i)
        style_link(ws.cell(3, i + 2, f"=-IS!{cur}16"))

    ws.cell(4, 1, "(Less) Cash Interest")
    for i in range(n):
        cur = col(i)
        style_link(ws.cell(4, i + 2, f"=-IS!{cur}13"))

    ws.cell(5, 1, "(Less) CapEx")
    for i in range(n):
        cur = col(i)
        if not _is_proj(years, i):
            style_link(ws.cell(5, i + 2, f"=-_CapIQ_Data!{cur}40"))
        else:
            style_formula(ws.cell(5, i + 2, f"=-IS!{cur}2*Inputs!{cur}${DRV_CAPEX}"))

    ws.cell(6, 1, "Levered Free Cash Flow")
    for i in range(n):
        cur = col(i)
        style_subtotal(ws.cell(6, i + 2, f"=SUM({cur}2:{cur}5)"))

    ws.cell(7, 1, "(Less) Debt Amortization")
    for i in range(n):
        cur = col(i)
        style_link(ws.cell(7, i + 2, f"=-(Debt!{cur}31+Debt!{cur}32)"))

    ws.cell(8, 1, "(Less) Dividends")
    first_proj = next(j for j, (_y, l) in enumerate(years) if l.endswith("E"))
    for i in range(n):
        cur = col(i)
        if not _is_proj(years, i):
            style_input(ws.cell(8, i + 2, 0))
        else:
            k = i - first_proj  # 0 for first projection year
            f = f"=-(inp_quarterly_dps*4)*((1+inp_dps_growth)^{k})*IS!{cur}19"
            style_formula(ws.cell(8, i + 2, f))

    ws.cell(9, 1, "Net Change in Cash")
    for i in range(n):
        cur = col(i)
        style_subtotal(ws.cell(9, i + 2, f"={cur}6+{cur}7+{cur}8"))

    ws.cell(10, 1, "Beginning Cash")
    for i in range(n):
        cur = col(i)
        if i == 0:
            style_input(ws.cell(10, i + 2, 0))
        else:
            prev = col(i - 1)
            style_formula(ws.cell(10, i + 2, f"={prev}11"))

    ws.cell(11, 1, "Ending Cash")
    for i in range(n):
        cur = col(i)
        style_subtotal(ws.cell(11, i + 2, f"={cur}10+{cur}9"))

    ws.freeze_panes = "B2"
    _set_widths(ws, n)


def build_debt(ws, years):
    n = len(years)
    _set_year_headers(ws, years)

    # Block 1: Revolver
    ws.cell(2, 1, "Revolver").font = SECTION_FONT
    ws.cell(3, 1, "Beginning Balance")
    ws.cell(4, 1, "Draws")
    ws.cell(5, 1, "Repayments")
    ws.cell(6, 1, "Ending Balance")
    ws.cell(7, 1, "Average Balance")
    ws.cell(8, 1, "Interest Rate %")
    ws.cell(9, 1, "Interest Expense")
    for i in range(n):
        cur = col(i)
        if i == 0:
            style_input(ws.cell(3, i + 2, 0))
        else:
            prev = col(i - 1)
            style_formula(ws.cell(3, i + 2, f"={prev}6"))
        style_input(ws.cell(4, i + 2, 0))
        style_input(ws.cell(5, i + 2, 0))
        style_subtotal(ws.cell(6, i + 2, f"={cur}3+{cur}4-{cur}5"))
        style_formula(ws.cell(7, i + 2, f"=({cur}3+{cur}6)/2"))
        style_input(ws.cell(8, i + 2, 0.06), num_format=PCT)
        style_formula(ws.cell(9, i + 2, f"={cur}7*{cur}8"))

    # Block 2: Term Loan
    ws.cell(11, 1, "Term Loan").font = SECTION_FONT
    ws.cell(12, 1, "Beginning Balance")
    ws.cell(13, 1, "Mandatory Amortization")
    ws.cell(14, 1, "Optional Prepayment (Cash Sweep)")
    ws.cell(15, 1, "Ending Balance")
    ws.cell(16, 1, "Average Balance")
    ws.cell(17, 1, "Interest Rate %")
    ws.cell(18, 1, "Interest Expense")
    for i in range(n):
        cur = col(i)
        if i == 0:
            style_input(ws.cell(12, i + 2, 0))
        else:
            prev = col(i - 1)
            style_formula(ws.cell(12, i + 2, f"={prev}15"))
        style_input(ws.cell(13, i + 2, 0))
        if not _is_proj(years, i):
            style_formula(ws.cell(14, i + 2, 0))
        else:
            style_link(ws.cell(14, i + 2, f"={cur}37"))
        style_subtotal(ws.cell(15, i + 2, f"={cur}12-{cur}13-{cur}14"))
        style_formula(ws.cell(16, i + 2, f"=({cur}12+{cur}15)/2"))
        style_input(ws.cell(17, i + 2, 0.07), num_format=PCT)
        style_formula(ws.cell(18, i + 2, f"={cur}16*{cur}17"))

    # Block 3: Senior Notes
    ws.cell(20, 1, "Senior Notes").font = SECTION_FONT
    ws.cell(21, 1, "Beginning Balance")
    ws.cell(22, 1, "Repayment at Maturity")
    ws.cell(23, 1, "Ending Balance")
    ws.cell(24, 1, "Interest Rate %")
    ws.cell(25, 1, "Interest Expense")
    for i in range(n):
        cur = col(i)
        if i == 0:
            style_input(ws.cell(21, i + 2, 0))
        else:
            prev = col(i - 1)
            style_formula(ws.cell(21, i + 2, f"={prev}23"))
        style_input(ws.cell(22, i + 2, 0))
        style_subtotal(ws.cell(23, i + 2, f"={cur}21-{cur}22"))
        style_input(ws.cell(24, i + 2, 0.05), num_format=PCT)
        style_formula(ws.cell(25, i + 2, f"={cur}21*{cur}24"))

    # Block 4: Totals
    ws.cell(27, 1, "Totals").font = SECTION_FONT
    ws.cell(28, 1, "Total Beginning Debt")
    ws.cell(29, 1, "Total Ending Debt")
    ws.cell(30, 1, "Total Interest Expense")
    ws.cell(31, 1, "Total Mandatory Amortization")
    ws.cell(32, 1, "Total Optional Prepayment")
    for i in range(n):
        cur = col(i)
        style_subtotal(ws.cell(28, i + 2, f"={cur}3+{cur}12+{cur}21"))
        style_subtotal(ws.cell(29, i + 2, f"={cur}6+{cur}15+{cur}23"))
        style_subtotal(ws.cell(30, i + 2, f"={cur}9+{cur}18+{cur}25"))
        style_formula(ws.cell(31, i + 2, f"={cur}13"))
        style_formula(ws.cell(32, i + 2, f"={cur}14"))

    # Block 5: Cash Sweep
    ws.cell(34, 1, "Cash Sweep").font = SECTION_FONT
    ws.cell(35, 1, "Available Cash for Sweep")
    ws.cell(36, 1, "Cash Sweep Applied")
    ws.cell(37, 1, "Allocated to Term Loan")
    for i in range(n):
        cur = col(i)
        if not _is_proj(years, i):
            style_formula(ws.cell(35, i + 2, 0))
            style_formula(ws.cell(36, i + 2, 0))
            style_formula(ws.cell(37, i + 2, 0))
        else:
            style_formula(ws.cell(35, i + 2, f"=MAX(0,CF!{cur}11-inp_min_cash)"))
            style_formula(ws.cell(36, i + 2, f"={cur}35*inp_cash_sweep_pct"))
            style_formula(ws.cell(37, i + 2, f"={cur}36"))

    ws.freeze_panes = "B2"
    _set_widths(ws, n)


def build_valuation(ws, years):
    n = len(years)
    _set_year_headers(ws, years, projection_only=True)

    proj_indices = [i for i, (_y, lbl) in enumerate(years) if lbl.endswith("E")]
    ws.cell(2, 1, "Year-end EBITDA")
    ws.cell(3, 1, "Exit EBITDA Multiple")
    ws.cell(4, 1, "Implied Enterprise Value")
    ws.cell(5, 1, "(Less) Year-end Total Debt")
    ws.cell(6, 1, "Plus: Year-end Cash")
    ws.cell(7, 1, "(Less) Minority Interest")
    ws.cell(8, 1, "Plus: Equity Investments")
    ws.cell(9, 1, "Implied Equity Value")
    ws.cell(10, 1, "/ Diluted Shares Outstanding")
    ws.cell(11, 1, "Implied Price per Share")
    ws.cell(12, 1, "Implied IRR from Today")
    ws.cell(13, 1, "Implied MOIC")

    for k, i in enumerate(proj_indices):
        cur = col(i)
        style_link(ws.cell(2, i + 2, f"=IS!{cur}8"))
        style_link(ws.cell(3, i + 2, f"=Inputs!{cur}${DRV_EXIT}"), num_format=MULT)
        style_subtotal(ws.cell(4, i + 2, f"={cur}2*{cur}3"))
        style_link(ws.cell(5, i + 2, f"=-Debt!{cur}29"))
        style_link(ws.cell(6, i + 2, f"=CF!{cur}11"))
        style_link(ws.cell(7, i + 2, "=-inp_minority_interest"))
        style_link(ws.cell(8, i + 2, "=inp_equity_investments"))
        style_subtotal(ws.cell(9, i + 2, f"={cur}4+{cur}5+{cur}6+{cur}7+{cur}8"))
        style_link(ws.cell(10, i + 2, f"=IS!{cur}19"))
        style_subtotal(ws.cell(11, i + 2, f"=IFERROR({cur}9/{cur}10,0)"), num_format=NUM_DEC)
        year_number = k + 1
        style_formula(
            ws.cell(12, i + 2, f"=IFERROR(({cur}11/inp_current_price)^(1/{year_number})-1,0)"),
            num_format=PCT,
        )
        style_formula(ws.cell(13, i + 2, f"=IFERROR({cur}11/inp_current_price,0)"), num_format=MULT)

    ws.freeze_panes = "B2"
    _set_widths(ws, n)


def build_summary(ws, years):
    n = len(years)

    ws.cell(1, 1, "Company Name").font = LABEL_BOLD
    style_link(ws.cell(1, 2, "=inp_company_name"), num_format=TEXT)
    ws.cell(2, 1, "Ticker").font = LABEL_BOLD
    style_link(ws.cell(2, 2, "=inp_ticker"), num_format=TEXT)
    ws.cell(3, 1, "Sector").font = LABEL_BOLD
    style_link(ws.cell(3, 2, "=inp_sector"), num_format=TEXT)

    ws.cell(5, 1, "Financial Summary").font = SECTION_FONT
    for i, (_yr, lbl) in enumerate(years):
        style_header(ws.cell(6, i + 2, lbl))

    ws.cell(7, 1, "Revenue")
    ws.cell(8, 1, "Revenue Growth %")
    ws.cell(9, 1, "EBITDA")
    ws.cell(10, 1, "EBITDA Margin %")
    ws.cell(11, 1, "Levered Free Cash Flow")
    ws.cell(12, 1, "LFCF Margin %")
    ws.cell(13, 1, "Net Debt")
    ws.cell(14, 1, "Net Leverage (Net Debt / EBITDA)")

    for i in range(n):
        cur = col(i)
        style_link(ws.cell(7, i + 2, f"=IS!{cur}2"))
        style_link(ws.cell(8, i + 2, f"=IS!{cur}3"), num_format=PCT)
        style_link(ws.cell(9, i + 2, f"=IS!{cur}8"))
        style_link(ws.cell(10, i + 2, f"=IS!{cur}9"), num_format=PCT)
        style_link(ws.cell(11, i + 2, f"=CF!{cur}6"))
        style_formula(ws.cell(12, i + 2, f"=IFERROR({cur}11/{cur}7,0)"), num_format=PCT)
        style_formula(ws.cell(13, i + 2, f"=Debt!{cur}29-CF!{cur}11"))
        style_subtotal(ws.cell(14, i + 2, f"=IFERROR({cur}13/{cur}9,0)"), num_format=MULT)

    ws.cell(16, 1, "Returns Table").font = SECTION_FONT
    for j, h in enumerate(["Year", "Implied Price per Share", "Implied IRR", "Implied MOIC"]):
        style_header(ws.cell(17, j + 1, h))

    proj_indices = [i for i, (_y, lbl) in enumerate(years) if lbl.endswith("E")]
    for k, (yr_off, label) in enumerate([(1, "Y1"), (3, "Y3"), (5, "Y5"), (10, "Y10")]):
        r = 18 + k
        i_col = proj_indices[yr_off - 1]
        cur = col(i_col)
        ws.cell(r, 1, f"{label} ({years[i_col][1]})").font = LABEL_BOLD
        style_link(ws.cell(r, 2, f"=Valuation!{cur}11"), num_format=NUM_DEC)
        style_link(ws.cell(r, 3, f"=Valuation!{cur}12"), num_format=PCT)
        style_link(ws.cell(r, 4, f"=Valuation!{cur}13"), num_format=MULT)

    ws.freeze_panes = "B2"
    _set_widths(ws, n)


def build_sensitivity(ws, _years):
    ws.cell(1, 1, "Exit Multiple Sensitivity").font = TITLE_FONT
    ws.cell(2, 1, "Target Year Offset (1 = first projection year)")
    style_input(ws.cell(2, 2, 5), num_format="0")  # named sens_target_year_offset

    ws.cell(4, 1, "Exit Multiple").font = LABEL_BOLD
    header_b = ws.cell(4, 2, '="Implied Price per Share at Year " & sens_target_year_offset')
    header_b.font = LABEL_BOLD

    multiples = [6, 7, 8, 9, 10, 11, 12, 13, 14]
    for k, m in enumerate(multiples):
        r = 5 + k
        style_input(ws.cell(r, 1, m), num_format=MULT)
        formula = (
            f"=(INDEX(IS!$E$8:$N$8,sens_target_year_offset)*$A{r}"
            f"-INDEX(Debt!$E$29:$N$29,sens_target_year_offset)"
            f"+INDEX(CF!$E$11:$N$11,sens_target_year_offset)"
            f"-inp_minority_interest+inp_equity_investments)"
            f"/INDEX(IS!$E$19:$N$19,sens_target_year_offset)"
        )
        style_formula(ws.cell(r, 2, formula), num_format=NUM_DEC)

    ws.cell(15, 1, "2D Sensitivity (Exit Multiple x Year)").font = SECTION_FONT
    for j, h in enumerate(["Multiple", "Y3", "Y5", "Y7", "Y10"]):
        style_header(ws.cell(16, j + 1, h))

    year_offsets = [3, 5, 7, 10]
    for k, m in enumerate(multiples):
        r = 17 + k
        style_input(ws.cell(r, 1, m), num_format=MULT)
        for j, off in enumerate(year_offsets):
            formula = (
                f"=(INDEX(IS!$E$8:$N$8,{off})*$A{r}"
                f"-INDEX(Debt!$E$29:$N$29,{off})"
                f"+INDEX(CF!$E$11:$N$11,{off})"
                f"-inp_minority_interest+inp_equity_investments)"
                f"/INDEX(IS!$E$19:$N$19,{off})"
            )
            style_formula(ws.cell(r, j + 2, formula), num_format=NUM_DEC)

    ws.column_dimensions["A"].width = 38
    for letter in ["B", "C", "D", "E"]:
        ws.column_dimensions[letter].width = 20


def build_capiq_data(ws):
    """Mirror of capiq_fetcher.xlsx Fetcher tab — values only, no CapIQ formulas.

    Layout sourced from shared.capiq_layout so this stays in lockstep with
    the fetcher. shared/fetch_capiq.py writes hardcoded values into the data
    cells; the cells are left empty here.
    """
    ws.sheet_state = "hidden"
    ws.sheet_properties.tabColor = "808080"

    ws.cell(capiq_layout.ROW_BANNER, 1,
            "_CapIQ_Data — DO NOT EDIT MANUALLY. Populated by shared/fetch_capiq.py."
            ).font = BANNER_FONT
    ws.cell(capiq_layout.ROW_LAST_FETCH, 1, "Last fetch:").font = LABEL_BOLD
    ws.cell(capiq_layout.ROW_TICKER, 1, "Ticker:").font = LABEL_BOLD

    for j, h in enumerate(capiq_layout.COL_HEADERS):
        style_header(ws.cell(capiq_layout.ROW_COL_HEADERS, j + 1, h))

    for r, label, _ in capiq_layout.METADATA:
        ws.cell(r, 1, label)
    for r, label, _ in capiq_layout.CURRENT:
        ws.cell(r, 1, label)
    for r, label, _ in capiq_layout.HISTORICAL:
        ws.cell(r, 1, label)

    ws.column_dimensions["A"].width = 35
    for letter in ["B", "C", "D", "E"]:
        ws.column_dimensions[letter].width = 16


def build_broker_data(ws):
    """Mirror of broker_fetcher.xlsx Fetcher tab.

    Rows 10-17 (P&L) and rows 28, 29, 31 (sentiment fetched values) are
    populated by shared/fetch_broker_estimates.py at runtime. Rows 20-25
    (implied growth/margins) and B30 (implied upside) are formulas that
    live in this workbook so they update when historicals refresh.
    """
    ws.sheet_state = "hidden"
    ws.sheet_properties.tabColor = "808080"

    ws.cell(broker_layout.ROW_BANNER, 1,
            "_Broker_Data — DO NOT EDIT MANUALLY. Populated by shared/fetch_broker_estimates.py."
            ).font = BANNER_FONT
    ws.cell(broker_layout.ROW_LAST_FETCH, 1, "Last fetch:").font = LABEL_BOLD
    ws.cell(broker_layout.ROW_TICKER, 1, "Ticker:").font = LABEL_BOLD
    ws.cell(broker_layout.ROW_FY1_YEAR, 1, "FY1 fiscal year:").font = LABEL_BOLD
    ws.cell(broker_layout.ROW_FY2_YEAR, 1, "FY2 fiscal year:").font = LABEL_BOLD
    ws.cell(broker_layout.ROW_FY3_YEAR, 1, "FY3 fiscal year:").font = LABEL_BOLD

    for j, h in enumerate(broker_layout.COL_HEADERS):
        style_header(ws.cell(broker_layout.ROW_COL_HEADERS, j + 1, h))

    # Section A labels (data populated by fetch script)
    for r, label, *_ in broker_layout.PNL:
        ws.cell(r, 1, label)

    # Section B: implied calcs as formulas (live in this workbook)
    for r, label, formula in broker_layout.IMPLIED:
        ws.cell(r, 1, label)
        c = ws.cell(r, 2, formula)
        style_formula(c, num_format=PCT)

    # Section C: sentiment labels; B30 is a formula here
    for r, label, _ in broker_layout.SENTIMENT:
        ws.cell(r, 1, label)
    for r, formula in broker_layout.SENTIMENT_FORMULAS_IN_MODEL.items():
        c = ws.cell(r, 2, formula)
        style_formula(c, num_format=PCT)

    ws.column_dimensions["A"].width = 38
    for letter in ["B", "C", "D", "E", "F", "G"]:
        ws.column_dimensions[letter].width = 14


def register_driver_named_ranges(wb):
    """Driver assumption rows on Inputs (E..N for the 10 projection years)."""
    drv_ranges = {
        "drv_revenue_growth": DRV_REV,
        "drv_gross_margin":   DRV_GM,
        "drv_opex_pct_rev":   DRV_OPEX,
        "drv_capex_pct_rev":  DRV_CAPEX,
        "drv_da_pct_capex":   DRV_DA,
        "drv_exit_multiple":  DRV_EXIT,
    }
    for name, row in drv_ranges.items():
        add_named_range(wb, name, "Inputs", f"$E${row}:$N${row}")


def build():
    wb = Workbook()

    # Pre-enable iterative calc so the interest <-> taxes <-> cash sweep
    # circularity converges when the workbook is opened.
    wb.calculation.iterate = True
    wb.calculation.iterateCount = 100
    wb.calculation.iterateDelta = 0.001

    years = make_year_columns()

    wb.remove(wb.active)
    cover_ws = wb.create_sheet("Cover")
    inputs_ws = wb.create_sheet("Inputs")
    is_ws = wb.create_sheet("IS")
    cf_ws = wb.create_sheet("CF")
    debt_ws = wb.create_sheet("Debt")
    val_ws = wb.create_sheet("Valuation")
    summary_ws = wb.create_sheet("Summary")
    sens_ws = wb.create_sheet("Sensitivity")
    capiq_ws = wb.create_sheet("_CapIQ_Data")
    broker_ws = wb.create_sheet("_Broker_Data")

    build_inputs(inputs_ws, years)
    build_cover(cover_ws, years)
    build_is(is_ws, years)
    build_cf(cf_ws, years)
    build_debt(debt_ws, years)
    build_valuation(val_ws, years)
    build_summary(summary_ws, years)
    build_sensitivity(sens_ws, years)
    build_capiq_data(capiq_ws)
    build_broker_data(broker_ws)

    register_inputs_named_ranges(wb)
    register_driver_named_ranges(wb)
    add_named_range(wb, "sens_target_year_offset", "Sensitivity", "$B$2")

    TEMPLATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(TEMPLATE_PATH)
    return TEMPLATE_PATH


if __name__ == "__main__":
    out = build()
    print(f"Wrote {out}")
