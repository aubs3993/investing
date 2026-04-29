"""Shared layout for the historical NTM EV/EBITDA multiple data flow.

Two files share this layout:
  - templates/multiple_history_fetcher.xlsx (Fetcher tab) — live CapIQ formulas
  - companies/output/<TICKER>/multiple_history_<TICKER>.xlsx — hardcoded values copy

Each row corresponds to a single US business day. Up to 1,500 data rows
(~6 years of trading days) — extend `ROW_DATA_END` if a longer lookback
is ever required, then rerun `shared.scaffold_multiple_history_fetcher`.

Universal conventions apply: column A is a 2.71-wide spacer, row 1 is
blank, B2 carries a dynamic title formula. Field labels live in column B
in the header section; the data grid runs B (date) through O (multiple).
"""
from __future__ import annotations

# --- Header rows ---
ROW_TITLE = 2
ROW_BANNER = 4
ROW_RUN_VIA = 5
ROW_TICKER = 6           # label B6, value C6 (named: mh_ticker)
ROW_END_DATE = 7         # label B7, value C7 (named: mh_end_date)
ROW_LOOKBACK_YRS = 8     # label B8, value C8 (named: mh_lookback_yrs)

# --- Column-header / data-grid rows ---
ROW_COL_HEADERS = 10
ROW_DATA_START = 11
ROW_DATA_END = 1510      # 1500 rows of capacity (~6 years of business days)

# Optional generated-timestamp row (used by the hardcoded copy only).
ROW_GENERATED = 9

# --- Data columns ---
# (letter, header, role)  role ∈ {"date", "input_formula", "calc"}
DATA_COLUMNS = [
    ("B", "Date",                   "date"),
    ("C", "Stock Price",            "input_formula"),
    ("D", "Diluted Shares Out",     "input_formula"),
    ("E", "Cash & Equivalents",     "input_formula"),
    ("F", "ST Investments",         "input_formula"),
    ("G", "Total Debt",             "input_formula"),
    ("H", "Preferred Equity",       "input_formula"),
    ("I", "Minority Interest",      "input_formula"),
    ("J", "Equity Investments",     "input_formula"),
    ("K", "Marketable Securities",  "input_formula"),
    ("L", "Market Cap",             "calc"),
    ("M", "Enterprise Value",       "calc"),
    ("N", "NTM EBITDA",             "input_formula"),
    ("O", "NTM EV/EBITDA Multiple", "calc"),
]

# Formula templates per input column. {t} = ticker named range,
# {d} = the row's date cell reference (e.g., "B11").
# Every CapIQ call is wrapped in IFERROR(..., "") so missing/erroring
# cells become blank rather than #N/A — keeps downstream copy + chart
# logic clean (cf. capiq_fetcher.xlsx for the underlying formula syntax).
INPUT_FORMULAS = {
    "C": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_LASTSALEPRICE", {d}), "")',
    "D": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_SHARESOUTSTANDING_OUT", {d}), "")',
    "E": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_CASH_EQUIV", , {d}), "")',
    "F": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_ST_INVEST", , {d}), "")',
    "G": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_TOTAL_DEBT_EXCL_OPER_LEASES", , {d}), "")',
    "H": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_PREF_EQUITY", , {d}), "")',
    "I": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_MINORITY_INTEREST", , {d}), "")',
    "J": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_EQUITY_METHOD", , {d}), "")',
    "K": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_LT_Mark_Securities", , {d}), "")',
    "N": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_EBITDA_EST", IQ_NTM, {d}, , , "USD"), "")',
}

# Calc templates. {r} = row number (e.g., 11).
CALC_FORMULAS = {
    "L": '=IFERROR(C{r}*D{r}, "")',
    "M": '=IFERROR(L{r}-E{r}-F{r}+G{r}+H{r}+I{r}-J{r}-K{r}, "")',
    "O": '=IFERROR(M{r}/N{r}, "")',
}

# Per-column number formats applied in both fetcher and hardcoded copy.
COLUMN_NUMBER_FORMATS = {
    "B": "mm/dd/yyyy",
    "C": "0.00",
    "D": "#,##0",
    "E": "#,##0;(#,##0);-",
    "F": "#,##0;(#,##0);-",
    "G": "#,##0;(#,##0);-",
    "H": "#,##0;(#,##0);-",
    "I": "#,##0;(#,##0);-",
    "J": "#,##0;(#,##0);-",
    "K": "#,##0;(#,##0);-",
    "L": "#,##0;(#,##0);-",
    "M": "#,##0;(#,##0);-",
    "N": "#,##0;(#,##0);-",
    "O": '0.0"x"',
}

COLUMN_WIDTHS = {
    "A": 2.71,   # spacer
    "B": 12,
    "C": 11,
    "D": 12,
    "E": 13, "F": 13, "G": 13, "H": 13, "I": 13, "J": 13, "K": 13,
    "L": 14,
    "M": 14,
    "N": 13,
    "O": 12,
}

# Named ranges (workbook-scoped).
NAME_TICKER = "mh_ticker"
NAME_END_DATE = "mh_end_date"
NAME_LOOKBACK_YRS = "mh_lookback_yrs"


def data_capacity() -> int:
    """Number of data rows the layout supports."""
    return ROW_DATA_END - ROW_DATA_START + 1


def col_letter(role_target: str) -> str | None:
    """Return the column letter of the first column with the given role."""
    for letter, _label, role in DATA_COLUMNS:
        if role == role_target:
            return letter
    return None
