"""Shared layout for the historical NTM EV/EBITDA multiple data flow.

Two files share this layout:
  - templates/multiple_history_fetcher.xlsx (Fetcher tab) — live CapIQ formulas
  - companies/output/<TICKER>/multiple_history_<TICKER>.xlsx — hardcoded values copy

Each row corresponds to a single US business day. Up to 1,500 data rows
(~6 years of trading days) — extend `ROW_DATA_END` if a longer lookback
is ever required, then rerun `shared.scaffold_multiple_history_fetcher`.

Universal conventions apply: column A is a 2.71-wide spacer, row 1 is
blank, B2 carries a dynamic title formula. Field labels live in column B
in the header section; the data grid runs B (date) through X (2Y fwd
growth %).
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
    ("B", "Date",                      "date"),
    ("C", "Stock Price",               "input_formula"),
    ("D", "Diluted Shares Out",        "input_formula"),
    ("E", "Cash & Equivalents",        "input_formula"),
    ("F", "ST Investments",            "input_formula"),
    ("G", "Total Debt",                "input_formula"),
    ("H", "Preferred Equity",          "input_formula"),
    ("I", "Minority Interest",         "input_formula"),
    ("J", "Equity Investments",        "input_formula"),
    ("K", "Marketable Securities",     "input_formula"),
    ("L", "Market Cap",                "calc"),
    ("M", "Enterprise Value",          "calc"),
    ("N", "IQ_CY EBITDA",              "input_formula"),
    ("O", "IQ_CY+1 EBITDA",            "input_formula"),
    ("P", "IQ_CY+2 EBITDA",            "input_formula"),
    ("Q", "IQ_CY+3 EBITDA",            "input_formula"),
    ("R", "LTM EBITDA",                "calc"),
    ("S", "NTM EBITDA",                "calc"),
    ("T", "NTM+12 EBITDA",             "calc"),
    ("U", "NTM EV/EBITDA",             "calc"),
    ("V", "2Y Fwd EV/EBITDA",          "calc"),
    ("W", "NTM Growth %",              "calc"),
    ("X", "2Y Fwd Growth % (CAGR)",    "calc"),
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
    "N": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_EBITDA_EST", IQ_CY, {d}, , , "USD"), "")',
    "O": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_EBITDA_EST", IQ_CY+1, {d}, , , "USD"), "")',
    "P": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_EBITDA_EST", IQ_CY+2, {d}, , , "USD"), "")',
    "Q": '=IFERROR(_xll.ciqfunctions.udf.CIQ({t}, "IQ_EBITDA_EST", IQ_CY+3, {d}, , , "USD"), "")',
}

# Calc templates. {r} = row number (e.g., 11).
# LTM/NTM/NTM+12 use weighted CY blends keyed off the row's date in column B:
#   frac_remaining = (DATE(YEAR(B)+1,1,1)-B)/365 — fraction of current CY ahead of B
#   frac_elapsed   = 1 - frac_remaining
# Symmetric construction so growth ratios (NTM/LTM, etc.) are clean.
CALC_FORMULAS = {
    "L": '=IFERROR(C{r}*D{r}, "")',
    "M": '=IFERROR(L{r}-E{r}-F{r}+G{r}+H{r}+I{r}-J{r}-K{r}, "")',
    "R": '=IFERROR(((DATE(YEAR(B{r})+1,1,1)-B{r})/365)*N{r}+(1-(DATE(YEAR(B{r})+1,1,1)-B{r})/365)*O{r}, "")',
    "S": '=IFERROR(((DATE(YEAR(B{r})+1,1,1)-B{r})/365)*O{r}+(1-(DATE(YEAR(B{r})+1,1,1)-B{r})/365)*P{r}, "")',
    "T": '=IFERROR(((DATE(YEAR(B{r})+1,1,1)-B{r})/365)*P{r}+(1-(DATE(YEAR(B{r})+1,1,1)-B{r})/365)*Q{r}, "")',
    "U": '=IFERROR(M{r}/S{r}, "")',
    "V": '=IFERROR(M{r}/T{r}, "")',
    "W": '=IFERROR(S{r}/R{r}-1, "")',
    "X": '=IFERROR((T{r}/R{r})^(1/2)-1, "")',
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
    "O": "#,##0;(#,##0);-",
    "P": "#,##0;(#,##0);-",
    "Q": "#,##0;(#,##0);-",
    "R": "#,##0;(#,##0);-",
    "S": "#,##0;(#,##0);-",
    "T": "#,##0;(#,##0);-",
    "U": '0.0"x"',
    "V": '0.0"x"',
    "W": "0.0%",
    "X": "0.0%",
}

COLUMN_WIDTHS = {
    "A": 2.71,   # spacer
    "B": 12,
    "C": 11,
    "D": 12,
    "E": 13, "F": 13, "G": 13, "H": 13, "I": 13, "J": 13, "K": 13,
    "L": 14,
    "M": 14,
    "N": 13, "O": 13, "P": 13, "Q": 13,
    "R": 13, "S": 13, "T": 13,
    "U": 12, "V": 12,
    "W": 11, "X": 11,
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
