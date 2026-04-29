"""Shared layout for the broker estimates data flow.

Mirrors the convention in shared/capiq_layout.py: two workbooks must share
this layout cell-for-cell.

  - templates/broker_fetcher.xlsx (Fetcher tab)             — live CapIQ EST formulas
  - templates/company_model.xlsx  (_Broker_Data tab)        — hardcoded values + computed formulas

All rows shifted +3 from the prior layout to make room for universal
conventions (row 1 blank, row 2 title, row 3 blank, row 4+ content).
Columns shifted +1 (column A is a 2.71-wide spacer).

Edit row positions / field labels here only — both scaffolders import this module.
"""
from __future__ import annotations

# Header section (universal: row 1 blank, B2 title, row 3 blank, content from row 4)
ROW_TITLE = 2
ROW_BANNER = 4
ROW_LAST_FETCH = 5      # was 2
ROW_TICKER = 6          # was 3
ROW_FY1_YEAR = 7        # was 4
ROW_FY2_YEAR = 8        # was 5
ROW_FY3_YEAR = 9        # was 6
ROW_COL_HEADERS = 11    # was 8

COL_HEADERS = ["Field", "FY1 Mean", "FY2 Mean", "FY3 Mean", "FY1 High", "FY1 Low", "FY1 # Est"]

# Section A: P&L estimates. Cols C/D/E = mean FY1/FY2/FY3, F/G = FY1 high/low, H = FY1 count.
# (row, label, mean_func, high_func, low_func, count_func, omit_currency_arg)
PNL = [
    (13, "Revenue",       "IQ_EST_REV",     "IQ_EST_REV_HIGH",    "IQ_EST_REV_LOW",    "IQ_EST_NUM_REV",    False),
    (14, "Gross Profit",  "IQ_EST_GP",      "IQ_EST_GP_HIGH",     "IQ_EST_GP_LOW",     "IQ_EST_NUM_GP",     False),
    (15, "EBITDA",        "IQ_EST_EBITDA",  "IQ_EST_EBITDA_HIGH", "IQ_EST_EBITDA_LOW", "IQ_EST_NUM_EBITDA", False),
    (16, "EBIT",          "IQ_EST_EBIT",    "IQ_EST_EBIT_HIGH",   "IQ_EST_EBIT_LOW",   "IQ_EST_NUM_EBIT",   False),
    (17, "Net Income",    "IQ_EST_NI",      "IQ_EST_NI_HIGH",     "IQ_EST_NI_LOW",     "IQ_EST_NUM_NI",     False),
    (18, "EPS (Diluted)", "IQ_EST_EPS",     "IQ_EST_EPS_HIGH",    "IQ_EST_EPS_LOW",    "IQ_EST_NUM_EPS",    True),
    (19, "CFO",           "IQ_EST_CFO",     "IQ_EST_CFO_HIGH",    "IQ_EST_CFO_LOW",    "IQ_EST_NUM_CFO",    False),
    (20, "CapEx",         "IQ_EST_CAPEX",   "IQ_EST_CAPEX_HIGH",  "IQ_EST_CAPEX_LOW",  "IQ_EST_NUM_CAPEX",  False),
]

# Section B: Implied growth/margins (computed as formulas in _Broker_Data;
# left blank in fetcher). Single value per row in column C (was B).
# Formulas reference _CapIQ_Data column E (IQ_FY = most recent completed year).
IMPLIED = [
    (23, "Revenue Growth FY1 %",  "=IFERROR(C13/_CapIQ_Data!E31-1,\"\")"),
    (24, "Revenue Growth FY2 %",  "=IFERROR(D13/C13-1,\"\")"),
    (25, "Revenue Growth FY3 %",  "=IFERROR(E13/D13-1,\"\")"),
    (26, "Gross Margin FY1 %",    "=IFERROR(C14/C13,\"\")"),
    (27, "EBITDA Margin FY1 %",   "=IFERROR(C15/C13,\"\")"),
    (28, "EBIT Margin FY1 %",     "=IFERROR(C16/C13,\"\")"),
]

# Section C: Analyst sentiment. Single value per row in column C (was B).
# (row, label, formula_or_None)  formula_or_None is the CapIQ formula for fetcher;
# main template just has the label and the fetched value (or computed for B33).
SENTIMENT = [
    (31, "Number of Analysts Covering",       "=IQ_EST_NUM_REV({t},IQ_FY1)"),
    (32, "Average Price Target",              "=IQ_PRICETARGET_AVG({t})"),
    (33, "Implied Upside %",                  None),  # computed in main template
    (34, "Average Recommendation (1=Buy,5=Sell)", "=IQ_RECOMMENDATION_AVG({t})"),
    (35, "Recommendation Distribution",       None),  # left blank for now
]

# In the main template, row 33 is computed:
SENTIMENT_FORMULAS_IN_MODEL = {
    33: "=IFERROR(C32/inp_current_price-1,\"\")",
}

PERIODS = ["IQ_FY1", "IQ_FY2", "IQ_FY3"]
MEAN_COLS = ["C", "D", "E"]   # was B/C/D
HIGH_COL = "F"                # was E
LOW_COL = "G"                 # was F
COUNT_COL = "H"               # was G


def all_field_rows():
    """Return [(row, label), ...] across all three sections, in row order."""
    rows = []
    for r, label, *_ in PNL:
        rows.append((r, label))
    for r, label, _ in IMPLIED:
        rows.append((r, label))
    for r, label, _ in SENTIMENT:
        rows.append((r, label))
    return rows
