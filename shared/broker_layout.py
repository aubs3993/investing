"""Shared layout for the broker estimates data flow.

Mirrors the convention in shared/capiq_layout.py: two workbooks must share
this layout cell-for-cell.

  - templates/broker_fetcher.xlsx (Fetcher tab)             — live CapIQ EST formulas
  - templates/company_model.xlsx  (_Broker_Data tab)        — hardcoded values + computed formulas

Edit row positions / field labels here only — both scaffolders import this module.
"""
from __future__ import annotations

# Header section
ROW_BANNER = 1
ROW_LAST_FETCH = 2
ROW_TICKER = 3
ROW_FY1_YEAR = 4
ROW_FY2_YEAR = 5
ROW_FY3_YEAR = 6
ROW_COL_HEADERS = 8

COL_HEADERS = ["Field", "FY1 Mean", "FY2 Mean", "FY3 Mean", "FY1 High", "FY1 Low", "FY1 # Est"]

# Section A: P&L estimates. Cols B/C/D = mean FY1/FY2/FY3, E/F = FY1 high/low, G = FY1 count.
# (row, label, mean_func, high_func, low_func, count_func, omit_currency_arg)
PNL = [
    (10, "Revenue",       "IQ_EST_REV",     "IQ_EST_REV_HIGH",    "IQ_EST_REV_LOW",    "IQ_EST_NUM_REV",    False),
    (11, "Gross Profit",  "IQ_EST_GP",      "IQ_EST_GP_HIGH",     "IQ_EST_GP_LOW",     "IQ_EST_NUM_GP",     False),
    (12, "EBITDA",        "IQ_EST_EBITDA",  "IQ_EST_EBITDA_HIGH", "IQ_EST_EBITDA_LOW", "IQ_EST_NUM_EBITDA", False),
    (13, "EBIT",          "IQ_EST_EBIT",    "IQ_EST_EBIT_HIGH",   "IQ_EST_EBIT_LOW",   "IQ_EST_NUM_EBIT",   False),
    (14, "Net Income",    "IQ_EST_NI",      "IQ_EST_NI_HIGH",     "IQ_EST_NI_LOW",     "IQ_EST_NUM_NI",     False),
    (15, "EPS (Diluted)", "IQ_EST_EPS",     "IQ_EST_EPS_HIGH",    "IQ_EST_EPS_LOW",    "IQ_EST_NUM_EPS",    True),
    (16, "CFO",           "IQ_EST_CFO",     "IQ_EST_CFO_HIGH",    "IQ_EST_CFO_LOW",    "IQ_EST_NUM_CFO",    False),
    (17, "CapEx",         "IQ_EST_CAPEX",   "IQ_EST_CAPEX_HIGH",  "IQ_EST_CAPEX_LOW",  "IQ_EST_NUM_CAPEX",  False),
]

# Section B: Implied growth/margins (computed as formulas in _Broker_Data;
# left blank in fetcher). Single value per row in column B.
# (row, label, formula_template_for_main_template_only)
# Formulas reference _CapIQ_Data D-column (FY-1 historical) for growth bases.
IMPLIED = [
    (20, "Revenue Growth FY1 %",  "=IFERROR(B10/_CapIQ_Data!D25-1,\"\")"),
    (21, "Revenue Growth FY2 %",  "=IFERROR(C10/B10-1,\"\")"),
    (22, "Revenue Growth FY3 %",  "=IFERROR(D10/C10-1,\"\")"),
    (23, "Gross Margin FY1 %",    "=IFERROR(B11/B10,\"\")"),
    (24, "EBITDA Margin FY1 %",   "=IFERROR(B12/B10,\"\")"),
    (25, "EBIT Margin FY1 %",     "=IFERROR(B13/B10,\"\")"),
]

# Section C: Analyst sentiment. Single value per row in column B.
# (row, label, formula_or_None)  formula_or_None is the CapIQ formula for fetcher;
# main template just has the label and the fetched value (or computed for B30).
SENTIMENT = [
    (28, "Number of Analysts Covering",       "=IQ_EST_NUM_REV({t},IQ_FY1)"),
    (29, "Average Price Target",              "=IQ_PRICETARGET_AVG({t})"),
    (30, "Implied Upside %",                  None),  # =B29/inp_current_price-1, set in main template
    (31, "Average Recommendation (1=Buy,5=Sell)", "=IQ_RECOMMENDATION_AVG({t})"),
    (32, "Recommendation Distribution",       None),  # left blank for now
]

# In the main template, B30 is computed:
SENTIMENT_FORMULAS_IN_MODEL = {
    30: "=IFERROR(B29/inp_current_price-1,\"\")",
}

PERIODS = ["IQ_FY1", "IQ_FY2", "IQ_FY3"]
MEAN_COLS = ["B", "C", "D"]
HIGH_COL = "E"
LOW_COL = "F"
COUNT_COL = "G"


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
