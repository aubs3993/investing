"""Shared layout for the CapIQ data flow.

Two workbooks must mirror this layout cell-for-cell:
  - templates/capiq_fetcher.xlsx (Fetcher tab)         — live CapIQ formulas
  - templates/company_model.xlsx (_CapIQ_Data tab)     — hardcoded values

shared/fetch_capiq.py copies the used range from Fetcher to _CapIQ_Data
verbatim, so layout drift between the two will misalign data. Edit row
positions / field labels in this one file to keep them in sync.
"""
from __future__ import annotations

# Header section (rows 1-5)
ROW_BANNER = 1
ROW_LAST_FETCH = 2
ROW_TICKER = 3
ROW_COL_HEADERS = 5

COL_HEADERS = ["Field", "FY-3", "FY-2", "FY-1", "Current/LTM"]

# Section A: Company metadata. Single value per row in column E.
# (row, field_label, capiq_formula_with_{t}_placeholder_for_ticker_ref)
METADATA = [
    (7,  "Company Name",   "=IQ_COMPANY_NAME({t})"),
    (8,  "Sector",         "=IQ_GICS_SECTOR({t})"),
    (9,  "Currency",       "=IQ_FILING_CURRENCY({t})"),
    (10, "Filing Status",  "=IQ_COMPANY_STATUS({t})"),
]

# Section B: Current state / point-in-time. Single value per row in column E.
CURRENT = [
    (13, "Current Price",           "=IQ_LASTSALEPRICE({t})"),
    (14, "Diluted Shares Out",      "=IQ_TOTAL_DILUT_SHARES_OUT({t})"),
    (15, "Quarterly DPS",           "=IQ_LAST_DIVID_QUARTERLY({t})"),
    (16, "Cash & Equivalents",      "=IQ_CASH_EQUIV({t},IQ_LP)"),
    (17, "Total Debt",              "=IQ_TOTAL_DEBT({t},IQ_LP)"),
    (18, "Minority Interest",       "=IQ_MINORITY_INTEREST({t},IQ_LP)"),
    (19, "Equity Investments",      "=IQ_LT_INVEST({t},IQ_LP)"),
    (20, "Effective Tax Rate",      "=IQ_EFFECT_TAX_RATE({t},IQ_LTM)"),
    (21, "Market Cap (CapIQ calc)", "=IQ_MARKETCAP({t})"),
    (22, "Net Debt (CapIQ calc)",   "=IQ_NET_DEBT({t},IQ_LP)"),
]

# Section C: Historical time-series. Three values per row in columns B/C/D.
# capiq_function = the IQ_* function name; columns B/C/D get IQ_FY-3/FY-2/FY-1.
# capiq_function = None means the row is computed from other rows on the tab.
HISTORICAL = [
    (25, "Revenue",                     "IQ_TOTAL_REV"),
    (26, "COGS",                        "IQ_COGS"),
    (27, "Gross Profit",                "IQ_GP"),
    (28, "SG&A",                        "IQ_SGA"),
    (29, "R&D Expense",                 "IQ_RD_EXP"),
    (30, "Total OpEx (SG&A + R&D)",     None),     # =SUM of rows 28+29
    (31, "D&A",                         "IQ_DA_SUPPL"),
    (32, "EBITDA",                      "IQ_EBITDA"),
    (33, "EBIT",                        "IQ_EBIT"),
    (34, "Interest Expense",            "IQ_INT_EXP"),
    (35, "Interest Income",             "IQ_INT_INC"),
    (36, "Pre-tax Income",              "IQ_EBT_INCL_UNUSUAL"),
    (37, "Taxes",                       "IQ_INC_TAX"),
    (38, "Net Income",                  "IQ_NI"),
    (39, "Diluted Weighted Avg Shares", "IQ_DILUT_WEIGHT"),
    (40, "CapEx",                       "IQ_CAPEX"),
]

HIST_COLS = ["B", "C", "D"]  # FY-3, FY-2, FY-1
HIST_PERIODS = ["IQ_FY-3", "IQ_FY-2", "IQ_FY-1"]
CURRENT_COL = "E"


def all_field_rows():
    """Return [(row, label), ...] across all three sections in row order."""
    rows = []
    for r, label, _ in METADATA:
        rows.append((r, label))
    for r, label, _ in CURRENT:
        rows.append((r, label))
    for r, label, _ in HISTORICAL:
        rows.append((r, label))
    return rows


def last_used_row():
    return max(r for r, _ in all_field_rows())
