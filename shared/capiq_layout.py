"""Shared layout for the CapIQ data flow.

Two workbooks must mirror this layout cell-for-cell:
  - templates/capiq_fetcher.xlsx (Fetcher tab)         — live CapIQ formulas
  - templates/company_model.xlsx (_CapIQ_Data tab)     — hardcoded values

Field labels in column B must match `templates/capiq_fetcher.xlsx` Fetcher
column B exactly at the same row positions (case-insensitive). The fetcher
is hand-maintained and is the source of truth; this module mirrors it so
shared/scaffold_template.py and shared/fetch_capiq.py stay aligned.

shared/fetch_capiq.py copies the used range from Fetcher to _CapIQ_Data
verbatim, so layout drift between the two will misalign data. Edit row
positions / field labels in this one file to keep them in sync.
"""
from __future__ import annotations

# --- Header rows ---
ROW_TITLE = 2          # B2 = =inp_ticker & " | CapIQ Data" (or " | CapIQ Fetcher" in fetcher)
ROW_BANNER = 4         # gray-italic instructions
ROW_RUN_VIA = 5        # "Run via:" / fetch command hint
ROW_TICKER = 6         # ticker label / value
ROW_DATE = 7           # "Date" / as-of date
ROW_FETCHER_DATE = 8   # "Fetcher Run-Date" / timestamp written by fetch_capiq
ROW_COL_HEADERS = 10   # blue header row

# --- Column positions ---
COL_FIELD_LABEL = 2    # B
COL_FY_M2 = 3          # C — IQ_FY-2
COL_FY_M1 = 4          # D — IQ_FY-1
COL_FY = 5             # E — IQ_FY (most recent completed)
COL_CURRENT = 6        # F — current / LTM / point-in-time

COL_HEADERS = ["Field", "IQ_FY-2", "IQ_FY-1", "IQ_FY", "Current"]

# Section A: Company metadata — single value per row in column F (current/LTM only).
# (row, label, capiq_formula_template). Kept for fetcher scaffolder; main
# template only uses (row, label).
METADATA = [
    (12, "Company Name",  "=IQ_COMPANY_NAME({t})"),
    (13, "Sector",        "=IQ_GICS_SECTOR({t})"),
    (14, "Currency",      "=IQ_FILING_CURRENCY({t})"),
    (15, "Filing Status", "=IQ_COMPANY_STATUS({t})"),
]

# Section B: Current state / point-in-time. Single value per row in column F.
# (row, label, is_calc). Two rows are computed in the fetcher itself:
#   row 20 Market Cap     = F18*F19
#   row 28 Enterprise Value = F20-F21-F22+F23+F24+F25-F26-F27
CURRENT_STATE = [
    (18, "Current Price",         False),
    (19, "Diluted Shares Out",    False),
    (20, "Market Cap",            True),
    (21, "Cash & Equivalents",    False),
    (22, "ST Investments",        False),
    (23, "Total Debt",            False),
    (24, "Preferred Equity",      False),
    (25, "Minority Interest",     False),
    (26, "Equity Investments",    False),
    (27, "Marketable Securities", False),
    (28, "Enterprise Value",      True),
]

# Section C: Historical time-series. Three values per row in columns C/D/E
# (FY-2, FY-1, FY). Column F (Current) is unused — not all historicals have
# a meaningful current/LTM value.
HISTORICALS = [
    (31, "Revenue"),
    (32, "COGS"),
    (33, "Gross Profit"),
    (34, "Total Opex"),
    (35, "D&A"),
    (36, "EBITDA"),
    (37, "EBIT"),
    (38, "Capex"),
    (39, "SBC"),
    (40, "DPS"),
]

HIST_COLS = ["C", "D", "E"]                  # FY-2, FY-1, FY
HIST_PERIODS = ["IQ_FY-2", "IQ_FY-1", "IQ_FY"]
CURRENT_COL = "F"


def all_field_rows():
    """Return [(row, label), ...] across all three sections in row order.

    Used by shared/fetch_capiq.py for cross-validation between the Fetcher
    sheet and the _CapIQ_Data sheet.
    """
    rows = []
    for r, label, _ in METADATA:
        rows.append((r, label))
    for r, label, _ in CURRENT_STATE:
        rows.append((r, label))
    for r, label in HISTORICALS:
        rows.append((r, label))
    return rows


def last_used_row():
    return max(r for r, _ in all_field_rows())
