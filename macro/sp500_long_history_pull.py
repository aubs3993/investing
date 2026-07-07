# Canonical long-run S&P Composite dataset, monthly Jan 1871 - present, built
# from Robert Shiller's ie_data.xls (shillerdata.com): price, dividends,
# earnings, CPI, a constructed nominal total-return index, real total return,
# and drawdowns.
#
# Source / citation: Robert J. Shiller, "Irrational Exuberance" (Princeton
# University Press). Shiller requests that users of this data cite the book.
#
# DATA CONTRACT — sp500_monthly.csv: the credit-cycle script consumes this
# file. Columns (exact names, do not rename): Date, P, D, E, CPI, TR, TR_real,
# drawdown. Date is YYYY-MM-DD (first of month); TR is the nominal
# total-return index (base 100 at 1871-01); TR_real is the real total-return
# index (base 100 at 1871-01); drawdown is TR/cummax(TR) - 1 (nominal basis,
# decimal, <= 0).
#
# INTENTIONAL EXCEPTION to the repo's 2006-01-01 start convention: the whole
# point of this series is the 150+ year return/drawdown history, so the
# primary charts cover 1871-present. A companion 2006+ drawdown chart
# (sp500_drawdown_2006.png) is emitted so it slots into the comparable-axis
# macro chart set.
#
# CAVEATS on the underlying data:
# - Shiller's P is the monthly AVERAGE of daily closes, not month-end, so
#   drawdowns are smoothed vs daily reality (Oct 1987 looks mild here; the GFC
#   trough is ~-50% on monthly averages vs -57% on daily closes). A daily
#   overlay via Yahoo/CapIQ is deliberately deferred — flagged as future work.
# - Dividends (D) are annual rates interpolated monthly from quarterly data;
#   the monthly cash dividend used in the TR construction is D/12.
# - The last month is partial (price = most recent close / partial-month
#   average) and the last 1-2 rows have price but no dividend yet; trailing
#   D-NaN rows are DROPPED from the dataset so TR is always defined.
#
# GOTCHA — Date column: dates are floats where the fractional part is the
# literal month string, so 1871.1 means OCTOBER (".10"), not January. We
# format with exactly 2 decimals and split on the decimal point (same approach
# as macro/shiller_pe_pull.py).
#
# GOTCHA — column names: the "Data" sheet has two "Price" headers; pandas
# dedups them, so "Price" = Real Price and "Price.1" = Real Total Return
# Price (both in latest-month CPI dollars). The <1% cross-check against the
# constructed nominal TR index doubles as validation that "Price.1" is still
# the real-TR column.
#
# Download: reuses the ie_data.xls already fetched by macro/shiller_pe_pull.py
# when present (copied into this topic's folder so it is self-contained);
# otherwise scrapes https://shillerdata.com/ for the current blob href with a
# last-known-URL fallback. FRED is used ONLY for recession shading.
from datetime import datetime
from pathlib import Path
import re
import shutil
import sys

import matplotlib.pyplot as plt
import pandas as pd
import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.fred_helpers import (
    get_fred_client,
    get_recession_periods,
    resolve_output_dir,
    series_stats,
    style_macro_chart,
)

SHILLER_PAGE = "https://shillerdata.com/"
# Last-known direct blob URL (verified 2026-07-07). The ?ver= token rotates
# whenever Shiller updates the file, so this is a fallback only.
FALLBACK_XLS_URL = (
    "https://img1.wsimg.com/blobby/go/e5e77e0b-59d1-44d9-ab25-4763ac982e53/"
    "downloads/dd48d685-0157-4aa8-9ad3-375fd4eef22b/ie_data.xls?ver=1783022873468"
)
# The site 403s default python-requests UAs; present a browser UA.
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"
    )
}

OUT_DIR = resolve_output_dir(__file__, "sp500_long_history")
(OUT_DIR / ".gitkeep").touch()

end = datetime.today()

# --- Obtain ie_data.xls (prefer the copy shiller_pe_pull.py already fetched) ---
raw_path = OUT_DIR / "ie_data.xls"  # raw .xls kept for reference; gitignored
shiller_pe_copy = OUT_DIR.parent / "shiller_pe" / "ie_data.xls"
if shiller_pe_copy.exists():
    shutil.copyfile(shiller_pe_copy, raw_path)
    print(f"Reused ie_data.xls from {shiller_pe_copy.parent.name}/")
elif raw_path.exists():
    print("Reusing existing ie_data.xls in this topic's folder")
else:
    xls_url = FALLBACK_XLS_URL
    try:
        page = requests.get(SHILLER_PAGE, headers=HEADERS, timeout=60)
        page.raise_for_status()
        hrefs = re.findall(r'href="([^"]*ie_data\.xls[^"]*)"', page.text)
        if hrefs:
            xls_url = hrefs[0]
            if xls_url.startswith("//"):
                xls_url = "https:" + xls_url
        else:
            print("WARNING: ie_data.xls link not found on page; using fallback URL")
    except requests.RequestException as exc:
        print(f"WARNING: could not scrape {SHILLER_PAGE} ({exc}); using fallback URL")
    resp = requests.get(xls_url, headers=HEADERS, timeout=120)
    resp.raise_for_status()
    raw_path.write_bytes(resp.content)
    print("Downloaded ie_data.xls from shillerdata.com")

# --- Parse ---
# Header block is messy multi-row; skiprows=7 lands on the effective header.
raw = pd.read_excel(raw_path, sheet_name="Data", skiprows=7, engine="xlrd")
needed = ["Date", "P", "D", "E", "CPI", "Price", "Price.1"]
missing = [c for c in needed if c not in raw.columns]
if missing:
    raise RuntimeError(
        f"ie_data.xls layout changed; missing columns {missing}. "
        f"Got: {list(raw.columns)}"
    )

df = raw[needed].copy()
# Coerce to numeric: the final row is a text footnote (Date NaN, P a string).
for col in needed:
    df[col] = pd.to_numeric(df[col], errors="coerce")
df = df.dropna(subset=["Date"]).reset_index(drop=True)

# Date floats: fractional part is the literal month, so 1871.1 == October.
date_str = df["Date"].map(lambda d: f"{d:.2f}")
year = date_str.str.split(".").str[0].astype(int)
month = date_str.str.split(".").str[1].astype(int)
df["Date"] = pd.to_datetime({"year": year, "month": month, "day": 1})
df = df.rename(columns={"Price": "P_real", "Price.1": "TR_real_price"})

# Latest raw price month (before the D-NaN tail is dropped) for the summary.
latest_price_row = df.dropna(subset=["P"]).iloc[-1]

# Drop trailing rows where D is NaN (last 1-2 months have price only) so the
# TR index is defined on every remaining row. Interior D gaps would be a data
# problem, not a tail artifact — fail loudly if any remain.
df = df.loc[: df["D"].last_valid_index()].reset_index(drop=True)
if df["D"].isna().any():
    raise RuntimeError("Interior NaNs in D after trimming the tail — inspect ie_data.xls")

# --- Construct nominal total-return index, base 100 at 1871-01 ---
# TR_t = TR_{t-1} * (P_t + D_t/12) / P_{t-1}: buy at last month's average
# price, collect one month of the annual-rate dividend, reinvest.
growth = (df["P"] + df["D"] / 12.0) / df["P"].shift(1)
growth.iloc[0] = 1.0
df["TR"] = 100.0 * growth.cumprod()

# Cross-check: Shiller's Real Total Return Price times CPI, rescaled to the
# same base, should reproduce the nominal TR index. This validates both our
# cumprod construction and the "Price.1" column identification.
nominal_check = df["TR_real_price"] * df["CPI"]
nominal_check = 100.0 * nominal_check / nominal_check.iloc[0]
rel_diff = (df["TR"] - nominal_check).abs() / nominal_check
max_dev = float(rel_diff.max())
if max_dev >= 0.01:
    raise RuntimeError(
        f"TR cross-check failed: max relative deviation {max_dev:.4%} vs "
        "Real-TR-Price x CPI (limit 1%)"
    )

# Real TR index (base 100 at 1871-01) comes free from the file column.
df["TR_real"] = 100.0 * df["TR_real_price"] / df["TR_real_price"].iloc[0]

# Drawdown on the nominal TR basis.
df["drawdown"] = df["TR"] / df["TR"].cummax() - 1.0

# Annualized nominal total return over the full sample.
n_months = len(df) - 1
tr_annualized = (df["TR"].iloc[-1] / df["TR"].iloc[0]) ** (12.0 / n_months) - 1.0

# --- Outputs ---
# CSV data contract (see header). Exact column names; consumed by the
# credit-cycle script.
csv_cols = ["Date", "P", "D", "E", "CPI", "TR", "TR_real", "drawdown"]
csv_path = OUT_DIR / "sp500_monthly.csv"
df[csv_cols].to_csv(csv_path, index=False, date_format="%Y-%m-%d")

summary_rows = {
    c: series_stats(df[c]) for c in ["P", "D", "E", "CPI", "TR", "TR_real", "drawdown"]
}
summary = pd.DataFrame(summary_rows).T[["min", "max", "mean", "median", "current"]]
# Extra reference rows: full-sample annualized nominal TR and the cross-check
# deviation, carried in the "current" column (other stats blank).
summary.loc["tr_annualized_nominal"] = [pd.NA, pd.NA, pd.NA, pd.NA, tr_annualized]
summary.loc["tr_crosscheck_max_dev"] = [pd.NA, pd.NA, pd.NA, pd.NA, max_dev]

xlsx_path = OUT_DIR / "sp500_long_history.xlsx"
with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    df[csv_cols + ["P_real"]].to_excel(writer, sheet_name="Data", index=False)
    summary.to_excel(writer, sheet_name="Summary")

# --- Charts ---
fred = get_fred_client()  # FRED used only for USREC recession shading
full_start = df["Date"].iloc[0]  # 1871-01-01; USREC goes back to 1854
recessions_full = get_recession_periods(fred, full_start, end)

# Chart 1: nominal + real TR indices, log scale (150 years of compounding is
# unreadable on a linear axis).
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df["Date"], df["TR_real"], color="#9ec5e8", linewidth=1.2,
        label="Real total return (base 100 = 1871)")
ax.plot(df["Date"], df["TR"], color="#1f3b73", linewidth=1.5,
        label="Nominal total return (base 100 = 1871)")
ax.set_yscale("log")
ax.set_xlim(full_start, pd.Timestamp(end))
style_macro_chart(
    ax,
    title="S&P Composite total-return index, 1871–present (log scale)",
    ylabel="Index (log scale)",
    recessions=recessions_full,
)
fig.tight_layout()
fig.savefig(OUT_DIR / "sp500_tr_log.png", dpi=150)
plt.close(fig)

# Chart 2: drawdown from all-time high (nominal TR basis), 1871+.
dd_pct = df["drawdown"] * 100.0
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df["Date"], dd_pct, color="#1f3b73", linewidth=1.0,
        label="Drawdown (nominal TR, monthly avg prices)")
ax.fill_between(df["Date"], dd_pct, 0, color="#9ec5e8", alpha=0.5)
ax.set_xlim(full_start, pd.Timestamp(end))
style_macro_chart(
    ax,
    title="S&P Composite drawdown from all-time high, 1871–present",
    ylabel="Drawdown (%)",
    recessions=recessions_full,
    hlines=[
        {"y": -20.0, "label": "-20% (bear market)", "color": "#c0392b"},
        {"y": -50.0, "label": "-50%", "color": "#c0392b", "linestyle": ":"},
    ],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "sp500_drawdown.png", dpi=150)
plt.close(fig)

# Chart 3: 2006+ companion so it slots into the comparable-axis macro chart
# set. Same full-history drawdown series, filtered — by 2006 the nominal TR
# index was already past its 2000 peak, so this equals a 2006-rebased
# drawdown everywhere it matters (GFC, COVID, today).
chart_start_2006 = datetime(2006, 1, 1)
df06 = df[df["Date"] >= pd.Timestamp(chart_start_2006)].reset_index(drop=True)
recessions_2006 = get_recession_periods(fred, chart_start_2006, end)
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df06["Date"], df06["drawdown"] * 100.0, color="#1f3b73", linewidth=1.2,
        label="Drawdown (nominal TR, monthly avg prices)")
ax.fill_between(df06["Date"], df06["drawdown"] * 100.0, 0, color="#9ec5e8", alpha=0.5)
ax.set_xlim(pd.Timestamp(chart_start_2006), pd.Timestamp(end))
style_macro_chart(
    ax,
    title="S&P Composite drawdown from all-time high, 2006–present",
    ylabel="Drawdown (%)",
    recessions=recessions_2006,
    hlines=[
        {"y": -20.0, "label": "-20% (bear market)", "color": "#c0392b"},
        {"y": -50.0, "label": "-50%", "color": "#c0392b", "linestyle": ":"},
    ],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "sp500_drawdown_2006.png", dpi=150)
plt.close(fig)

# --- Print summary ---
# Era drawdown troughs (monthly-average basis — see smoothing caveat above).
eras = {
    "Depression": ("1929-01-01", "1940-12-31"),
    "Dot-com": ("2000-01-01", "2003-12-31"),
    "GFC": ("2007-10-01", "2009-12-31"),
    "COVID 2020": ("2020-01-01", "2020-12-31"),
}
era_troughs = {}
for name, (w_start, w_stop) in eras.items():
    win = df[(df["Date"] >= w_start) & (df["Date"] <= w_stop)]
    i = win["drawdown"].idxmin()
    era_troughs[name] = (df.loc[i, "Date"], df.loc[i, "drawdown"])

print(f"Data start:              {df['Date'].iloc[0].date()}")
print(f"Data end (TR, D avail):  {df['Date'].iloc[-1].date()}")
print(f"Latest price month:      {latest_price_row['Date'].date()} "
      f"(P = {latest_price_row['P']:,.2f}, partial month, dropped from TR rows)")
print(f"Rows (CSV/xlsx):         {len(df)}")
print(f"TR index (nominal):      {df['TR'].iloc[-1]:,.0f} (base 100 = 1871-01)")
print(f"TR index (real):         {df['TR_real'].iloc[-1]:,.0f} (base 100 = 1871-01)")
print(f"Annualized nominal TR:   {tr_annualized:.2%} over {n_months / 12:.0f} years")
print(f"TR cross-check max dev:  {max_dev:.4%} (vs Real-TR-Price x CPI; limit 1%)")
print(f"Max drawdown:            {df['drawdown'].min():.1%} "
      f"({df.loc[df['drawdown'].idxmin(), 'Date'].date()})")
for name, (t_date, t_dd) in era_troughs.items():
    print(f"{name + ' trough:':<25}{t_dd:.1%} ({t_date.date()})")
print(f"Current drawdown:        {df['drawdown'].iloc[-1]:.1%}")
print(f"Wrote {csv_path.name}, {xlsx_path.name}, ie_data.xls, sp500_tr_log.png, "
      f"sp500_drawdown.png, sp500_drawdown_2006.png to {OUT_DIR}")
