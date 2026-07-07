# Long-run S&P 500 valuation from Robert Shiller's ie_data.xls (shillerdata.com).
#
# Source / citation: Robert J. Shiller, "Irrational Exuberance" (Princeton
# University Press). Shiller requests that users of this data cite the book.
# The dataset is monthly, Jan 1871 to present: S&P composite price (P),
# trailing-12m as-reported earnings monthly-interpolated (E), CPI, CAPE, TR CAPE.
#
# INTENTIONAL EXCEPTION to the repo's 2006-01-01 start convention: the whole
# point of this series is the 150+ year valuation history, so the primary
# charts cover 1871/1881–present. A companion 2006+ chart (shiller_pe_2006.png)
# is also emitted so it slots into the comparable-axis macro chart set.
#
# GOTCHA — Date column: dates are floats where the fractional part is the
# literal month string, so 1871.1 means OCTOBER (".10"), not January. Naive
# float arithmetic corrupts every October. We format with exactly 2 decimals
# and split on the decimal point to recover year/month.
#
# GOTCHA — earnings tail: E lags price by roughly two quarters (and the most
# recent months are S&P estimate-based interpolations). The last row(s) have
# P and CPI but NaN E; we do NOT forward-fill — earnings-based series simply
# end earlier than price-based ones. Because the E tail is interpolated toward
# estimated future quarters, the latest trailing P/E here runs ~5-7 points
# BELOW last-reported-TTM sources (multpl/WSJ) — expected, not a bug.
#
# Download: the ie_data.xls link on https://shillerdata.com/ points at an
# img1.wsimg.com blob URL whose ?ver= token rotates on every data update, so
# we scrape the page for the current href and fall back to the last-known
# blob URL if scraping fails. FRED is used ONLY for recession shading.
from datetime import datetime
from pathlib import Path
import re
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
# whenever Shiller updates the file, so this is a fallback only — it keeps
# working for a while after rotation but may eventually serve stale data.
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

OUT_DIR = resolve_output_dir(__file__, "shiller_pe")
(OUT_DIR / ".gitkeep").touch()

end = datetime.today()

# --- Download ie_data.xls (scrape current link, fall back to known blob URL) ---
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
raw_path = OUT_DIR / "ie_data.xls"  # raw .xls kept for reference; gitignored
raw_path.write_bytes(resp.content)

# --- Parse ---
# Header block is messy multi-row; skiprows=7 lands on the effective header.
# Verified column names as read by pandas: Date, P, D, E, CPI, Fraction,
# Rate GS10, ..., CAPE, ..., TR CAPE. Select by name and fail loudly if the
# layout ever shifts.
raw = pd.read_excel(raw_path, sheet_name="Data", skiprows=7, engine="xlrd")
needed = ["Date", "P", "E", "CPI", "CAPE", "TR CAPE"]
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
# Format with exactly 2 decimals and split — never treat the fraction as a
# numeric month.
date_str = df["Date"].map(lambda d: f"{d:.2f}")
year = date_str.str.split(".").str[0].astype(int)
month = date_str.str.split(".").str[1].astype(int)
df["Date"] = pd.to_datetime({"year": year, "month": month, "day": 1})
df = df.rename(columns={"TR CAPE": "TR_CAPE"})

# Derived: trailing P/E on as-reported 12m earnings. E is NaN in the last
# row(s) (earnings lag), so PE ends earlier than P — intentionally not filled.
df["PE"] = df["P"] / df["E"]

# --- Summary stats + percentile ranks (full-history window = charted window) ---
summary_rows = {
    "P": series_stats(df["P"]),
    "E": series_stats(df["E"]),
    "CPI": series_stats(df["CPI"]),
    "PE": series_stats(df["PE"]),
    "CAPE": series_stats(df["CAPE"]),
    "TR_CAPE": series_stats(df["TR_CAPE"]),
}
summary = pd.DataFrame(summary_rows).T[["min", "max", "mean", "median", "current"]]
# Percentile rank of the current value within each series' own full history.
pct_ranks = {
    name: df[name].dropna().rank(pct=True).iloc[-1] * 100
    for name in ["PE", "CAPE", "TR_CAPE"]
}
summary["pct_rank_current"] = pd.Series(pct_ranks)

xlsx_path = OUT_DIR / "shiller_pe.xlsx"
with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    df[["Date", "P", "E", "CPI", "PE", "CAPE", "TR_CAPE"]].to_excel(
        writer, sheet_name="Data", index=False
    )
    summary.to_excel(writer, sheet_name="Summary")

# --- Charts ---
fred = get_fred_client()  # FRED used only for USREC recession shading
full_start = df["Date"].iloc[0]  # 1871-01-01; USREC goes back to 1854
recessions_full = get_recession_periods(fred, full_start, end)

pe_median = df["PE"].median()
cape_median = df["CAPE"].median()
tr_cape_median = df["TR_CAPE"].median()

# Chart 1: trailing P/E + CAPE, 1871–present, full-window median hlines.
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df["Date"], df["PE"], color="#9ec5e8", linewidth=1.0,
        label="Trailing P/E (as-reported)")
ax.plot(df["Date"], df["CAPE"], color="#1f3b73", linewidth=1.5, label="CAPE")
ax.set_xlim(full_start, pd.Timestamp(end))
style_macro_chart(
    ax,
    title="S&P 500 valuation — trailing P/E and CAPE, 1871–present",
    ylabel="Multiple (x)",
    # Clipped at 60x: the 2009 as-reported-earnings collapse spiked trailing
    # P/E to ~120, which would flatten 150 years of history. The spike
    # visibly runs off-axis by design.
    ylim=(0, 60),
    recessions=recessions_full,
    hlines=[
        {"y": pe_median, "label": f"P/E median ({pe_median:.1f}x)",
         "color": "#9ec5e8"},
        {"y": cape_median, "label": f"CAPE median ({cape_median:.1f}x)",
         "color": "#1f3b73"},
    ],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "shiller_pe_full.png", dpi=150)
plt.close(fig)

# Chart 2: CAPE + TR CAPE, 1881–present (CAPE needs 10y of earnings history).
cape_start = df.loc[df["CAPE"].notna(), "Date"].iloc[0]
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df["Date"], df["CAPE"], color="#1f3b73", linewidth=1.5, label="CAPE")
ax.plot(df["Date"], df["TR_CAPE"], color="#9ec5e8", linewidth=1.5,
        label="TR CAPE (total-return)")
ax.set_xlim(cape_start, pd.Timestamp(end))
style_macro_chart(
    ax,
    title="S&P 500 — CAPE and total-return CAPE, 1881–present",
    ylabel="Multiple (x)",
    recessions=recessions_full,
    hlines=[
        {"y": cape_median, "label": f"CAPE median ({cape_median:.1f}x)",
         "color": "#1f3b73"},
        {"y": tr_cape_median, "label": f"TR CAPE median ({tr_cape_median:.1f}x)",
         "color": "#9ec5e8"},
    ],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "shiller_cape_full.png", dpi=150)
plt.close(fig)

# Chart 3: 2006+ companion so this slots into the comparable-axis chart set.
# Reference medians recomputed over the 2006+ window (what's on the chart).
chart_start_2006 = datetime(2006, 1, 1)
df06 = df[df["Date"] >= pd.Timestamp(chart_start_2006)].reset_index(drop=True)
recessions_2006 = get_recession_periods(fred, chart_start_2006, end)
pe_median_06 = df06["PE"].median()
cape_median_06 = df06["CAPE"].median()
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df06["Date"], df06["PE"], color="#9ec5e8", linewidth=1.2,
        label="Trailing P/E (as-reported)")
ax.plot(df06["Date"], df06["CAPE"], color="#1f3b73", linewidth=1.5, label="CAPE")
ax.set_xlim(pd.Timestamp(chart_start_2006), pd.Timestamp(end))
style_macro_chart(
    ax,
    title="S&P 500 valuation — trailing P/E and CAPE, 2006–present",
    ylabel="Multiple (x)",
    # Clipped at 60x so the 2009 trailing-P/E spike to ~120 (earnings
    # collapse) doesn't compress the rest of the window; spike runs off-axis.
    ylim=(0, 60),
    recessions=recessions_2006,
    hlines=[
        {"y": pe_median_06, "label": f"P/E median 2006+ ({pe_median_06:.1f}x)",
         "color": "#9ec5e8"},
        {"y": cape_median_06, "label": f"CAPE median 2006+ ({cape_median_06:.1f}x)",
         "color": "#1f3b73"},
    ],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "shiller_pe_2006.png", dpi=150)
plt.close(fig)

# --- Print summary ---
last_pe_row = df.dropna(subset=["PE"]).iloc[-1]
last_cape_row = df.dropna(subset=["CAPE"]).iloc[-1]
peak_pe_row = df.loc[df["PE"].idxmax()]
peak_cape_row = df.loc[df["CAPE"].idxmax()]
print(f"Data start:            {df['Date'].iloc[0].date()}")
print(f"Data end (price):      {df['Date'].iloc[-1].date()}")
print(f"Data end (earnings):   {last_pe_row['Date'].date()}")
print(f"Rows:                  {len(df)}")
print(f"Latest trailing P/E:   {last_pe_row['PE']:.1f}x "
      f"(pct rank {pct_ranks['PE']:.0f}%)")
print(f"Latest CAPE:           {last_cape_row['CAPE']:.1f}x "
      f"(pct rank {pct_ranks['CAPE']:.0f}%)")
print(f"Latest TR CAPE:        {last_cape_row['TR_CAPE']:.1f}x "
      f"(pct rank {pct_ranks['TR_CAPE']:.0f}%)")
print(f"P/E median (full):     {pe_median:.1f}x")
print(f"CAPE median (full):    {cape_median:.1f}x")
print(f"Peak trailing P/E:     {peak_pe_row['PE']:.1f}x "
      f"({peak_pe_row['Date'].date()})")
print(f"Peak CAPE:             {peak_cape_row['CAPE']:.1f}x "
      f"({peak_cape_row['Date'].date()})")
print(f"Wrote {xlsx_path.name}, ie_data.xls, shiller_pe_full.png, "
      f"shiller_cape_full.png, shiller_pe_2006.png to {OUT_DIR}")
