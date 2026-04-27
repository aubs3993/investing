# NOTE on data availability:
# In April 2026, ICE restricted BAMLH0A0HYM2 (and the rest of the BofA OAS
# family) on FRED to a rolling 3-year window. The FRED series notes read:
# "Starting in April 2026, this series will only include 3 years of
# observations. For more data, go to the source."
#
# As a result the median and percentile-rank computed here reflect only the
# trailing ~3 years that FRED still serves, not a 2006+ long-term window.
# To build our own longer history independent of ICE's restriction, each run
# appends the pulled rows to macro/output/credit_spreads/hy_oas_archive.csv
# (deduped by date, latest pull wins on revisions). That CSV is tracked in
# git despite the macro/output/**/*.xlsx-style ignore rule — see .gitignore
# for the explicit override.

from datetime import datetime
from pathlib import Path
import sys

import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.fred_helpers import (
    get_fred_client,
    get_recession_periods,
    pull_series,
    resolve_output_dir,
    style_macro_chart,
)

fred = get_fred_client()

# Requested window: 2006-01-01 to today. Effective window will be ~3 years
# due to the FRED restriction noted above.
start = datetime(2006, 1, 1)
end = datetime.today()

series = {
    "HY_OAS": "BAMLH0A0HYM2",
}

df = pull_series(fred, series, start, end)
df = df.dropna(subset=["HY_OAS"]).reset_index(drop=True)

OUT_DIR = resolve_output_dir(__file__, "credit_spreads")
archive_path = OUT_DIR / "hy_oas_archive.csv"

# Merge today's pull into the persistent archive. The archive accumulates daily
# pulls so we can eventually compute statistics over a window longer than the
# 3 years FRED serves. New pull wins on dedupe so FRED revisions overwrite
# previously-archived values for the same date.
new_rows = df[["Date", "HY_OAS"]].copy()
if archive_path.exists():
    archive = pd.read_csv(archive_path, parse_dates=["Date"])
    combined = pd.concat([archive, new_rows], ignore_index=True)
else:
    combined = new_rows
combined = (
    combined.drop_duplicates(subset="Date", keep="last")
    .sort_values("Date")
    .reset_index(drop=True)
)
combined.to_csv(archive_path, index=False)

# Statistics are computed over the FRED-served window (the same df used for
# plotting), not the archive — keeps the chart and the xlsx Summary internally
# consistent for any single run.
window_start = df["Date"].iloc[0].date()
window_end = df["Date"].iloc[-1].date()
current = float(df["HY_OAS"].iloc[-1])
median_oas = float(df["HY_OAS"].median())
pct_rank = float((df["HY_OAS"] <= current).mean() * 100)

summary = pd.DataFrame({
    "value": [
        window_start.isoformat(),
        df["HY_OAS"].min(),
        df["HY_OAS"].max(),
        df["HY_OAS"].mean(),
        median_oas,
        current,
        pct_rank,
    ],
}, index=["window_start", "min", "max", "mean", "median", "current", "pct_rank"])

xlsx_path = OUT_DIR / "credit_spreads.xlsx"
with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Data", index=False)
    summary.to_excel(writer, sheet_name="Summary")

recessions = get_recession_periods(fred, df["Date"].min(), end)

fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df["Date"], df["HY_OAS"], color="#1f3b73", linewidth=1.2,
        label="ICE BofA US HY OAS")
ax.set_xlim(pd.Timestamp(window_start), pd.Timestamp(end))
style_macro_chart(
    ax,
    title="US High Yield OAS (3-year window — FRED restricted by ICE April 2026)",
    ylabel="OAS (percentage points)",
    recessions=recessions,
    hlines=[
        {"y": median_oas, "label": "Median (last 3 years)"},
        {"y": 8.0, "label": "Stress threshold", "color": "#c0392b"},
    ],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "hy_oas.png", dpi=150)
plt.close(fig)

print(f"Window start:      {window_start}")
print(f"Window end:        {window_end}")
print(f"Rows (this run):   {len(df)}")
print(f"Rows (archive):    {len(combined)}")
print(f"Latest HY OAS:     {current:.2f}")
print(f"Median (window):   {median_oas:.2f}")
print(f"Percentile rank:   {pct_rank:.1f}%")
print(f"Wrote {xlsx_path.name}, hy_oas.png, hy_oas_archive.csv to {OUT_DIR}")
