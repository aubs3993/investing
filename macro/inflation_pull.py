from datetime import datetime, timedelta
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

end = datetime.today()
chart_start = end - timedelta(days=365 * 20)
# Pull ~13 months of history before chart_start so YoY is defined at the chart's
# left edge rather than starting blank for the first year.
data_start = chart_start - timedelta(days=400)

series = {
    "CPI_headline": "CPIAUCSL",
    "CPI_core": "CPILFESL",
    "PCE_headline": "PCEPI",
    "PCE_core": "PCEPILFE",
}

df = pull_series(fred, series, data_start, end)
df = df.dropna(subset=list(series.keys())).reset_index(drop=True)

# Monthly indices: pct_change(12) gives YoY %.
yoy_cols = []
for name in series:
    yoy_col = f"{name}_yoy"
    df[yoy_col] = df[name].pct_change(12) * 100
    yoy_cols.append(yoy_col)

df = df.dropna(subset=yoy_cols).reset_index(drop=True)
# Trim to the chart window now that YoY is computed; xlsx and charts share this range.
df = df[df["Date"] >= pd.Timestamp(chart_start)].reset_index(drop=True)

summary = pd.DataFrame({
    "min": df[yoy_cols].min(),
    "max": df[yoy_cols].max(),
    "mean": df[yoy_cols].mean(),
    "most_recent": df[yoy_cols].iloc[-1],
})

OUT_DIR = resolve_output_dir(__file__, "inflation")
xlsx_path = OUT_DIR / "inflation.xlsx"
with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Data", index=False)
    summary.to_excel(writer, sheet_name="Summary")

recessions = get_recession_periods(fred, chart_start, end)

# Headline vs. core distinguished by color + weight, both solid.
HEADLINE_COLOR = "#1f77b4"  # blue, lighter weight
CORE_COLOR = "#d62728"      # red, heavier weight — core is the cleaner signal
HEADLINE_LW = 1.2
CORE_LW = 2.0
YLIM = (-2, 10)
XLIM = (pd.Timestamp(chart_start), pd.Timestamp(end))
TITLE_RANGE = f"{chart_start.year}–present"

fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df["Date"], df["CPI_headline_yoy"], color=HEADLINE_COLOR,
        linewidth=HEADLINE_LW, label="Headline CPI")
ax.plot(df["Date"], df["CPI_core_yoy"], color=CORE_COLOR,
        linewidth=CORE_LW, label="Core CPI")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title=f"CPI year-over-year, {TITLE_RANGE}",
    ylabel="YoY % change",
    ylim=YLIM,
    recessions=recessions,
    hline=2.0,
    hline_label="Fed target (PCE basis)",
)
fig.tight_layout()
fig.savefig(OUT_DIR / "cpi_yoy.png", dpi=150)
plt.close(fig)

fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df["Date"], df["PCE_headline_yoy"], color=HEADLINE_COLOR,
        linewidth=HEADLINE_LW, label="Headline PCE")
ax.plot(df["Date"], df["PCE_core_yoy"], color=CORE_COLOR,
        linewidth=CORE_LW, label="Core PCE")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title=f"PCE year-over-year, {TITLE_RANGE}",
    ylabel="YoY % change",
    ylim=YLIM,
    recessions=recessions,
    hline=2.0,
    hline_label=None,
)
fig.tight_layout()
fig.savefig(OUT_DIR / "pce_yoy.png", dpi=150)
plt.close(fig)

print(f"Start date:  {df['Date'].iloc[0].date()}")
print(f"End date:    {df['Date'].iloc[-1].date()}")
print(f"Rows:        {len(df)}")
for col in yoy_cols:
    print(f"Latest {col}: {df[col].iloc[-1]:.2f}%")
print(f"Wrote {xlsx_path.name}, cpi_yoy.png, pce_yoy.png to {OUT_DIR}")
