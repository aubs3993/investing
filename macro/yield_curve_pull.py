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

end = datetime.today()
start = datetime(2006, 1, 1)

series = {
    "10Y": "DGS10",
    "2Y": "DGS2",
    "FedFunds": "DFF",
}

df = pull_series(fred, series, start, end)
df["10Y_minus_2Y"] = df["10Y"] - df["2Y"]
# Require all three series for joint analysis. DFF typically lags DGS10/DGS2
# by ~1 trading day, so this drops the most recent rows where DFF is unpublished.
df = df.dropna(subset=["10Y", "2Y", "FedFunds"]).reset_index(drop=True)

summary_cols = ["10Y", "2Y", "FedFunds", "10Y_minus_2Y"]
summary = pd.DataFrame({
    "min": df[summary_cols].min(),
    "max": df[summary_cols].max(),
    "mean": df[summary_cols].mean(),
    "most_recent": df[summary_cols].iloc[-1],
})

OUT_DIR = resolve_output_dir(__file__, "yield_curve")
out_path = OUT_DIR / "yields.xlsx"
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Data", index=False)
    summary.to_excel(writer, sheet_name="Summary")

recessions = get_recession_periods(fred, start, end)

fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df["Date"], df["10Y_minus_2Y"], color="#1f3b73", linewidth=1.5,
        label="10Y-2Y spread")
ax.set_xlim(pd.Timestamp(start), pd.Timestamp(end))
style_macro_chart(
    ax,
    title=f"10Y-2Y Treasury spread, {start.year}–present",
    ylabel="Spread (percentage points)",
    ylim=(-1.5, 3.5),
    recessions=recessions,
    hlines=[{"y": 0.0, "label": "Inversion threshold"}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "spread_chart.png", dpi=150)
plt.close(fig)

print(f"Start date:  {df['Date'].iloc[0].date()}")
print(f"End date:    {df['Date'].iloc[-1].date()}")
print(f"Rows:        {len(df)}")
print(f"Latest 10Y-2Y spread: {df['10Y_minus_2Y'].iloc[-1]:.2f}")
print(
    f"Pulled {len(df)} rows. Note: rows are dropped if any of DGS10/DGS2/DFF "
    "is missing — DFF typically lags by 1 trading day, so the pull may end "
    "1 day before the chart's T10Y2Y series."
)
