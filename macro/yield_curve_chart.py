from datetime import datetime, timedelta
from dotenv import load_dotenv
from fredapi import Fred
from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import os
import pandas as pd

load_dotenv()

api_key = os.getenv("FRED_API_KEY")
if not api_key:
    raise RuntimeError("FRED_API_KEY not found in environment or .env file")

fred = Fred(api_key=api_key.strip())

end = datetime.today()
start = end - timedelta(days=365 * 20)

spread = fred.get_series("T10Y2Y", observation_start=start, observation_end=end)
recession = fred.get_series("USREC", observation_start=start, observation_end=end)

spread.name = "T10Y2Y"
recession.name = "USREC"

daily_idx = spread.index
recession_daily = recession.reindex(daily_idx, method="ffill")

df = pd.DataFrame({"T10Y2Y": spread, "USREC": recession_daily})
df.index.name = "Date"
df = df.dropna(subset=["T10Y2Y"])

rec_flag = df["USREC"].fillna(0).astype(int)
in_recession = rec_flag == 1
bands = []
band_start = None
for date, flag in in_recession.items():
    if flag and band_start is None:
        band_start = date
    elif not flag and band_start is not None:
        bands.append((band_start, date))
        band_start = None
if band_start is not None:
    bands.append((band_start, df.index[-1]))

fig, ax = plt.subplots(figsize=(12, 6))
ax.plot(df.index, df["T10Y2Y"], color="#1f4e79", linewidth=1.2, label="10Y-2Y Spread")
ax.axhline(0, color="red", linestyle="--", linewidth=1, label="Inversion threshold (0)")

for i, (b_start, b_end) in enumerate(bands):
    ax.axvspan(
        b_start, b_end,
        color="gray", alpha=0.3,
        label="NBER Recession" if i == 0 else None,
    )

ax.set_title("10Y-2Y Treasury Spread with NBER Recessions")
ax.set_xlabel("Date")
ax.set_ylabel("Spread (percentage points)")
ax.grid(True, linestyle=":", alpha=0.6)
ax.legend(loc="lower left")
ax.xaxis.set_major_locator(mdates.YearLocator(2))
ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
fig.autofmt_xdate()
fig.tight_layout()

OUT_DIR = Path(__file__).resolve().parent / "output" / "yield_curve"
OUT_DIR.mkdir(parents=True, exist_ok=True)
chart_path = OUT_DIR / "spread_chart.png"
fig.savefig(chart_path, dpi=300)
print(f"Saved {chart_path}")

chart_df = df.reset_index()
xlsx_path = OUT_DIR / "yields.xlsx"
if xlsx_path.exists():
    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        chart_df.to_excel(writer, sheet_name="Chart Data", index=False)
else:
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        chart_df.to_excel(writer, sheet_name="Chart Data", index=False)

print(f"Appended 'Chart Data' sheet to {xlsx_path} ({len(chart_df)} rows)")
print(f"Recession periods detected: {len(bands)}")
