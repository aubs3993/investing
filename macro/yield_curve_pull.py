from datetime import datetime, timedelta
from dotenv import load_dotenv
from fredapi import Fred
from pathlib import Path
import os
import pandas as pd

load_dotenv()

api_key = os.getenv("FRED_API_KEY")
if not api_key:
    raise RuntimeError("FRED_API_KEY not found in environment or .env file")

fred = Fred(api_key=api_key.strip())

end = datetime.today()
start = end - timedelta(days=365 * 20)

series = {
    "10Y": "DGS10",
    "2Y": "DGS2",
    "FedFunds": "DFF",
}

frames = {
    name: fred.get_series(sid, observation_start=start, observation_end=end)
    for name, sid in series.items()
}

df = pd.DataFrame(frames)
df.index.name = "Date"
df = df.reset_index()
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

OUT_DIR = Path(__file__).resolve().parent / "output" / "yield_curve"
OUT_DIR.mkdir(parents=True, exist_ok=True)
out_path = OUT_DIR / "yields.xlsx"
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Data", index=False)
    summary.to_excel(writer, sheet_name="Summary")

print(f"Start date:  {df['Date'].iloc[0].date()}")
print(f"End date:    {df['Date'].iloc[-1].date()}")
print(f"Rows:        {len(df)}")
print(f"Latest 10Y-2Y spread: {df['10Y_minus_2Y'].iloc[-1]:.2f}")
print(
    f"Pulled {len(df)} rows. Note: rows are dropped if any of DGS10/DGS2/DFF "
    "is missing — DFF typically lags by 1 trading day, so the pull may end "
    "1 day before the chart's T10Y2Y series."
)
