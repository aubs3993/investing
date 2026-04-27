# NOTE on indicator choice:
# (a) ISM Manufacturing PMI was removed from FRED in June 2016 due to ISM's
#     licensing change. To stay on a free, long-history data source we use
#     the regional Federal Reserve manufacturing surveys instead.
# (b) Of the five regional Fed manufacturing surveys, only three are
#     distributed on FRED: New York Empire State, Philadelphia Manufacturing
#     Business Outlook, and Dallas Texas Manufacturing Outlook. The Richmond
#     and Kansas City surveys are published only on those Feds' own sites.
#     The composite below is therefore an equal-weighted average of NY,
#     Philly, and Dallas — the three most-watched, and the three with the
#     longest FRED history. Composite is computed only on dates where all
#     three are non-null. Effective composite start is 2004-06-01 (Dallas's
#     FRED start), comfortably before the 2006-01-01 chart window.
# (c) CFNAI (Chicago Fed National Activity Index) is added as a broader-
#     economy cross-check — it weights 85 monthly indicators across
#     production, employment, consumption, and sales/orders. CFNAIMA3 is
#     the 3-month MA, which the Chicago Fed recommends for interpretation
#     because the monthly series is noisy.

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
chart_start = datetime(2006, 1, 1)
# Modest buffer so any month-boundary edge cases land before the chart window.
# No moving averages are computed here (CFNAIMA3 is fetched pre-computed), so
# the buffer is small.
data_start = datetime(2005, 1, 1)

regional_series = {
    "NY": "GACDISA066MSFRBNY",
    "Philly": "GACDFSA066MSFRBPHI",
    "Dallas": "BACTSAMFRBDAL",
}
cfnai_series = {
    "CFNAI": "CFNAI",
    "CFNAIMA3": "CFNAIMA3",
}

df_regional = pull_series(fred, regional_series, data_start, end)
# Composite is the row-wise mean across the three surveys, defined only on
# dates where all three are non-null (skipna=False). Anything before Dallas's
# 2004-06 start gets NaN here and is dropped at the chart-start filter below.
df_regional["composite"] = df_regional[["NY", "Philly", "Dallas"]].mean(
    axis=1, skipna=False
)
df_regional = df_regional[df_regional["Date"] >= pd.Timestamp(chart_start)].reset_index(drop=True)

df_cfnai = pull_series(fred, cfnai_series, data_start, end)
df_cfnai = df_cfnai[df_cfnai["Date"] >= pd.Timestamp(chart_start)].reset_index(drop=True)


def _stats(series: pd.Series) -> dict:
    s = series.dropna()
    return {
        "min": s.min(),
        "max": s.max(),
        "mean": s.mean(),
        "median": s.median(),
        "current": s.iloc[-1] if len(s) else None,
    }


summary_rows = {
    "composite": _stats(df_regional["composite"]),
    "CFNAIMA3": _stats(df_cfnai["CFNAIMA3"]),
}
summary = pd.DataFrame(summary_rows).T[["min", "max", "mean", "median", "current"]]

current_composite = float(df_regional["composite"].dropna().iloc[-1])
current_cfnaima3 = float(df_cfnai["CFNAIMA3"].dropna().iloc[-1])
composite_below_zero = current_composite < 0
cfnai_below_neg07 = current_cfnaima3 < -0.7

summary.loc["composite_below_zero"] = [pd.NA, pd.NA, pd.NA, pd.NA, composite_below_zero]
summary.loc["cfnaima3_below_neg07"] = [pd.NA, pd.NA, pd.NA, pd.NA, cfnai_below_neg07]

OUT_DIR = resolve_output_dir(__file__, "leading_indicators")
xlsx_path = OUT_DIR / "leading_indicators.xlsx"
with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    df_regional.to_excel(writer, sheet_name="Data_Regional", index=False)
    df_cfnai.to_excel(writer, sheet_name="Data_CFNAI", index=False)
    summary.to_excel(writer, sheet_name="Summary")

recessions = get_recession_periods(fred, chart_start, end)
XLIM = (pd.Timestamp(chart_start), pd.Timestamp(end))

fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df_regional["Date"], df_regional["composite"], color="#1f3b73",
        linewidth=2.0, label="Composite (avg of NY, Philly, Dallas)")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="Regional Fed Manufacturing Composite (avg of NY, Philly, Dallas), 2006–present",
    ylabel="Diffusion index",
    # Clipped at +/-40 so COVID and GFC extremes run off-axis but normal
    # cycle variation is readable.
    ylim=(-40, 40),
    recessions=recessions,
    hlines=[{"y": 0.0, "label": "Expansion / contraction"}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "regional_fed_composite.png", dpi=150)
plt.close(fig)

fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df_cfnai["Date"], df_cfnai["CFNAI"], color="#9ec5e8", linewidth=1.0,
        label="CFNAI (monthly)")
ax.plot(df_cfnai["Date"], df_cfnai["CFNAIMA3"], color="#1f3b73", linewidth=2.0,
        label="CFNAIMA3 (3-month MA)")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="Chicago Fed National Activity Index, 2006–present",
    ylabel="Index value",
    # Tightened to (-3, 2) so the GFC trough still sits cleanly inside the
    # frame while normal-cycle variation around zero is more legible. COVID
    # extremes run off-axis as designed.
    ylim=(-3, 2),
    recessions=recessions,
    hlines=[
        {"y": 0.0, "label": "Zero growth"},
        {"y": -0.7, "label": "Recession threshold", "color": "#c0392b"},
    ],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "cfnai.png", dpi=150)
plt.close(fig)

print(f"Start date:                {chart_start.date()}")
print(f"End date:                  {end.date()}")
print(f"Regional rows:             {len(df_regional)}")
print(f"CFNAI rows:                {len(df_cfnai)}")
print(f"Latest NY:                 {df_regional['NY'].dropna().iloc[-1]:+.1f}")
print(f"Latest Philly:             {df_regional['Philly'].dropna().iloc[-1]:+.1f}")
print(f"Latest Dallas:             {df_regional['Dallas'].dropna().iloc[-1]:+.1f}")
print(f"Latest composite:          {current_composite:+.1f}")
print(f"Latest CFNAI (monthly):    {df_cfnai['CFNAI'].dropna().iloc[-1]:+.2f}")
print(f"Latest CFNAIMA3:           {current_cfnaima3:+.2f}")
print(f"Composite below zero?      {composite_below_zero}")
print(f"CFNAIMA3 below -0.7?       {cfnai_below_neg07}")
print(f"Wrote {xlsx_path.name}, regional_fed_composite.png, cfnai.png to {OUT_DIR}")
