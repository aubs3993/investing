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
# 12-month buffer so derived calcs (monthly diff + 3-month MA, 4-week MA on
# weekly claims) are defined at chart_start rather than starting blank.
data_start = datetime(2005, 1, 1)

monthly_series = {
    "PAYEMS": "PAYEMS",
    "UNRATE": "UNRATE",
    "SAHMREALTIME": "SAHMREALTIME",
}
weekly_series = {
    "ICSA": "ICSA",
}

df_m = pull_series(fred, monthly_series, data_start, end)
df_m["PAYEMS_change"] = df_m["PAYEMS"].diff()
df_m["PAYEMS_change_ma3"] = df_m["PAYEMS_change"].rolling(3).mean()
df_m = df_m[df_m["Date"] >= pd.Timestamp(chart_start)].reset_index(drop=True)

df_w = pull_series(fred, weekly_series, data_start, end)
df_w["ICSA_ma4"] = df_w["ICSA"].rolling(4).mean()
df_w = df_w[df_w["Date"] >= pd.Timestamp(chart_start)].reset_index(drop=True)

# Build Summary stats per series. Each row is one series; columns are stats.
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
    "PAYEMS": _stats(df_m["PAYEMS"]),
    "PAYEMS_change": _stats(df_m["PAYEMS_change"]),
    "PAYEMS_change_ma3": _stats(df_m["PAYEMS_change_ma3"]),
    "UNRATE": _stats(df_m["UNRATE"]),
    "SAHMREALTIME": _stats(df_m["SAHMREALTIME"]),
    "ICSA": _stats(df_w["ICSA"]),
    "ICSA_ma4": _stats(df_w["ICSA_ma4"]),
}
summary = pd.DataFrame(summary_rows).T[["min", "max", "mean", "median", "current"]]

current_sahm = float(df_m["SAHMREALTIME"].dropna().iloc[-1])
sahm_triggered = current_sahm >= 0.5
# Append the flag as a row, with True/False in the "current" column. Other
# stat columns stay blank — the SAHMREALTIME row above already carries them.
summary.loc["sahm_triggered"] = [pd.NA, pd.NA, pd.NA, pd.NA, sahm_triggered]

OUT_DIR = resolve_output_dir(__file__, "labor")
xlsx_path = OUT_DIR / "labor.xlsx"
with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    df_m.to_excel(writer, sheet_name="Data_Monthly", index=False)
    df_w.to_excel(writer, sheet_name="Data_Weekly", index=False)
    summary.to_excel(writer, sheet_name="Summary")

recessions = get_recession_periods(fred, chart_start, end)
XLIM = (pd.Timestamp(chart_start), pd.Timestamp(end))

# Payrolls: bars for monthly change, green positive / red negative, with a
# black 3-month MA line overlaid for the underlying-trend signal.
POS_COLOR = "#2ca02c"
NEG_COLOR = "#c0392b"
fig, ax = plt.subplots(figsize=(11, 5))
bar_colors = [POS_COLOR if v >= 0 else NEG_COLOR for v in df_m["PAYEMS_change"]]
ax.bar(df_m["Date"], df_m["PAYEMS_change"], color=bar_colors, width=25,
       label="Monthly change")
ax.plot(df_m["Date"], df_m["PAYEMS_change_ma3"], color="black", linewidth=1.5,
        label="3-month MA")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="Nonfarm payrolls — monthly change, 2006–present",
    ylabel="Jobs added (thousands)",
    # Clipped at +/-5000 so the COVID swing (-20M / +5M) doesn't compress the
    # rest of the cycle. COVID bars/MA visibly run off-axis by design.
    ylim=(-5000, 5000),
    recessions=recessions,
    hlines=[{"y": 0.0}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "payrolls.png", dpi=150)
plt.close(fig)

# Jobless claims: weekly ICSA (lighter) with 4-week MA (heavier — cleaner signal).
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df_w["Date"], df_w["ICSA"], color="#9ec5e8", linewidth=1.0,
        label="Weekly initial claims")
ax.plot(df_w["Date"], df_w["ICSA_ma4"], color="#1f3b73", linewidth=2.0,
        label="4-week MA")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="Initial jobless claims, 2006–present",
    ylabel="Claims",
    # Clipped at 1M so the COVID 6M+ peak doesn't compress GFC and current.
    ylim=(0, 1_000_000),
    recessions=recessions,
)
fig.tight_layout()
fig.savefig(OUT_DIR / "jobless_claims.png", dpi=150)
plt.close(fig)

# Unemployment rate: single line.
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df_m["Date"], df_m["UNRATE"], color="#1f3b73", linewidth=1.5,
        label="Unemployment rate")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="Unemployment rate, 2006–present",
    ylabel="Percent",
    # Clipped at 10% so the COVID ~14.8% spike doesn't compress GFC and trend.
    ylim=(0, 10),
    recessions=recessions,
)
fig.tight_layout()
fig.savefig(OUT_DIR / "unemployment.png", dpi=150)
plt.close(fig)

# Sahm rule (real-time): line with 0.5pp recession-trigger threshold.
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df_m["Date"], df_m["SAHMREALTIME"], color="#1f3b73", linewidth=1.5,
        label="Sahm Rule (real-time)")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="Sahm Rule recession indicator (real-time), 2006–present",
    ylabel="Percentage points",
    # Clipped at 4pp so the COVID 9.5pp peak doesn't compress GFC and current.
    # Lower bound preserves the small post-GFC negative readings.
    ylim=(-0.5, 4),
    recessions=recessions,
    hlines=[{"y": 0.5, "label": "Recession trigger", "color": "#c0392b"}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "sahm_rule.png", dpi=150)
plt.close(fig)

print(f"Start date:           {chart_start.date()}")
print(f"End date:             {end.date()}")
print(f"Monthly rows:         {len(df_m)}")
print(f"Weekly rows:          {len(df_w)}")
print(f"Latest PAYEMS level:  {df_m['PAYEMS'].dropna().iloc[-1]:,.0f}k")
print(f"Latest PAYEMS change: {df_m['PAYEMS_change'].dropna().iloc[-1]:+,.0f}k")
print(f"Latest UNRATE:        {df_m['UNRATE'].dropna().iloc[-1]:.1f}%")
print(f"Latest SAHMREALTIME:  {current_sahm:.2f}pp")
print(f"Latest ICSA:          {df_w['ICSA'].dropna().iloc[-1]:,.0f}")
print(f"Latest ICSA 4w MA:    {df_w['ICSA_ma4'].dropna().iloc[-1]:,.0f}")
print(f"Sahm Rule triggered:  {sahm_triggered}")
print(f"Wrote {xlsx_path.name}, payrolls.png, jobless_claims.png, "
      f"unemployment.png, sahm_rule.png to {OUT_DIR}")
