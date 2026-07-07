# NIPA / Z.1 macro-accounts P/E for the US corporate sector, quarterly 1947Q1-present.
#
# INTENTIONAL EXCEPTION to the repo's 2006-01-01 start convention: this valuation
# series is only meaningful across full valuation cycles (post-WWII, Nifty Fifty,
# 1970s-80s single-digit trough, dot-com, GFC, today), so the primary charts run
# 1947-present. A 2006+ companion chart (nipa_pe_2006.png) is emitted so the
# headline ratio also slots into the comparable-axis macro chart set.
#
# Construction: numerator = Financial Accounts (Z.1) market value of corporate
# equities outstanding (liability side); denominator = NIPA corporate profits
# (SAAR flows, already annualized — no further annualization). Equities come in
# MILLIONS of $ and are divided by 1,000 to billions to match profits.
#
# Scope-matched ratio pairs:
#   1. HEADLINE  BOGZ1LM883164105Q / CPATAX  — all-corporate (incl. financials)
#      equities over after-tax profits WITH IVA+CCAdj ("economic" after-tax P/E).
#      Also computed with a 4-quarter trailing average of profits — smoother and
#      closest analogue to a trailing P/E.
#   2. BOGZ1LM883164105Q / CP        — after-tax profits WITHOUT IVA+CCAdj
#      ("book-style" profits variant).
#   3. NCBEILQ027S / NFCPATAX        — nonfinancial corporate pair (the series to
#      line up with the Z.1 nonfinancial EV/EBITDA measure later).
#   4. BOGZ1LM883164105Q / CPROFIT   — PRETAX profits WITH IVA+CCAdj; isolates
#      the 2017 TCJA tax step-change from valuation moves.
#
# CAVEATS: this is an economy-wide multiple. Z.1 equities include closely-held /
# unlisted equity, and NIPA profits include private firms, S-corps, and foreign
# earnings of US corporates — so its LEVEL is not comparable to the S&P 500 P/E;
# only its own history is the reference. IVA/CCAdj variants (CP vs CPATAX,
# CPROFIT) diverge most in high-inflation and high-capex eras (1970s-80s), when
# inventory and depreciation distortions in book profits are largest.
# MVEONWMVBSNNCB is discontinued (2017) and deliberately not used.
#
# Data starts: equities 1945Q4, profits 1947Q1. NOTE: the Z.1 equity series are
# ANNUAL frequency before 1952 (observations only in Q4, dated Oct 1), so the
# ratios are actually defined from 1947Q4 and are Q4-only through 1951; true
# quarterly coverage begins 1952Q1. Charts connect the five annual points with
# straight segments. Pull uses a 1945-01-01 buffer.

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
    series_stats,
    style_macro_chart,
)

fred = get_fred_client()

end = datetime.today()
chart_start = datetime(1947, 1, 1)   # profits begin 1947Q1
chart_start_2006 = datetime(2006, 1, 1)
data_start = datetime(1945, 1, 1)    # buffer; equities exist from 1945Q4

series_map = {
    # Z.1 market value of corporate equities outstanding, MILLIONS of $, NSA,
    # end of period.
    "EQ_ALLCORP_MM": "BOGZ1LM883164105Q",  # all domestic sectors (incl. financials)
    "EQ_NCB_MM": "NCBEILQ027S",            # nonfinancial corporate business
    # NIPA corporate profits, BILLIONS of $, SAAR.
    "CP": "CP",                # after tax, w/o IVA+CCAdj
    "CPATAX": "CPATAX",        # after tax, WITH IVA+CCAdj
    "NFCPATAX": "NFCPATAX",    # nonfinancial, after tax, w/o adjustments
    "CPROFIT": "CPROFIT",      # total PRETAX, WITH IVA+CCAdj
}

df = pull_series(fred, series_map, data_start, end)

# Millions -> billions so equities match the profit series' units.
df["EQ_ALLCORP"] = df["EQ_ALLCORP_MM"] / 1000.0
df["EQ_NCB"] = df["EQ_NCB_MM"] / 1000.0
df = df.drop(columns=["EQ_ALLCORP_MM", "EQ_NCB_MM"])

# Profits are SAAR (annualized) flows, so the plain quarterly ratio is already a
# P/E. The 4q trailing average of profits is a smoother variant for the headline.
df["CPATAX_4Q"] = df["CPATAX"].rolling(4).mean()
df["CP_4Q"] = df["CP"].rolling(4).mean()

df["PE_ALLCORP_CPATAX"] = df["EQ_ALLCORP"] / df["CPATAX"]
df["PE_ALLCORP_CPATAX_4Q"] = df["EQ_ALLCORP"] / df["CPATAX_4Q"]
df["PE_ALLCORP_CP"] = df["EQ_ALLCORP"] / df["CP"]
df["PE_ALLCORP_CP_4Q"] = df["EQ_ALLCORP"] / df["CP_4Q"]
df["PE_NCB_NFCPATAX"] = df["EQ_NCB"] / df["NFCPATAX"]
df["PE_ALLCORP_PRETAX"] = df["EQ_ALLCORP"] / df["CPROFIT"]

# Drop the 1945Q4-1946Q4 buffer rows (equities only, no profits -> no ratios).
df = df[df["Date"] >= pd.Timestamp(chart_start)].reset_index(drop=True)

RATIO_COLS = [
    "PE_ALLCORP_CPATAX",
    "PE_ALLCORP_CPATAX_4Q",
    "PE_ALLCORP_CP",
    "PE_ALLCORP_CP_4Q",
    "PE_NCB_NFCPATAX",
    "PE_ALLCORP_PRETAX",
]

# Summary: series_stats over the full 1947+ window (the charted window for the
# primary charts) plus the current value's percentile rank within that window.
summary_rows = {c: series_stats(df[c]) for c in RATIO_COLS}
summary = pd.DataFrame(summary_rows).T[["min", "max", "mean", "median", "current"]]
summary["current_pctile"] = [
    (df[c].dropna() <= df[c].dropna().iloc[-1]).mean() * 100 for c in RATIO_COLS
]

OUT_DIR = resolve_output_dir(__file__, "nipa_pe")
(OUT_DIR / ".gitkeep").touch()
xlsx_path = OUT_DIR / "nipa_pe.xlsx"
with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Data", index=False)
    summary.to_excel(writer, sheet_name="Summary")

recessions = get_recession_periods(fred, chart_start, end)
XLIM = (pd.Timestamp(chart_start), pd.Timestamp(end))

headline_median = df["PE_ALLCORP_CPATAX_4Q"].median()

# Per-ratio NaN-dropped frames for plotting: pre-1952 Z.1 equity data is annual
# (Q4-only), so plotting the raw columns would leave those five early
# observations as invisible isolated points between NaNs. Dropping NaNs lets
# matplotlib connect them with straight segments.
plot_df = {c: df[["Date", c]].dropna() for c in RATIO_COLS}

# Chart 1 — headline: all-corporate equities / after-tax economic profits (4q-avg
# profits), with the book-style CP variant (also 4q-avg, so both lines carry the
# same smoothing) lighter behind it.
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(plot_df["PE_ALLCORP_CP_4Q"]["Date"], plot_df["PE_ALLCORP_CP_4Q"]["PE_ALLCORP_CP_4Q"],
        color="#9ec5e8", linewidth=1.2, label="vs after-tax book profits (CP, 4q avg)")
ax.plot(plot_df["PE_ALLCORP_CPATAX_4Q"]["Date"], plot_df["PE_ALLCORP_CPATAX_4Q"]["PE_ALLCORP_CPATAX_4Q"],
        color="#1f3b73", linewidth=1.8, label="vs after-tax economic profits (CPATAX, 4q avg)")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="US corporate sector P/E — Z.1 equities / NIPA after-tax profits, 1947–present",
    ylabel="Ratio (x)",
    recessions=recessions,
    hlines=[{"y": headline_median,
             "label": f"Median {headline_median:.1f}x (1947–present)"}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "nipa_pe_full.png", dpi=150)
plt.close(fig)

# Chart 2 — scope: nonfinancial pair vs all-corporate pair (plain quarterly
# ratios; the gap shows the financials-sector effect on the aggregate multiple).
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(plot_df["PE_NCB_NFCPATAX"]["Date"], plot_df["PE_NCB_NFCPATAX"]["PE_NCB_NFCPATAX"],
        color="#9ec5e8", linewidth=1.2, label="Nonfinancial (NCB equities / NFCPATAX)")
ax.plot(plot_df["PE_ALLCORP_CPATAX"]["Date"], plot_df["PE_ALLCORP_CPATAX"]["PE_ALLCORP_CPATAX"],
        color="#1f3b73", linewidth=1.8, label="All corporate (incl. financials / CPATAX)")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="Macro P/E scope — nonfinancial vs all-corporate pairs, 1947–present",
    ylabel="Ratio (x)",
    # Clipped at 60x so the 2001Q4 nonfinancial spike (92.6x — dot-com profits
    # collapse) doesn't compress the rest of the history. Spike runs off-axis
    # by design; the ~54x dot-com peak remains visible.
    ylim=(0, 60),
    recessions=recessions,
)
fig.tight_layout()
fig.savefig(OUT_DIR / "nipa_pe_scope.png", dpi=150)
plt.close(fig)

# Chart 3 — after-tax vs pretax overlay, plotted as raw ratios on ONE shared
# axis (not indexed, no second axis): both are P/E-style ratios in identical
# units and the vertical gap between them IS the effective-tax wedge, so the
# TCJA divergence reads directly — after 2018Q1 the after-tax line pulls away
# from the pretax line as the statutory rate dropped 35% -> 21%.
tcja = pd.Timestamp("2018-01-01")
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(plot_df["PE_ALLCORP_PRETAX"]["Date"], plot_df["PE_ALLCORP_PRETAX"]["PE_ALLCORP_PRETAX"],
        color="#9ec5e8", linewidth=1.2, label="vs PRETAX economic profits (CPROFIT)")
ax.plot(plot_df["PE_ALLCORP_CPATAX"]["Date"], plot_df["PE_ALLCORP_CPATAX"]["PE_ALLCORP_CPATAX"],
        color="#1f3b73", linewidth=1.8, label="vs after-tax economic profits (CPATAX)")
ax.axvline(tcja, color="#c0392b", linestyle=":", linewidth=1.2,
           label="TCJA effective (2018Q1)")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="Macro P/E — after-tax vs pretax profits (TCJA tax step-change), 1947–present",
    ylabel="Ratio (x)",
    recessions=recessions,
)
fig.tight_layout()
fig.savefig(OUT_DIR / "nipa_pe_pretax.png", dpi=150)
plt.close(fig)

# Chart 4 — 2006+ companion of the headline ratio so it slots into the
# comparable-axis macro chart set. Median recomputed over the 2006+ window
# (what's on this chart), per repo convention for reference stats.
df_2006 = df[df["Date"] >= pd.Timestamp(chart_start_2006)].reset_index(drop=True)
median_2006 = df_2006["PE_ALLCORP_CPATAX_4Q"].median()
recessions_2006 = [r for r in recessions if r[1] >= pd.Timestamp(chart_start_2006)]
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df_2006["Date"], df_2006["PE_ALLCORP_CPATAX_4Q"], color="#1f3b73",
        linewidth=1.8, label="vs after-tax economic profits (CPATAX, 4q avg)")
ax.set_xlim(pd.Timestamp(chart_start_2006), pd.Timestamp(end))
style_macro_chart(
    ax,
    title="US corporate sector P/E — Z.1 equities / NIPA profits, 2006–present",
    ylabel="Ratio (x)",
    recessions=recessions_2006,
    hlines=[{"y": median_2006,
             "label": f"Median {median_2006:.1f}x (2006–present)"}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "nipa_pe_2006.png", dpi=150)
plt.close(fig)

latest = df.dropna(subset=["PE_ALLCORP_CPATAX"]).iloc[-1]
print(f"Start date:                {chart_start.date()}")
print(f"End date:                  {end.date()}")
print(f"Rows (1947Q1+):            {len(df)}")
print(f"Latest ratio quarter:      {latest['Date'].date()}")
print(f"Latest AllCorp equities:   ${latest['EQ_ALLCORP']:,.0f}B")
print(f"Latest CPATAX (SAAR):      ${latest['CPATAX']:,.0f}B")
first_ratio = df.dropna(subset=["PE_ALLCORP_CPATAX"]).iloc[0]
print(f"First PE_ALLCORP_CPATAX:   {first_ratio['PE_ALLCORP_CPATAX']:.1f}x "
      f"({first_ratio['Date'].date()} -- Z.1 equities are annual pre-1952)")
print(f"Latest PE_ALLCORP_CPATAX:  {latest['PE_ALLCORP_CPATAX']:.1f}x")
print(f"Latest headline (4q avg):  {df['PE_ALLCORP_CPATAX_4Q'].dropna().iloc[-1]:.1f}x "
      f"(median {headline_median:.1f}x, "
      f"{summary.loc['PE_ALLCORP_CPATAX_4Q', 'current_pctile']:.0f}th pctile since 1947)")
print(f"Latest PE_ALLCORP_CP:      {df['PE_ALLCORP_CP'].dropna().iloc[-1]:.1f}x")
print(f"Latest PE_NCB_NFCPATAX:    {df['PE_NCB_NFCPATAX'].dropna().iloc[-1]:.1f}x")
print(f"Latest PE_ALLCORP_PRETAX:  {df['PE_ALLCORP_PRETAX'].dropna().iloc[-1]:.1f}x")
print(f"Wrote {xlsx_path.name}, nipa_pe_full.png, nipa_pe_scope.png, "
      f"nipa_pe_pretax.png, nipa_pe_2006.png to {OUT_DIR}")
