# US credit-cycle module: price / standards / performance / quantity panels plus
# a two-block composite gauge (STRESS and FROTH), paired with equities and rates.
#
# INTENTIONAL EXCEPTION to the repo's 2006-01-01 start convention: the credit
# cycle only reads across full cycles (1930s depression, 1974-75, 1980-82, 1990,
# 2001-02, 2008-09, 2020), so the primary charts run from 1919 (Baa/Aaa spread
# start). A 2006+ companion chart (credit_composite_2006.png) is emitted so the
# composites also slot into the comparable-axis macro chart set.
#
# METHOD NOTES:
# - Z-scores are computed over each component's FULL monthly sample. That uses
#   future information (lookahead) — fine for a descriptive "where are we vs.
#   history" chart, but this is NOT a backtest and must not be treated as one.
# - SLOOS series (DRTSCILM etc.) are net-percent-tightening DIFFUSION indexes:
#   they measure the CHANGE in standards each quarter, not the level of
#   standards. A long run of small positives is cumulative tightening.
# - Loan performance (delinquencies, charge-offs) LAGS the cycle by roughly 2-4
#   quarters. It confirms a downturn; it does not time one. Deliberately
#   excluded from the composites.
# - EBP (excess bond premium, Favara-Gilchrist-Lewis-Zakrajsek FEDS Notes) is
#   RE-ESTIMATED over its full history every monthly update, so we re-download
#   each run and archive a copy in the output folder (gitignored, overwritten).
# - FROTH omission: the Greenwood-Hanson high-yield-share-of-gross-issuance leg
#   is deliberately OMITTED — the SIFMA issuance source requires registration
#   and carries redistribution-license restrictions. Flagged as a future manual
#   addition; until then FROTH = spread ease + credit impulses + SLOOS easing.
# - DO NOT read or touch macro/output/credit_spreads/hy_oas_archive.csv — that
#   belongs to the separate credit_spreads module.
#
# COMPONENT START DATES (composites gain legs as data begins):
#   Baa-Aaa spread 1919-01; spread 12m change 1920-01; EBP 1973-01;
#   SLOOS C&I standards 1990Q2; aggregate & household credit impulses 1953Q3 —
#   Z.1 flow series are ANNUAL frequency (Q4-only observations) before 1952, so
#   the 4q rolling mean is first defined 1952Q3 and the 4q change 1953Q3.
#
# Flow units: Z.1 FA/HNO transactions series are Mil$ SAAR -> /1000 to Bil$ to
# match GDP and DPI (Bil$ SAAR).

from datetime import datetime
from pathlib import Path
import sys

import matplotlib.pyplot as plt
import pandas as pd
import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.fred_helpers import (
    get_fred_client,
    get_recession_periods,
    pull_series,
    resolve_output_dir,
    series_stats,
    style_macro_chart,
)

MAIN = "#1f3b73"
LIGHT = "#9ec5e8"
POS = "#2ca02c"
NEG = "#c0392b"

fred = get_fred_client()
end = datetime.today()
data_start = datetime(1900, 1, 1)     # buffer; earliest series (BAA/AAA) is 1919-01
data_start_q = datetime(1946, 1, 1)   # quarterly buffer; Z.1 flows begin 1946Q4
chart_start_2006 = datetime(2006, 1, 1)

OUT_DIR = resolve_output_dir(__file__, "credit_cycle")
(OUT_DIR / ".gitkeep").touch()

# ---------------------------------------------------------------------------
# Section 1 — FRED pulls (grouped by native frequency)
# ---------------------------------------------------------------------------
monthly_map = {
    "BAA": "BAA",              # Moody's Baa yield, 1919+
    "AAA": "AAA",              # Moody's Aaa yield, 1919+
    "BAA10YM": "BAA10YM",      # Baa minus 10y Treasury, 1953+
    "BUSLOANS": "BUSLOANS",    # C&I loans, all commercial banks, 1947+
    "REVOLSL": "REVOLSL",      # revolving consumer credit, 1968+
    "GS10": "GS10",            # 10y Treasury yield, 1953+
    "FEDFUNDS": "FEDFUNDS",    # effective fed funds, 1954+
}
weekly_map = {
    "TOTBKCR": "TOTBKCR",          # total bank credit, weekly 1973+
    "NFCI": "NFCI",                # Chicago Fed NFCI, weekly 1971+
    "NFCICREDIT": "NFCICREDIT",    # NFCI credit subindex, weekly 1971+
}
quarterly_map = {
    # SLOOS standards (diffusion: net % tightening / increased willingness)
    "DRTSCILM": "DRTSCILM",    # C&I standards, large/medium firms, 1990Q2+
    "DRIWCIL": "DRIWCIL",      # willingness to make consumer loans, 1982Q2+
    "DRTSCLCC": "DRTSCLCC",    # credit card standards, 1996Q1+
    # Performance (lagging confirmation, NOT in composites)
    "DRBLACBS": "DRBLACBS",    # business loan delinquency rate, 1987Q1+
    "DRCCLACBS": "DRCCLACBS",  # credit card delinquency rate, 1991Q1+
    "CORBLACBS": "CORBLACBS",  # business loan charge-off rate, 1985Q1+
    # Quantity / flows
    "QUSPAM770A": "QUSPAM770A",              # private nonfin credit % GDP, 1947Q4+
    "FLOW_AGG_MM": "BOGZ1FA384104005Q",      # all dom. nonfin net borrowing, Mil$ SAAR
    "FLOW_HH_MM": "BOGZ1FA154104005Q",       # household net borrowing, Mil$ SAAR
    "FLOW_CC_MM": "HNOCCLQ027S",             # household consumer-credit flow, Mil$ SAAR
    # Denominators
    "GDP": "GDP",              # Bil$ SAAR, 1947Q1+
    "DPI": "DPI",              # disposable personal income, Bil$ SAAR, 1947Q1+
}

df_mraw = pull_series(fred, monthly_map, data_start, end)
df_wraw = pull_series(fred, weekly_map, data_start, end)
df_q = pull_series(fred, quarterly_map, data_start_q, end)

# ---------------------------------------------------------------------------
# Section 2 — EBP download (Fed revises FULL history monthly: re-download every
# run and archive a copy in the output folder, overwriting; gitignored)
# ---------------------------------------------------------------------------
EBP_URL = "https://www.federalreserve.gov/econres/notes/feds-notes/ebp_csv.csv"
resp = requests.get(EBP_URL, timeout=120, headers={"User-Agent": "Mozilla/5.0"})
resp.raise_for_status()
ebp_path = OUT_DIR / "ebp_csv.csv"
ebp_path.write_bytes(resp.content)
ebp = pd.read_csv(ebp_path)
needed = ["date", "gz_spread", "ebp", "est_prob"]
missing = [c for c in needed if c not in ebp.columns]
if missing:
    raise RuntimeError(f"EBP csv layout changed; missing {missing}. Got: {list(ebp.columns)}")
ebp["Date"] = pd.to_datetime(ebp["date"])
ebp = ebp.set_index("Date")[["gz_spread", "ebp", "est_prob"]]

# ---------------------------------------------------------------------------
# Section 3 — quarterly constructions: credit impulses and borrowing rate
# ---------------------------------------------------------------------------
df_q = df_q.set_index("Date")
for mm, bil in [("FLOW_AGG_MM", "FLOW_AGG"), ("FLOW_HH_MM", "FLOW_HH"),
                ("FLOW_CC_MM", "FLOW_CC")]:
    df_q[bil] = df_q[mm] / 1000.0   # Mil$ SAAR -> Bil$ SAAR
df_q = df_q.drop(columns=["FLOW_AGG_MM", "FLOW_HH_MM", "FLOW_CC_MM"])

# 4Q rolling MEAN of the SAAR flow smooths quarterly Z.1 noise; the impulse is
# the 4-quarter change in that smoothed flow, scaled by nominal income.
df_q["FLOW4Q_AGG"] = df_q["FLOW_AGG"].rolling(4).mean()
df_q["FLOW4Q_HH"] = df_q["FLOW_HH"].rolling(4).mean()
df_q["FLOW4Q_CC"] = df_q["FLOW_CC"].rolling(4).mean()
df_q["DPI4Q"] = df_q["DPI"].rolling(4).mean()

# Aggregate impulse over GDP; household legs over 4q-mean DPI (per spec).
df_q["IMPULSE_AGG"] = 100 * (df_q["FLOW4Q_AGG"] - df_q["FLOW4Q_AGG"].shift(4)) / df_q["GDP"]
df_q["CI_HH"] = 100 * (df_q["FLOW4Q_HH"] - df_q["FLOW4Q_HH"].shift(4)) / df_q["DPI4Q"]
df_q["CI_CC"] = 100 * (df_q["FLOW4Q_CC"] - df_q["FLOW4Q_CC"].shift(4)) / df_q["DPI4Q"]
# Household borrowing rate: smoothed net borrowing as % of disposable income.
# Anchors: peak ~13.9% 2006Q2, trough ~-2.1% 2009Q3, recent ~3%.
df_q["B_HH"] = 100 * df_q["FLOW4Q_HH"] / df_q["DPI4Q"]

# ---------------------------------------------------------------------------
# Section 4 — monthly master frame: spreads, EBP, resampled weeklies, and
# quarterly series forward-filled to monthly
# ---------------------------------------------------------------------------
df_m = df_mraw.set_index("Date")
midx = pd.date_range(df_m["BAA"].dropna().index.min(), df_m.index.max(), freq="MS")
df_m = df_m.reindex(midx)
df_m.index.name = "Date"

df_m["SPREAD"] = df_m["BAA"] - df_m["AAA"]
df_m["SPREAD_CHG"] = df_m["SPREAD"].diff(12)   # 12m change

# Weekly -> monthly means (TOTBKCR, NFCI, NFCICREDIT)
df_w = df_wraw.set_index("Date")
for col in ["TOTBKCR", "NFCI", "NFCICREDIT"]:
    df_m[f"{col}_M"] = df_w[col].resample("MS").mean().reindex(midx)

df_m = df_m.join(ebp)   # gz_spread, ebp, est_prob (monthly, 1973-01+)

# Quarterly -> monthly forward-fill. Quarter-start observations carry forward
# up to 8 months (one quarter of coverage plus reporting lag) so a stale or
# discontinued series can't ffill for years.
for qcol, mcol in [("DRTSCILM", "DRTSCILM_M"), ("IMPULSE_AGG", "IMPULSE_AGG_M"),
                   ("CI_HH", "CI_HH_M")]:
    df_m[mcol] = df_q[qcol].reindex(midx).ffill(limit=8)

# YoY growth of quantity series (context columns for the xlsx, not charted)
for col in ["BUSLOANS", "TOTBKCR_M", "REVOLSL"]:
    df_m[f"{col}_YOY"] = 100 * (df_m[col] / df_m[col].shift(12) - 1)

# ---------------------------------------------------------------------------
# Section 5 — composites. Full-sample z-scores winsorized at +/-3; each
# composite is the equal-weight mean of whichever components exist that month,
# so STRESS starts 1919 spread-only and gains legs as series begin.
# ---------------------------------------------------------------------------
def zscore(s: pd.Series) -> pd.Series:
    """Full-sample z-score, winsorized at +/-3 (descriptive use; lookahead OK)."""
    z = (s - s.mean()) / s.std()
    return z.clip(-3, 3)

df_m["z_spread"] = zscore(df_m["SPREAD"])
df_m["z_spread_chg"] = zscore(df_m["SPREAD_CHG"])
df_m["z_ebp"] = zscore(df_m["ebp"])
df_m["z_sloos"] = zscore(df_m["DRTSCILM_M"])
df_m["z_impulse"] = zscore(df_m["IMPULSE_AGG_M"])
df_m["z_ci_hh"] = zscore(df_m["CI_HH_M"])

# STRESS: high = credit tight / stressed
df_m["STRESS"] = df_m[["z_spread", "z_spread_chg", "z_ebp", "z_sloos"]].mean(axis=1)
# FROTH: high = ease / exuberance (leads reversals by ~2 years).
# NOTE: Greenwood-Hanson HY-issuance-share leg omitted (SIFMA licensing) — see header.
df_m["FROTH"] = pd.concat(
    [-df_m["z_spread"], df_m["z_impulse"], -df_m["z_sloos"], df_m["z_ci_hh"]],
    axis=1,
).mean(axis=1)

component_starts = {
    name: df_m[col].first_valid_index()
    for name, col in [("spread", "z_spread"), ("spread 12m chg", "z_spread_chg"),
                      ("EBP", "z_ebp"), ("SLOOS C&I", "z_sloos"),
                      ("agg impulse", "z_impulse"), ("HH impulse", "z_ci_hh")]
}

# ---------------------------------------------------------------------------
# Section 6 — sanity check: STRESS vs NFCI credit subindex, 1971+ overlap
# ---------------------------------------------------------------------------
# INVESTIGATED (2026-07): the pooled 1971+ corr runs ~0.50, below the ~0.7 rule
# of thumb, but every subperiod correlates HIGHER than the pooled figure
# (1971-90 ~0.55, 1990-2010 ~0.86, 2010+ ~0.66) — a between-era mean-shift
# artifact, not a construction bug. NFCICREDIT embeds rate-level-sensitive
# leverage components that were structurally elevated through the Volcker era
# (annual means ~+2.8 in 1980-82) while the spread-based STRESS legs peaked
# ~+1.9 there; once SLOOS joins in 1990 the two track closely. Both the pooled
# and the 1990+ correlations are printed below.
overlap = df_m[["STRESS", "NFCICREDIT_M"]].dropna()
nfci_corr = overlap["STRESS"].corr(overlap["NFCICREDIT_M"])
overlap_90 = overlap.loc["1990":]
nfci_corr_90 = overlap_90["STRESS"].corr(overlap_90["NFCICREDIT_M"])

# ---------------------------------------------------------------------------
# Section 7 — xlsx: Data_Monthly, Data_Quarterly, Summary
# ---------------------------------------------------------------------------
monthly_cols = [
    "BAA", "AAA", "SPREAD", "SPREAD_CHG", "BAA10YM",
    "gz_spread", "ebp", "est_prob",
    "DRTSCILM_M", "IMPULSE_AGG_M", "CI_HH_M",
    "z_spread", "z_spread_chg", "z_ebp", "z_sloos", "z_impulse", "z_ci_hh",
    "STRESS", "FROTH",
    "NFCI_M", "NFCICREDIT_M",
    "BUSLOANS", "BUSLOANS_YOY", "TOTBKCR_M", "TOTBKCR_M_YOY",
    "REVOLSL", "REVOLSL_YOY", "GS10", "FEDFUNDS",
]
quarterly_cols = [
    "DRTSCILM", "DRIWCIL", "DRTSCLCC",
    "DRBLACBS", "DRCCLACBS", "CORBLACBS",
    "QUSPAM770A", "GDP", "DPI",
    "FLOW_AGG", "FLOW_HH", "FLOW_CC",
    "FLOW4Q_AGG", "FLOW4Q_HH", "FLOW4Q_CC",
    "IMPULSE_AGG", "CI_HH", "CI_CC", "B_HH",
]

# Full-window stats (each series' own charted window = its full sample here)
summary_rows = {}
for name in ["SPREAD", "SPREAD_CHG", "ebp", "STRESS", "FROTH", "NFCICREDIT_M"]:
    summary_rows[f"{name} (full)"] = series_stats(df_m[name])
for name in ["DRTSCILM", "DRBLACBS", "DRCCLACBS", "CORBLACBS", "QUSPAM770A",
             "IMPULSE_AGG", "CI_HH", "CI_CC", "B_HH"]:
    summary_rows[f"{name} (full)"] = series_stats(df_q[name])
summary = pd.DataFrame(summary_rows).T[["min", "max", "mean", "median", "current"]]
summary["window"] = "full history (charted)"
summary["pct_rank_current"] = pd.NA

# 2006+ stats for the composites per repo convention, with percentile ranks
df_m06 = df_m[df_m.index >= pd.Timestamp(chart_start_2006)]
rows_06 = {}
for name in ["STRESS", "FROTH", "SPREAD", "ebp"]:
    rows_06[f"{name} (2006+)"] = series_stats(df_m06[name])
summary_06 = pd.DataFrame(rows_06).T[["min", "max", "mean", "median", "current"]]
summary_06["window"] = "2006+"
summary_06["pct_rank_current"] = [
    df_m06[name].dropna().rank(pct=True).iloc[-1] * 100
    for name in ["STRESS", "FROTH", "SPREAD", "ebp"]
]
summary = pd.concat([summary, summary_06])

xlsx_path = OUT_DIR / "credit_cycle.xlsx"
with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    df_m.reset_index()[["Date"] + monthly_cols].to_excel(
        writer, sheet_name="Data_Monthly", index=False)
    df_q.reset_index()[["Date"] + quarterly_cols].to_excel(
        writer, sheet_name="Data_Quarterly", index=False)
    summary.to_excel(writer, sheet_name="Summary")

# ---------------------------------------------------------------------------
# Section 8 — charts
# ---------------------------------------------------------------------------
full_start = df_m.index.min()          # 1919-01-01
recessions = get_recession_periods(fred, full_start, end)
recessions_06 = [r for r in recessions if r[1] >= pd.Timestamp(chart_start_2006)]
XLIM_FULL = (full_start, pd.Timestamp(end))
XLIM_06 = (pd.Timestamp(chart_start_2006), pd.Timestamp(end))


def combined_legend(ax1, ax2, loc="best"):
    """Merge legends of a twin-axis pair onto ax1 (frameless, house style)."""
    h1, l1 = ax1.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax1.legend(h1 + h2, l1 + l2, loc=loc, frameon=False)


# Chart 1 — price of credit: Baa-Aaa spread 1919+ with EBP overlaid from 1973.
# Both are percentage-point spreads, so one shared axis is fine.
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df_m.index, df_m["SPREAD"], color=MAIN, linewidth=1.3,
        label="Baa - Aaa spread (1919+)")
ax.plot(df_m.index, df_m["ebp"], color=NEG, linewidth=1.0,
        label="Excess bond premium (1973+)")
ax.set_xlim(*XLIM_FULL)
style_macro_chart(
    ax,
    title="Price of credit — Baa-Aaa spread and excess bond premium, 1919–present",
    ylabel="Percentage points",
    recessions=recessions,
    hlines=[{"y": 0.0}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "credit_price_panel.png", dpi=150)
plt.close(fig)

# Chart 2 — quantity of credit: credit/GDP level + aggregate credit impulse.
# X-axis starts 1947 (quarterly data start), not 1919 — no dead space.
XLIM_1947 = (pd.Timestamp("1947-01-01"), pd.Timestamp(end))
recessions_47 = [r for r in recessions if r[1] >= XLIM_1947[0]]
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8), sharex=True)
q770 = df_q["QUSPAM770A"].dropna()
ax1.plot(q770.index, q770, color=MAIN, linewidth=1.5,
         label="Private nonfinancial credit / GDP (BIS, QUSPAM770A)")
ax1.set_xlim(*XLIM_1947)
style_macro_chart(
    ax1,
    title="Quantity of credit — private nonfinancial credit to GDP, 1947–present",
    ylabel="% of GDP",
    recessions=recessions_47,
)
imp = df_q["IMPULSE_AGG"].dropna()
ax2.plot(imp.index, imp, color=MAIN, linewidth=1.5,
         label="Aggregate credit impulse (all domestic nonfinancial)")
ax2.set_xlim(*XLIM_1947)
style_macro_chart(
    ax2,
    title="Aggregate credit impulse — 4q change in smoothed net borrowing, % of GDP",
    ylabel="% of GDP",
    recessions=recessions_47,
    hlines=[{"y": 0.0}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "credit_quantity_panel.png", dpi=150)
plt.close(fig)

# Chart 3 — standards (SLOOS, leading-ish) + performance (lagging confirmation).
# DRIWCIL is willingness to LEND (higher = easier), so it is inverted to read
# in the same direction as the tightening series.
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8), sharex=True)
sloos_start = df_q["DRIWCIL"].dropna().index.min()
d1 = df_q["DRTSCILM"].dropna()
d2 = df_q["DRTSCLCC"].dropna()
d3 = df_q["DRIWCIL"].dropna()
ax1.plot(d1.index, d1, color=MAIN, linewidth=1.5,
         label="C&I standards, net % tightening (DRTSCILM)")
ax1.plot(d2.index, d2, color=LIGHT, linewidth=1.3,
         label="Credit card standards, net % tightening (DRTSCLCC)")
ax1.plot(d3.index, -d3, color=NEG, linewidth=1.0,
         label="Willingness to lend to consumers, INVERTED (-DRIWCIL)")
ax1.set_xlim(sloos_start, pd.Timestamp(end))
style_macro_chart(
    ax1,
    title="Lending standards (SLOOS diffusion: quarterly CHANGES, not levels)",
    ylabel="Net % (higher = tightening)",
    recessions=[r for r in recessions if r[1] >= sloos_start],
    hlines=[{"y": 0.0}],
)
p1 = df_q["DRBLACBS"].dropna()
p2 = df_q["DRCCLACBS"].dropna()
p3 = df_q["CORBLACBS"].dropna()
ax2.plot(p1.index, p1, color=MAIN, linewidth=1.5,
         label="Business loan delinquency rate")
ax2.plot(p2.index, p2, color=LIGHT, linewidth=1.3,
         label="Credit card delinquency rate")
ax2.plot(p3.index, p3, color=NEG, linewidth=1.0,
         label="Business loan charge-off rate")
ax2.set_xlim(sloos_start, pd.Timestamp(end))
style_macro_chart(
    ax2,
    title="Loan performance — LAGGING confirmation (2-4 quarters), not timing",
    ylabel="Percent of loans",
    recessions=[r for r in recessions if r[1] >= sloos_start],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "credit_standards_performance.png", dpi=150)
plt.close(fig)

# Chart 4 — composite gauge, full history. FROTH is plotted advanced 24 months
# (a froth reading at t is drawn at t+24m) because ease/exuberance historically
# leads credit reversals by ~2 years — the shifted line overlays the stress it
# preceded.
froth_shifted = df_m["FROTH"].dropna()
froth_shifted.index = froth_shifted.index + pd.DateOffset(months=24)
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df_m.index, df_m["STRESS"], color=MAIN, linewidth=1.4,
        label="STRESS composite (high = tight/stressed)")
ax.plot(froth_shifted.index, froth_shifted, color=NEG, linewidth=1.1,
        label="FROTH composite, advanced 24m (high = ease/exuberance)")
ax.set_xlim(*XLIM_FULL)
style_macro_chart(
    ax,
    title="Credit-cycle composites — STRESS vs FROTH (froth, advanced 24m), 1919–present",
    ylabel="Z-score (equal-weight mean)",
    recessions=recessions,
    hlines=[{"y": 0.0}],
)
starts_note = "Components join as data begins: " + "; ".join(
    f"{k} {v:%Y-%m}" for k, v in component_starts.items())
fig.text(0.01, 0.005, starts_note + ". HY-issuance-share leg omitted (SIFMA license).",
         fontsize=6.5, color="0.35")
fig.tight_layout(rect=(0, 0.02, 1, 1))
fig.savefig(OUT_DIR / "credit_cycle_composite.png", dpi=150)
plt.close(fig)

# Chart 5 — 2006+ companion, both composites UNSHIFTED.
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df_m06.index, df_m06["STRESS"], color=MAIN, linewidth=1.5,
        label="STRESS composite (high = tight/stressed)")
ax.plot(df_m06.index, df_m06["FROTH"], color=NEG, linewidth=1.2,
        label="FROTH composite, unshifted (high = ease/exuberance)")
ax.set_xlim(*XLIM_06)
style_macro_chart(
    ax,
    title="Credit-cycle composites — STRESS and FROTH (unshifted), 2006–present",
    ylabel="Z-score (equal-weight mean)",
    recessions=recessions_06,
    hlines=[{"y": 0.0}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "credit_composite_2006.png", dpi=150)
plt.close(fig)

# Chart 6 — STRESS vs the stock market (sp500_long_history sibling module).
spx_path = Path(__file__).resolve().parent / "output" / "sp500_long_history" / "sp500_monthly.csv"
if not spx_path.exists():
    print(f"SKIP credit_vs_spx.png: {spx_path} not found — "
          f"run python macro/sp500_long_history_pull.py first.")
else:
    spx = pd.read_csv(spx_path, parse_dates=["Date"]).set_index("Date")
    spx = spx[spx.index >= full_start]   # align to STRESS start (1919)
    dd = spx["drawdown"].copy()
    if dd.min() >= -1.01:                # fraction -> percent if needed
        dd = dd * 100
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8), sharex=True)
    ax1.plot(df_m.index, df_m["STRESS"], color=MAIN, linewidth=1.2,
             label="STRESS composite (left)")
    ax1.set_xlim(*XLIM_FULL)
    style_macro_chart(
        ax1,
        title="Credit STRESS composite vs S&P 500 total return (log), 1919–present",
        ylabel="Z-score",
        recessions=recessions,
        hlines=[{"y": 0.0}],
    )
    ax1b = ax1.twinx()
    ax1b.plot(spx.index, spx["TR"], color=LIGHT, linewidth=1.2,
              label="S&P 500 nominal TR index (right, log)")
    ax1b.set_yscale("log")
    ax1b.set_ylabel("TR index (log)")
    ax1b.spines["top"].set_visible(False)
    combined_legend(ax1, ax1b)
    ax2.plot(df_m.index, df_m["STRESS"], color=MAIN, linewidth=1.2,
             label="STRESS composite (left)")
    ax2.set_xlim(*XLIM_FULL)
    style_macro_chart(
        ax2,
        title="Credit STRESS composite vs S&P 500 drawdown, 1919–present",
        ylabel="Z-score",
        recessions=recessions,
        hlines=[{"y": 0.0}],
    )
    ax2b = ax2.twinx()
    ax2b.plot(spx.index, dd, color=NEG, linewidth=1.0,
              label="S&P 500 real-TR drawdown, % (right)")
    ax2b.set_ylabel("Drawdown (%)")
    ax2b.spines["top"].set_visible(False)
    # "best" lands the legend on the drawdown line's 0% ceiling; lower left is clear.
    combined_legend(ax2, ax2b, loc="lower left")
    fig.tight_layout()
    fig.savefig(OUT_DIR / "credit_vs_spx.png", dpi=150)
    plt.close(fig)

# Chart 7 — STRESS vs rates (10y Treasury and fed funds), twin axis.
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df_m.index, df_m["STRESS"], color=MAIN, linewidth=1.3,
        label="STRESS composite (left)")
ax.set_xlim(*XLIM_FULL)
style_macro_chart(
    ax,
    title="Credit STRESS composite vs interest rates, 1919–present",
    ylabel="Z-score",
    recessions=recessions,
    hlines=[{"y": 0.0}],
)
axb = ax.twinx()
axb.plot(df_m.index, df_m["GS10"], color=LIGHT, linewidth=1.2,
         label="10y Treasury yield, % (right, 1953+)")
axb.plot(df_m.index, df_m["FEDFUNDS"], color=POS, linewidth=1.0,
         label="Fed funds rate, % (right, 1954+)")
axb.set_ylabel("Percent")
axb.spines["top"].set_visible(False)
combined_legend(ax, axb)
fig.tight_layout()
fig.savefig(OUT_DIR / "credit_vs_rates.png", dpi=150)
plt.close(fig)

# Chart 8 — household credit impulses: bridge series shared with the
# consumer_pullforward module.
ci_hh = df_q["CI_HH"].dropna()
ci_cc = df_q["CI_CC"].dropna()
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(ci_hh.index, ci_hh, color=MAIN, linewidth=1.5,
        label="Household credit impulse (all debt, % of DPI)")
ax.plot(ci_cc.index, ci_cc, color=LIGHT, linewidth=1.3,
        label="Consumer-credit impulse (% of DPI)")
ax.set_xlim(ci_hh.index.min(), pd.Timestamp(end))
style_macro_chart(
    ax,
    title="Household credit impulses — bridge series (shared with consumer_pullforward)",
    ylabel="% of disposable income",
    recessions=[r for r in recessions if r[1] >= ci_hh.index.min()],
    hlines=[{"y": 0.0}],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "credit_impulse_household.png", dpi=150)
plt.close(fig)

# ---------------------------------------------------------------------------
# Section 9 — print summary + sanity checks
# ---------------------------------------------------------------------------
spread_now = df_m["SPREAD"].dropna().iloc[-1]
stress_now = df_m["STRESS"].dropna().iloc[-1]
froth_now = df_m["FROTH"].dropna().iloc[-1]
peak_spread = df_m["SPREAD"].idxmax()
print(f"Data start (monthly):     {df_m.index.min().date()}")
print(f"Data end:                 {end.date()}")
print(f"Monthly rows:             {len(df_m)}")
print(f"Quarterly rows:           {len(df_q)}")
print(f"Latest Baa-Aaa spread:    {spread_now:.2f}pp "
      f"(max {df_m['SPREAD'].max():.2f}pp on {peak_spread:%Y-%m})")
print(f"Latest EBP:               {df_m['ebp'].dropna().iloc[-1]:+.2f} "
      f"({df_m['ebp'].dropna().index[-1]:%Y-%m})")
print(f"Latest DRTSCILM:          {df_q['DRTSCILM'].dropna().iloc[-1]:+.1f} "
      f"(peak {df_q['DRTSCILM'].max():.1f} on {df_q['DRTSCILM'].idxmax():%Y-%m})")
print(f"Latest QUSPAM770A:        {df_q['QUSPAM770A'].dropna().iloc[-1]:.1f}% of GDP")
print(f"Latest agg impulse:       {df_q['IMPULSE_AGG'].dropna().iloc[-1]:+.2f}% of GDP")
print(f"Latest CI_HH / CI_CC:     {df_q['CI_HH'].dropna().iloc[-1]:+.2f} / "
      f"{df_q['CI_CC'].dropna().iloc[-1]:+.2f} % of DPI")
bhh = df_q["B_HH"].dropna()
print(f"B_HH peak/trough/latest:  {bhh.max():.1f}% ({bhh.idxmax():%Y-%m}) / "
      f"{bhh.min():.1f}% ({bhh.idxmin():%Y-%m}) / {bhh.iloc[-1]:.1f}%")
print(f"Latest STRESS:            {stress_now:+.2f} "
      f"({summary.loc['STRESS (2006+)', 'pct_rank_current']:.0f}th pctile since 2006)")
print(f"Latest FROTH:             {froth_now:+.2f} "
      f"({summary.loc['FROTH (2006+)', 'pct_rank_current']:.0f}th pctile since 2006)")
print(f"SANITY corr(STRESS, NFCICREDIT) 1971+ monthly: {nfci_corr:.3f} "
      f"({'OK, >= 0.7' if nfci_corr >= 0.7 else 'below 0.7 — Volcker-era mean shift, see Section 6 comment'})")
print(f"SANITY corr(STRESS, NFCICREDIT) 1990+ monthly: {nfci_corr_90:.3f} "
      f"({'OK, >= 0.7' if nfci_corr_90 >= 0.7 else 'LOW — investigate'})")
print("SANITY — STRESS peaks by stress era (expect 2008-09 highest post-1970):")
eras = [("1974-75", "1973-06", "1976-06"), ("1980-82", "1979-06", "1983-06"),
        ("1990", "1989-01", "1992-01"), ("2001-02", "2000-06", "2003-06"),
        ("2008-09", "2007-06", "2010-06"), ("2020", "2020-01", "2021-06")]
for label, s, e in eras:
    win = df_m.loc[s:e, "STRESS"]
    print(f"  {label:8s} max {win.max():+.2f} ({win.idxmax():%Y-%m})")
print(f"Wrote {xlsx_path.name}, ebp_csv.csv, credit_price_panel.png, "
      f"credit_quantity_panel.png, credit_standards_performance.png, "
      f"credit_cycle_composite.png, credit_composite_2006.png, credit_vs_spx.png, "
      f"credit_vs_rates.png, credit_impulse_household.png to {OUT_DIR}")
