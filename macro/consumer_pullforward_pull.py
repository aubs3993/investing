# US consumer pull-forward & sustainability gauge — ex-post trend gaps,
# stock-adjustment, excess savings, an ex-ante survey leg, and a composite.
# Data: FRED (verified 2026-07-07) + UMich Surveys of Consumers data archive.
# See macro/ANALYSES.md section 3 for the episode library and design notes.
#
# COMPONENTS
#   1. Durables trend gap: log real durables (DDURRA3M086SBEA) minus a
#      log-linear OLS trend fit on the FROZEN 2011-01..2019-12 window only,
#      extrapolated forward. HP-style filters are deliberately NOT used —
#      endpoint bias plus they absorb exactly the persistent deviation we want
#      to measure (Hamilton critique); the frozen pre-COVID window matches the
#      SF Fed excess-savings method. The CUMULATIVE gap (sum of monthly % gaps,
#      units = %-months; divide by 12 for %-years) is the payback readout:
#      pull-forward is repaid only when it re-crosses zero.
#      A LINEAR-level trend companion (DUR_GAP_LIN) is computed alongside:
#      2011-19 log growth is 6.2%/yr, so six years of exponential
#      extrapolation compounds far above a straight-line path, and the two
#      variants give OPPOSITE payback verdicts today (see caveats). The
#      published "+25-30% above pre-COVID trend, never paid back" magnitudes
#      reproduce under the LINEAR trend, not the log-linear one.
#   2. Auto replacement gap: ALTSALES minus REPLACEMENT_SAAR (13.0M, see
#      constant below).
#   3. Stock-adjustment ratio (quarterly): gross durables purchases over
#      depreciation-implied replacement demand (annual BEA Fixed Assets dep
#      rate x prior-quarter Z.1 stock). VERIFIED CORRECTION: BOGZ1FA155111005Q
#      is the Z.1 NET transactions flow (gross purchases minus current-cost
#      depreciation — cum flows 1995-99 = $713bn ~ stock change $587bn ~
#      PCEDG-minus-dep $520bn, vs gross PCEDG $3.8T), NOT gross purchases.
#      Gross is reconstructed as net flow + replacement demand, so the ratio
#      = 1 + net/replacement. Ratio > 1 = buying above replacement (net
#      additions to the stock). Both legs nominal, so price effects cancel.
#   4. Front-running: retail inventories/sales vs its 2011-19 mean, and YoY %
#      of nominal consumer-goods imports (importers stocking ahead of tariffs).
#   5. Excess savings, SF Fed replication: saving level = PSAVERT*DSPI/100
#      (monthly, SAAR $bn); linear trend fit Mar-2016..Feb-2020; cumulative
#      sum of (actual - trend)/12 from Mar-2020 (SAAR flows -> /12 to
#      accumulate in $). Anchors: peak ~ +$2.1T around Aug-2021, crosses zero
#      ~ Dec-2024 on the current data vintage (SF Fed's original estimate was
#      ~Mar-2024; later annual PSAVERT/DSPI revisions raised measured saving),
#      increasingly negative after.
#   6. Ex-ante survey leg: UMich SCA reason-split tables 36/38/42 — share
#      saying now is a good time to buy durables/vehicles/houses BECAUSE
#      "prices will increase" (buy-in-advance share). Monthly 1978+.
#   7. CI_CC consumer-credit impulse bridge series — IDENTICAL formula to the
#      credit_cycle module (4q-mean flow minus its 4q lag, over 4q-mean DPI):
#      the two analyses split the credit block, see ANALYSES.md section 2.
#
# COMPOSITE (documented sign conventions also live in the chart legends):
#   PULL-FORWARD sub-index (higher = demand borrowed from the future) =
#     mean of available z-scores of: durables gap, auto gap, stock-adjustment
#     ratio, front-running average, buy-in-advance intent average.
#   SUSTAINABILITY-STRAIN sub-index (higher = financing stretched) =
#     mean of z(-(PSAVERT - 2011-19 mean)), z(REVOLSL YoY - DSPI YoY),
#     z(card delinquency rate), z(CI_CC).
#   Overall gauge = mean of the two sub-indexes. All z-scores standardized
#   over the 2006+ charted window per repo convention. (Component history
#   would support a composite back to ~1992 — RETAILIRSA start — but charts
#   follow the standard 2006 window; the episodes of interest are 2009+.)
#
# CAVEATS
# - Frozen-trend caveat: the 2011-19 trend gap grows less meaningful the
#   further we get from 2020 (any trend-growth change compounds). A rolling
#   10-year z-score of the gap (GAP_Z_ROLL10Y) is reported alongside as the
#   drift-robust cross-check. BACKCAST readings before the 2011 fit start are
#   artifacts (the 2006-10 gap reads ~+30% only because 2006-11 realized
#   growth was far below the 2011-19 fit) — they bias the composite's
#   durables leg upward pre-2011.
# - Payback verdict is TREND-METHOD-DEPENDENT (annotated on the chart):
#   under the linear-level trend the cumulative gap never re-crosses zero
#   (COVID durables excess partly a permanent level shift, per the published
#   result); under the primary log-linear trend the excess is fully repaid by
#   ~2023-24 and the cumulative gap is now negative.
# - BEA Fixed Assets are ANNUAL with a ~1-year lag (last obs 2024): the
#   depreciation rate is linearly interpolated to quarterly and held flat past
#   the last Fixed-Assets year.
# - REPLACEMENT_SAAR is a hardcoded constant — refresh annually (see below).
# - BNPL balances are essentially invisible in REVOLSL, so credit-financed
#   spending is somewhat understated in the strain block.
# - UMich survey moved phone->web during 2024; treat level comparisons across
#   that break with care (marked on the intent charts). Reason shares allow
#   multiple mentions — columns can sum above 100; never renormalize.
# - CDSP on FRED now starts 2005 (older vintages reached 1980); it is pulled
#   for the data sheet only and is not a composite component.
# - Intent chart runs 1978+ — INTENTIONAL EXCEPTION to the repo's 2006 start
#   convention: the 1979-80 advance-buying episode (shares ~45-53) is the
#   historical yardstick for the 2024-25 tariff-anticipation spike (~20-25).
#   A 2006+ companion chart is emitted per convention.

from datetime import datetime
from pathlib import Path
import re
import sys

import matplotlib.pyplot as plt
import numpy as np
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

# Light-vehicle replacement demand, million units SAAR: US fleet 289M x 4.5%
# scrappage rate ~= 13.0M (S&P Global Mobility, May-2025 release). Refresh
# annually when S&P publishes the new fleet/scrappage figures.
REPLACEMENT_SAAR = 13.0

UMICH_URL = "https://data.sca.isr.umich.edu/data-archive/mine.php"
# Site 403s default python-requests UAs; present a browser UA.
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"
    )
}
# UMich SCA reason-split tables: durables (36), vehicles (38), houses (42).
UMICH_TABLES = {"INTENT_DURABLES": 36, "INTENT_VEHICLES": 38, "INTENT_HOMES": 42}

# Policy events that produced measurable pull-forward (see ANALYSES.md).
EVENTS = [
    (pd.Timestamp("2009-07-01"), "Cash for Clunkers"),
    (pd.Timestamp("2009-11-01"), "Homebuyer credit #1"),
    (pd.Timestamp("2010-06-01"), "Homebuyer credit #2"),
    (pd.Timestamp("2020-04-01"), "CARES checks"),
    (pd.Timestamp("2021-01-01"), "Stimulus #2"),
    (pd.Timestamp("2021-03-01"), "ARP checks"),
    (pd.Timestamp("2025-03-01"), "Tariff front-running"),
    (pd.Timestamp("2025-09-01"), "EV credit expiry"),
]
# UMich survey transitioned phone->web across 2024 (fully web by mid-2024).
UMICH_BREAK = pd.Timestamp("2024-07-01")

fred = get_fred_client()

end = datetime.today()
chart_start = datetime(2006, 1, 1)
# Buffer well before 2006: the 2011-19 frozen trend windows sit inside 2006+,
# but rolling z (10y), YoY changes and 4q MAs need earlier data, and the
# stock-adjustment ratio is meaningful from the early 1990s.
data_start = datetime(1990, 1, 1)

monthly_series = {
    "DUR_REAL": "DDURRA3M086SBEA",  # real PCE durables quantity index, 2017=100
    "PCEDG": "PCEDG",               # nominal durables, $bn SAAR
    "ALTSALES": "ALTSALES",         # light-vehicle sales, M units SAAR
    "TOTALSA": "TOTALSA",           # total vehicle sales, M units SAAR
    "RETAILIRSA": "RETAILIRSA",     # retail inventories/sales ratio (1992+)
    "PSAVERT": "PSAVERT",           # personal saving rate, %
    "DSPI": "DSPI",                 # disposable personal income, $bn SAAR (monthly)
    "REVOLSL": "REVOLSL",           # revolving consumer credit, $M
}
quarterly_series = {
    "IMP_CONS": "A652RC1Q027SBEA",  # nominal consumer-goods imports, $bn SAAR
    "DRCCLACBS": "DRCCLACBS",       # card delinquency rate, % (1991+)
    "CDSP": "CDSP",                 # consumer debt service % of DPI (2005+ on FRED)
    "DPI_Q": "DPI",                 # disposable personal income, quarterly, $bn SAAR
    "CC_FLOW": "HNOCCLQ027S",       # consumer credit borrowing flow, $M SAAR
    "DUR_STOCK": "BOGZ1LM155111005Q",  # Z.1 consumer durables stock, $M
    "DUR_PURCH": "BOGZ1FA155111005Q",  # Z.1 durables NET transactions, $M SAAR
}
annual_series = {
    "FA_STOCK": "K1CTOTL1CD000",   # BEA FA net stock of consumer durables, $M
    "FA_DEP": "M1CTOTL1CD000",     # BEA FA depreciation of consumer durables, $M
}

df_m = pull_series(fred, monthly_series, data_start, end)
df_q = pull_series(fred, quarterly_series, data_start, end)
df_a = pull_series(fred, annual_series, data_start, end)

# ---------------------------------------------------------------------------
# (1) Durables trend gap vs frozen 2011-01..2019-12 log-linear trend
# ---------------------------------------------------------------------------
FIT_START, FIT_END = pd.Timestamp("2011-01-01"), pd.Timestamp("2019-12-01")
t_months = (df_m["Date"].dt.year - 2011) * 12 + (df_m["Date"].dt.month - 1)
fit_mask = (df_m["Date"] >= FIT_START) & (df_m["Date"] <= FIT_END)
slope, intercept = np.polyfit(
    t_months[fit_mask], np.log(df_m.loc[fit_mask, "DUR_REAL"]), 1
)
df_m["DUR_TREND"] = np.exp(intercept + slope * t_months)
# Log gap in % (log-points x100 ~= % for these magnitudes).
df_m["DUR_GAP"] = (np.log(df_m["DUR_REAL"]) - np.log(df_m["DUR_TREND"])) * 100
# Linear-level trend companion on the same frozen window (see header: the two
# variants disagree on the payback verdict; published magnitudes are linear).
slope_lin, int_lin = np.polyfit(
    t_months[fit_mask], df_m.loc[fit_mask, "DUR_REAL"], 1
)
df_m["DUR_TREND_LIN"] = int_lin + slope_lin * t_months
df_m["DUR_GAP_LIN"] = (df_m["DUR_REAL"] / df_m["DUR_TREND_LIN"] - 1) * 100
# Cumulative gaps from Mar-2020 (COVID episode start, mirroring the excess-
# savings accumulation start). Units: %-months (divide by 12 for %-years).
covid = df_m["Date"] >= pd.Timestamp("2020-03-01")
for gap_col, cum_col in (("DUR_GAP", "DUR_GAP_CUM"),
                         ("DUR_GAP_LIN", "DUR_GAP_CUM_LIN")):
    df_m[cum_col] = np.nan
    df_m.loc[covid, cum_col] = df_m.loc[covid, gap_col].cumsum()
# Drift-robust cross-check: z of the gap vs its own trailing 10y distribution.
roll = df_m["DUR_GAP"].rolling(120, min_periods=60)
df_m["GAP_Z_ROLL10Y"] = (df_m["DUR_GAP"] - roll.mean()) / roll.std()

# ---------------------------------------------------------------------------
# (2) Auto replacement gap
# ---------------------------------------------------------------------------
df_m["AUTO_GAP"] = df_m["ALTSALES"] - REPLACEMENT_SAAR

# ---------------------------------------------------------------------------
# (3) Stock-adjustment ratio (quarterly)
# ---------------------------------------------------------------------------
# Annual dep rate delta = FA depreciation / FA net stock, linearly (time-)
# interpolated to quarterly dates and held flat past the last Fixed-Assets
# year (annual publication lag).
delta_a = pd.Series(
    (df_a["FA_DEP"] / df_a["FA_STOCK"]).values, index=df_a["Date"].values
)
q_idx = pd.DatetimeIndex(df_q["Date"])
delta_q = (
    delta_a.reindex(delta_a.index.union(q_idx))
    .interpolate(method="time")
    .ffill()
    .reindex(q_idx)
)
df_q["DELTA_Q"] = delta_q.values
# DUR_PURCH (BOGZ1FA155111005Q) is the NET transactions flow (SAAR, gross
# purchases minus current-cost depreciation — see verified correction in the
# header), so gross purchases are reconstructed as net + replacement demand
# and the ratio is 1 + net/replacement. Both legs annualized nominal $M.
df_q["REPLACEMENT_DEMAND"] = df_q["DELTA_Q"] * df_q["DUR_STOCK"].shift(1)
df_q["GROSS_PURCH"] = df_q["DUR_PURCH"] + df_q["REPLACEMENT_DEMAND"]
df_q["STOCKADJ_RATIO"] = df_q["GROSS_PURCH"] / df_q["REPLACEMENT_DEMAND"]

# ---------------------------------------------------------------------------
# (4) Front-running: inventories/sales deviation + consumer-goods imports YoY
# ---------------------------------------------------------------------------
retail_mean_1119 = df_m.loc[fit_mask, "RETAILIRSA"].mean()
df_m["RETAIL_DEV"] = df_m["RETAILIRSA"] - retail_mean_1119
df_q["IMP_YOY"] = df_q["IMP_CONS"].pct_change(4) * 100

# ---------------------------------------------------------------------------
# (5) Excess savings (SF Fed method)
# ---------------------------------------------------------------------------
df_m["SAVING_LEVEL"] = df_m["PSAVERT"] * df_m["DSPI"] / 100  # $bn SAAR
ES_FIT_START, ES_FIT_END = pd.Timestamp("2016-03-01"), pd.Timestamp("2020-02-01")
es_mask = (df_m["Date"] >= ES_FIT_START) & (df_m["Date"] <= ES_FIT_END)
es_t = (df_m["Date"].dt.year - 2016) * 12 + (df_m["Date"].dt.month - 3)
es_slope, es_int = np.polyfit(es_t[es_mask], df_m.loc[es_mask, "SAVING_LEVEL"], 1)
df_m["SAVING_TREND"] = es_int + es_slope * es_t
# Monthly flows are SAAR -> divide by 12 to accumulate actual $ saved.
df_m["EXCESS_SAVINGS"] = np.where(
    covid, (df_m["SAVING_LEVEL"] - df_m["SAVING_TREND"]) / 12, np.nan
)
df_m.loc[covid, "EXCESS_SAVINGS"] = df_m.loc[covid, "EXCESS_SAVINGS"].cumsum()

# ---------------------------------------------------------------------------
# (6) UMich SCA buy-in-advance intent (tables 36/38/42), cached CSVs
# ---------------------------------------------------------------------------
OUT_DIR = resolve_output_dir(__file__, "consumer_pullforward")
(OUT_DIR / ".gitkeep").touch()


def fetch_umich_share(table_no: int, cache_path: Path) -> pd.Series | None:
    """Fetch one SCA reason-split table; return the 'Good Time — Prices will
    increase' (buy-in-advance) share as a Date-indexed monthly Series.

    Falls back to the cached CSV on fetch failure; returns None if neither is
    available. Header cells embed literal '<br>'/'<Br>'; data rows end with a
    trailing comma. Shares allow multiple mentions (columns can sum above
    100) — never renormalized.
    """
    text = None
    try:
        resp = requests.post(
            UMICH_URL,
            headers=HEADERS,
            timeout=60,
            data={
                "table": str(table_no),
                "year": "1978",
                "qorm": "M",
                "order": "asc",
                "format": "Comma-Separated (CSV)",
            },
        )
        resp.raise_for_status()
        if "csv" not in resp.headers.get("Content-Type", ""):
            raise RuntimeError(f"unexpected content type for table {table_no}")
        text = resp.text
        cache_path.write_text(text, encoding="utf-8")
    except Exception as exc:  # noqa: BLE001 — any fetch problem -> cache
        print(f"WARNING: UMich table {table_no} fetch failed ({exc}); trying cache")
        if cache_path.exists():
            text = cache_path.read_text(encoding="utf-8")
    if text is None:
        return None
    lines = text.splitlines()
    header = re.sub(r"<br>", " ", lines[1], flags=re.I)
    cols = [c.strip() for c in header.split(",")]
    target_idx = next(
        (
            i
            for i, c in enumerate(cols)
            if "good time" in c.lower() and "prices will increase" in c.lower()
        ),
        None,
    )
    if target_idx is None:
        print(f"WARNING: UMich table {table_no}: advance-buying column not found")
        return None
    rows = [ln.split(",") for ln in lines[2:] if ln.strip()]
    month = pd.to_numeric([r[0] for r in rows], errors="coerce")
    year = pd.to_numeric([r[1] for r in rows], errors="coerce")
    vals = pd.to_numeric([r[target_idx] for r in rows], errors="coerce")
    dates = pd.to_datetime(
        {"year": year, "month": month, "day": np.ones(len(rows))}, errors="coerce"
    )
    s = pd.Series(vals, index=dates).dropna()
    return s[s.index.notna()].sort_index()


intent_shares = {}
for name, table_no in UMICH_TABLES.items():
    s = fetch_umich_share(table_no, OUT_DIR / f"umich_table{table_no}.csv")
    if s is not None:
        intent_shares[name] = s
df_intent = pd.DataFrame(intent_shares) if intent_shares else None
if df_intent is not None:
    df_intent.index.name = "Date"
    df_intent["INTENT_AVG"] = df_intent.mean(axis=1)
    df_intent = df_intent.reset_index()
    df_m = df_m.merge(
        df_intent[["Date", "INTENT_AVG"]], on="Date", how="left"
    )
else:
    print("WARNING: no UMich data (fetch failed, no cache) — intent chart and "
          "composite leg skipped")

# ---------------------------------------------------------------------------
# (7) CI_CC consumer-credit impulse — bridge series shared with credit_cycle
# ---------------------------------------------------------------------------
cc_ma4 = (df_q["CC_FLOW"] / 1000).rolling(4).mean()  # $M SAAR -> $bn, 4q mean
df_q["CI_CC"] = (cc_ma4 - cc_ma4.shift(4)) / df_q["DPI_Q"].rolling(4).mean() * 100

# --- Quarterly components forward-filled onto the monthly grid -------------
df_m = pd.merge_asof(
    df_m.sort_values("Date"),
    df_q[["Date", "STOCKADJ_RATIO", "IMP_YOY", "DRCCLACBS", "CI_CC"]]
    .sort_values("Date"),
    on="Date",
    direction="backward",
)

# ---------------------------------------------------------------------------
# Composite (z-scores standardized over the 2006+ charted window)
# ---------------------------------------------------------------------------
df_m["REVOLSL_YOY"] = df_m["REVOLSL"].pct_change(12) * 100
df_m["DSPI_YOY"] = df_m["DSPI"].pct_change(12) * 100
psavert_mean_1119 = df_m.loc[fit_mask, "PSAVERT"].mean()
df_m["PSAVERT_DEV"] = df_m["PSAVERT"] - psavert_mean_1119

dfm06 = df_m[df_m["Date"] >= pd.Timestamp(chart_start)].reset_index(drop=True)
df_q = df_q[df_q["Date"] >= pd.Timestamp(chart_start)].reset_index(drop=True)


def z06(col: str) -> pd.Series:
    """Z-score of a dfm06 column over the 2006+ charted window."""
    s = dfm06[col]
    return (s - s.mean()) / s.std()


dfm06["Z_DUR_GAP"] = z06("DUR_GAP")
dfm06["Z_AUTO_GAP"] = z06("AUTO_GAP")
dfm06["Z_STOCKADJ"] = z06("STOCKADJ_RATIO")
dfm06["FRONTRUN_AVG"] = pd.concat(
    [z06("RETAIL_DEV"), z06("IMP_YOY")], axis=1
).mean(axis=1)
dfm06["Z_FRONTRUN"] = z06("FRONTRUN_AVG")
pull_cols = ["Z_DUR_GAP", "Z_AUTO_GAP", "Z_STOCKADJ", "Z_FRONTRUN"]
if "INTENT_AVG" in dfm06.columns:
    # 2024 phone->web break noted rather than z-ing within regime: the web
    # regime is too short (~2y) for a stable own-regime z.
    dfm06["Z_INTENT"] = z06("INTENT_AVG")
    pull_cols.append("Z_INTENT")
dfm06["PULLFORWARD"] = dfm06[pull_cols].mean(axis=1)

dfm06["Z_SAVERT"] = z06("PSAVERT_DEV") * -1          # low saving rate = strain
dfm06["CREDIT_INC_GAP"] = dfm06["REVOLSL_YOY"] - dfm06["DSPI_YOY"]
dfm06["Z_CREDIT_GAP"] = z06("CREDIT_INC_GAP")
dfm06["Z_DELINQ"] = z06("DRCCLACBS")
dfm06["Z_CICC"] = z06("CI_CC")
strain_cols = ["Z_SAVERT", "Z_CREDIT_GAP", "Z_DELINQ", "Z_CICC"]
dfm06["STRAIN"] = dfm06[strain_cols].mean(axis=1)

dfm06["GAUGE"] = dfm06[["PULLFORWARD", "STRAIN"]].mean(axis=1)

# ---------------------------------------------------------------------------
# Summary + xlsx
# ---------------------------------------------------------------------------
summary_cols = [
    "DUR_GAP", "DUR_GAP_LIN", "DUR_GAP_CUM", "DUR_GAP_CUM_LIN",
    "GAP_Z_ROLL10Y", "ALTSALES", "AUTO_GAP",
    "STOCKADJ_RATIO", "RETAILIRSA", "IMP_YOY", "PSAVERT", "EXCESS_SAVINGS",
    "REVOLSL_YOY", "DRCCLACBS", "CI_CC", "PULLFORWARD", "STRAIN", "GAUGE",
]
if "INTENT_AVG" in dfm06.columns:
    summary_cols.insert(8, "INTENT_AVG")
summary_rows = {c: series_stats(dfm06[c]) for c in summary_cols}
summary = pd.DataFrame(summary_rows).T[["min", "max", "mean", "median", "current"]]
summary.loc["PSAVERT_mean_2011_19"] = [pd.NA] * 4 + [psavert_mean_1119]
summary.loc["RETAILIRSA_mean_2011_19"] = [pd.NA] * 4 + [retail_mean_1119]
summary.loc["REPLACEMENT_SAAR"] = [pd.NA] * 4 + [REPLACEMENT_SAAR]

xlsx_path = OUT_DIR / "consumer_pullforward.xlsx"
with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    dfm06.to_excel(writer, sheet_name="Data_Monthly", index=False)
    df_q.to_excel(writer, sheet_name="Data_Quarterly", index=False)
    if df_intent is not None:
        df_intent.to_excel(writer, sheet_name="Data_Intent", index=False)
    summary.to_excel(writer, sheet_name="Summary")

# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------
recessions = get_recession_periods(fred, chart_start, end)
XLIM = (pd.Timestamp(chart_start), pd.Timestamp(end))


def add_event_lines(ax, xlim) -> None:
    """Dotted vlines + small rotated labels for the pull-forward event dates."""
    y_top = ax.get_ylim()[1]
    for ts, label in EVENTS:
        if not (xlim[0] <= ts <= xlim[1]):
            continue
        ax.axvline(ts, color="0.35", linestyle=":", linewidth=0.9, zorder=1)
        ax.text(ts, y_top, " " + label, rotation=90, fontsize=6.5,
                va="top", ha="right", color="0.35")


# Chart 1 — durables level vs frozen trend (top) + % gap and cumulative gap
# (bottom, twin axes).
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8), sharex=True)
ax1.plot(dfm06["Date"], dfm06["DUR_REAL"], color="#1f3b73", linewidth=1.6,
         label="Real PCE durables (2017=100)")
ax1.plot(dfm06["Date"], dfm06["DUR_TREND"], color="#c0392b", linewidth=1.3,
         linestyle="--", label="Frozen log-linear trend (fit 2011–2019)")
ax1.plot(dfm06["Date"], dfm06["DUR_TREND_LIN"], color="#2ca02c", linewidth=1.1,
         linestyle="--", label="Frozen linear trend (same window)")
ax1.set_xlim(*XLIM)
style_macro_chart(
    ax1,
    title="Real durables consumption vs frozen pre-COVID trends, 2006–present",
    ylabel="Quantity index (2017=100)",
    recessions=recessions,
)
ax2b = ax2.twinx()
ax2b.plot(dfm06["Date"], dfm06["DUR_GAP_CUM"], color="#9ec5e8", linewidth=1.6,
          label="Cum. gap since Mar-2020, log-linear (%-months, right)")
ax2b.plot(dfm06["Date"], dfm06["DUR_GAP_CUM_LIN"], color="#9ec5e8",
          linewidth=1.4, linestyle="--",
          label="Cum. gap since Mar-2020, linear (%-months, right)")
ax2b.axhline(0.0, color="#9ec5e8", linestyle=":", linewidth=0.8)
ax2b.set_ylabel("%-months")
ax2.plot(dfm06["Date"], dfm06["DUR_GAP"], color="#1f3b73", linewidth=1.4,
         label="Gap vs log-linear trend (%, left)")
ax2.plot(dfm06["Date"], dfm06["DUR_GAP_LIN"], color="#2ca02c", linewidth=1.1,
         label="Gap vs linear trend (%, left)")
ax2.set_xlim(*XLIM)
style_macro_chart(
    ax2,
    title="Durables trend gap — payback verdict is trend-method-dependent: "
          "linear = never repaid; log-linear = repaid by ~2024",
    ylabel="Gap (%)",
    # Clipped at +40%: the 2006-08 backcast readings (linear gap up to +80%)
    # are pre-fit-window artifacts and would compress the COVID episode.
    ylim=(-35, 40),
    recessions=recessions,
    hlines=[{"y": 0.0}],
)
h1, l1 = ax2.get_legend_handles_labels()
h2, l2 = ax2b.get_legend_handles_labels()
ax2.legend(h1 + h2, l1 + l2, loc="upper left", frameon=False, fontsize=8)
fig.tight_layout()
fig.savefig(OUT_DIR / "durables_trend_gap.png", dpi=150)
plt.close(fig)

# Chart 2 — light-vehicle sales vs replacement demand, with event vlines.
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(dfm06["Date"], dfm06["ALTSALES"], color="#1f3b73", linewidth=1.5,
        label="Light-vehicle sales (SAAR)")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="Light-vehicle sales vs replacement demand, 2006–present",
    ylabel="Million units, SAAR",
    recessions=recessions,
    hlines=[{"y": REPLACEMENT_SAAR, "color": "#c0392b",
             "label": f"Replacement demand ~{REPLACEMENT_SAAR:.1f}M "
                      "(fleet x scrappage, S&P Mobility 2025)"}],
)
add_event_lines(ax, XLIM)
fig.tight_layout()
fig.savefig(OUT_DIR / "auto_replacement_gap.png", dpi=150)
plt.close(fig)

# Chart 3 — stock-adjustment ratio (quarterly).
ratio_mean06 = df_q["STOCKADJ_RATIO"].mean()
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(df_q["Date"], df_q["STOCKADJ_RATIO"], color="#1f3b73", linewidth=1.6,
        label="Durables purchases / depreciation-implied replacement")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="Durables stock-adjustment ratio — >1 = adding to the stock, 2006–present",
    ylabel="Ratio",
    recessions=recessions,
    hlines=[
        {"y": 1.0, "color": "#c0392b", "label": "Replacement only (1.0)"},
        {"y": ratio_mean06, "label": f"2006+ mean ({ratio_mean06:.2f})"},
    ],
)
fig.tight_layout()
fig.savefig(OUT_DIR / "stock_adjustment_ratio.png", dpi=150)
plt.close(fig)

# Chart 4 — cumulative excess savings ($bn), peak and zero-cross annotated.
es = dfm06.dropna(subset=["EXCESS_SAVINGS"])
es_peak = es.loc[es["EXCESS_SAVINGS"].idxmax()]
below = es[es["EXCESS_SAVINGS"] < 0]
es_cross = below.iloc[0] if len(below) else None
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(es["Date"], es["EXCESS_SAVINGS"], color="#1f3b73", linewidth=1.8,
        label="Cumulative excess savings (SF Fed method, trend fit 2016–Feb-2020)")
ax.set_xlim(pd.Timestamp("2020-01-01"), pd.Timestamp(end))
style_macro_chart(
    ax,
    title="Pandemic excess savings — accumulated, then drawn down and negative",
    ylabel="$ billions",
    recessions=[r for r in recessions if r[1] >= pd.Timestamp("2020-01-01")],
    hlines=[{"y": 0.0}],
)
ax.annotate(f"Peak +${es_peak['EXCESS_SAVINGS']:,.0f}B\n"
            f"({es_peak['Date']:%b-%Y})",
            xy=(es_peak["Date"], es_peak["EXCESS_SAVINGS"]),
            xytext=(20, -10), textcoords="offset points", fontsize=8,
            arrowprops={"arrowstyle": "->", "color": "0.4"})
if es_cross is not None:
    ax.annotate(f"Crosses zero {es_cross['Date']:%b-%Y}",
                xy=(es_cross["Date"], 0), xytext=(15, 25),
                textcoords="offset points", fontsize=8,
                arrowprops={"arrowstyle": "->", "color": "0.4"})
fig.tight_layout()
fig.savefig(OUT_DIR / "excess_savings.png", dpi=150)
plt.close(fig)

# Chart 5 — buy-in-advance intent, 1978+ full window (documented exception)
# + 2006+ companion. Skipped entirely if no UMich data.
if df_intent is not None:
    intent_labels = {
        "INTENT_DURABLES": "Durables (table 36)",
        "INTENT_VEHICLES": "Vehicles (table 38)",
        "INTENT_HOMES": "Houses (table 42)",
    }
    intent_colors = {
        "INTENT_DURABLES": "#1f3b73",
        "INTENT_VEHICLES": "#2ca02c",
        "INTENT_HOMES": "#9ec5e8",
    }
    full_start = df_intent["Date"].iloc[0]
    recessions_full = get_recession_periods(fred, full_start, end)
    for suffix, x0, recs in (
        ("", full_start, recessions_full),
        ("_2006", pd.Timestamp(chart_start), recessions),
    ):
        # Subset to the charted window so the y-axis scales to it (the 1979-80
        # peaks would otherwise waste half the 2006+ panel).
        win = df_intent[df_intent["Date"] >= x0]
        fig, ax = plt.subplots(figsize=(11, 5))
        for col, label in intent_labels.items():
            if col not in win.columns:
                continue
            sub = win.dropna(subset=[col])
            ax.plot(sub["Date"], sub[col], color=intent_colors[col],
                    linewidth=1.2, label=f"{label} — good time, prices will rise")
        ax.axvline(UMICH_BREAK, color="0.35", linestyle=":", linewidth=1.0,
                   label="Survey phone→web transition (2024)")
        ax.set_xlim(x0, pd.Timestamp(end))
        style_macro_chart(
            ax,
            title="Buy-in-advance-of-rising-prices share (UMich SCA), "
                  + ("1978–present" if not suffix else "2006–present"),
            ylabel="% of respondents (multiple mentions allowed)",
            recessions=recs,
        )
        fig.tight_layout()
        fig.savefig(OUT_DIR / f"pullforward_intent{suffix}.png", dpi=150)
        plt.close(fig)

# Chart 6 — composite: two sub-indexes + overall gauge, with event vlines.
fig, ax = plt.subplots(figsize=(11, 5))
ax.plot(dfm06["Date"], dfm06["PULLFORWARD"], color="#9ec5e8", linewidth=1.3,
        label="PULL-FORWARD sub-index (higher = demand borrowed from future)")
ax.plot(dfm06["Date"], dfm06["STRAIN"], color="#c0392b", linewidth=1.3,
        label="STRAIN sub-index (higher = financing stretched)")
ax.plot(dfm06["Date"], dfm06["GAUGE"], color="#1f3b73", linewidth=2.0,
        label="Overall gauge (mean of the two)")
ax.set_xlim(*XLIM)
style_macro_chart(
    ax,
    title="Consumer pull-forward & sustainability gauge, 2006–present "
          "(z-scores vs 2006+ history)",
    ylabel="Z-score (2006+ window)",
    recessions=recessions,
    hlines=[{"y": 0.0}],
)
add_event_lines(ax, XLIM)
fig.tight_layout()
fig.savefig(OUT_DIR / "pullforward_composite.png", dpi=150)
plt.close(fig)

# ---------------------------------------------------------------------------
# Print summary
# ---------------------------------------------------------------------------
# Peak search restricted to 2020+ — backcast pre-2011 gaps are artifacts.
post20 = dfm06[dfm06["Date"] >= pd.Timestamp("2020-01-01")]
gap_peak = post20.loc[post20["DUR_GAP"].idxmax()]
gap_peak_lin = post20.loc[post20["DUR_GAP_LIN"].idxmax()]
alt_mar25 = dfm06.loc[dfm06["Date"] == pd.Timestamp("2025-03-01"), "ALTSALES"]
latest = dfm06.dropna(subset=["GAUGE"]).iloc[-1]
print(f"Chart start:              {chart_start.date()}")
print(f"End date:                 {end.date()}")
print(f"Monthly rows (2006+):     {len(dfm06)}")
print(f"Quarterly rows (2006+):   {len(df_q)}")
if df_intent is not None:
    print(f"Intent rows (1978+):      {len(df_intent)}")
print(f"Trend fit (frozen):       {FIT_START.date()}..{FIT_END.date()}, "
      f"growth {slope * 12 * 100:.1f}%/yr")
print(f"Durables gap peak 2020+:  log-linear {gap_peak['DUR_GAP']:+.1f}% "
      f"({gap_peak['Date']:%b-%Y}) | linear {gap_peak_lin['DUR_GAP_LIN']:+.1f}% "
      f"({gap_peak_lin['Date']:%b-%Y})")
print(f"Durables gap latest:      log-linear "
      f"{dfm06['DUR_GAP'].dropna().iloc[-1]:+.1f}% | linear "
      f"{dfm06['DUR_GAP_LIN'].dropna().iloc[-1]:+.1f}% "
      f"(rolling-10y z {dfm06['GAP_Z_ROLL10Y'].dropna().iloc[-1]:+.1f})")
print(f"Cumulative gap latest:    log-linear "
      f"{dfm06['DUR_GAP_CUM'].dropna().iloc[-1]:+,.0f} %-months | linear "
      f"{dfm06['DUR_GAP_CUM_LIN'].dropna().iloc[-1]:+,.0f} %-months")
if len(alt_mar25):
    print(f"ALTSALES Mar-2025:        {float(alt_mar25.iloc[0]):.1f}M SAAR")
payback = dfm06[(dfm06["Date"] >= pd.Timestamp("2025-04-01"))
                & (dfm06["Date"] <= pd.Timestamp("2025-12-01"))]
if len(payback.dropna(subset=["ALTSALES"])):
    trough = payback.loc[payback["ALTSALES"].idxmin()]
    print(f"ALTSALES 2025 payback low: {trough['ALTSALES']:.1f}M "
          f"({trough['Date']:%b-%Y})")
print(f"ALTSALES latest:          {dfm06['ALTSALES'].dropna().iloc[-1]:.1f}M "
      f"(gap vs replacement {dfm06['AUTO_GAP'].dropna().iloc[-1]:+.1f}M)")
print(f"Stock-adj ratio latest:   {df_q['STOCKADJ_RATIO'].dropna().iloc[-1]:.2f} "
      f"(2006+ mean {ratio_mean06:.2f})")
print(f"PSAVERT latest:           {dfm06['PSAVERT'].dropna().iloc[-1]:.1f}% "
      f"(2011-19 mean {psavert_mean_1119:.1f}%)")
print(f"Card delinquency latest:  {dfm06['DRCCLACBS'].dropna().iloc[-1]:.2f}%")
print(f"CI_CC latest:             {dfm06['CI_CC'].dropna().iloc[-1]:+.2f}% of DPI")
print(f"Excess savings peak:      +${es_peak['EXCESS_SAVINGS']:,.0f}B "
      f"({es_peak['Date']:%b-%Y})")
if es_cross is not None:
    print(f"Excess savings <0 since:  {es_cross['Date']:%b-%Y} "
          f"(latest {es['EXCESS_SAVINGS'].iloc[-1]:+,.0f}B)")
if df_intent is not None and "INTENT_DURABLES" in df_intent.columns:
    dur_int = df_intent.dropna(subset=["INTENT_DURABLES"])
    spike = dur_int[(dur_int["Date"] >= pd.Timestamp("2024-07-01"))
                    & (dur_int["Date"] <= pd.Timestamp("2025-12-01"))]
    spike_row = spike.loc[spike["INTENT_DURABLES"].idxmax()]
    print(f"Intent (durables) latest: "
          f"{dur_int['INTENT_DURABLES'].iloc[-1]:.0f} "
          f"({dur_int['Date'].iloc[-1]:%b-%Y}; 1979-80 peak "
          f"{dur_int['INTENT_DURABLES'].max():.0f}; tariff spike "
          f"{spike_row['INTENT_DURABLES']:.0f} {spike_row['Date']:%b-%Y})")
print(f"Gauge latest ({latest['Date']:%b-%Y}):   "
      f"pull-forward {latest['PULLFORWARD']:+.2f} | strain {latest['STRAIN']:+.2f} "
      f"| overall {latest['GAUGE']:+.2f}")
print(f"Wrote {xlsx_path.name}, durables_trend_gap.png, auto_replacement_gap.png, "
      f"stock_adjustment_ratio.png, excess_savings.png, pullforward_intent.png, "
      f"pullforward_intent_2006.png, pullforward_composite.png to {OUT_DIR}")
