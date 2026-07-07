# Independent audit of macro/consumer_pullforward_pull.py and
# macro/output/consumer_pullforward/.
#
# Re-pulls the raw FRED series itself (DDURRA3M086SBEA, PSAVERT, DSPI,
# ALTSALES, BOGZ1FA155111005Q, BOGZ1LM155111005Q, K1CTOTL1CD000,
# M1CTOTL1CD000, PCEDG, DPI) and re-derives the frozen-trend durables gaps,
# the SF-Fed excess-savings track, and the stock-adjustment ratio with its
# own code; re-parses the CACHED UMich SCA csv files independently. Does NOT
# import or exec consumer_pullforward_pull.py; reads
# consumer_pullforward.xlsx only to compare (plus a static text scan of the
# script source to confirm event vlines are chart-only).
#
# Tolerances:
#   - Gap re-derivations vs xlsx: 0.3pp (audit runs against a same-day FRED
#     vintage; real formula errors — wrong fit window, log/linear mix-up,
#     off-by-one t — are >1pp).
#   - Excess savings vs xlsx: $5B (same-vintage float round-trip only).
#     vs the published SF-Fed-style anchors (+$2,184B peak Aug-2021): $30B
#     PASS, $30-100B WARN (data-vintage drift), >$100B FAIL.
#   - Stock-adjustment ratio vs xlsx: 0.02. The discriminating test: if the
#     xlsx ratio matched net/replacement (~0.2) instead of
#     1 + net/replacement (~1.2), the script mishandled the NET flow -> FAIL.
#   - Composite recomputation from stored z columns: 0.01 (pure arithmetic).
#
# Prints one line per check: PASS|FAIL|WARN <name>: <detail>. Exits 1 on FAIL.

from datetime import datetime
from pathlib import Path
import re
import sys

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent      # macro/audits
MACRO_DIR = HERE.parent                     # macro/
REPO_ROOT = MACRO_DIR.parent
sys.path.insert(0, str(REPO_ROOT))

from shared.fred_helpers import get_fred_client  # noqa: E402

OUT_DIR = MACRO_DIR / "output" / "consumer_pullforward"
XLSX = OUT_DIR / "consumer_pullforward.xlsx"
SCRIPT = MACRO_DIR / "consumer_pullforward_pull.py"

TODAY = pd.Timestamp(datetime.today().date())

RESULTS = []


def check(status: str, name: str, detail: str) -> None:
    RESULTS.append(status)
    print(f"{status} {name}: {detail}")


# ---------------------------------------------------------------------------
# 1. Output file existence + non-trivial size
# ---------------------------------------------------------------------------
EXPECTED_FILES = {
    "consumer_pullforward.xlsx": 20_000,
    "durables_trend_gap.png": 30_000,
    "auto_replacement_gap.png": 30_000,
    "stock_adjustment_ratio.png": 30_000,
    "excess_savings.png": 30_000,
    "pullforward_intent.png": 30_000,
    "pullforward_intent_2006.png": 30_000,
    "pullforward_composite.png": 30_000,
    "umich_table36.csv": 5_000,
    "umich_table38.csv": 5_000,
    "umich_table42.csv": 5_000,
}
for fname, min_size in EXPECTED_FILES.items():
    p = OUT_DIR / fname
    if not p.exists():
        check("FAIL", f"file {fname}", "missing")
    elif p.stat().st_size < min_size:
        check("FAIL", f"file {fname}",
              f"only {p.stat().st_size} bytes (< {min_size})")
    else:
        check("PASS", f"file {fname}", f"{p.stat().st_size:,} bytes")

if not XLSX.exists():
    print("\nSummary: cannot continue without consumer_pullforward.xlsx")
    sys.exit(1)

dm = pd.read_excel(XLSX, sheet_name="Data_Monthly")
dm["Date"] = pd.to_datetime(dm["Date"])
dq = pd.read_excel(XLSX, sheet_name="Data_Quarterly")
dq["Date"] = pd.to_datetime(dq["Date"])
di = pd.read_excel(XLSX, sheet_name="Data_Intent")
di["Date"] = pd.to_datetime(di["Date"])
dmx = dm.set_index("Date")

# ---------------------------------------------------------------------------
# 2. Independent raw pulls from FRED
# ---------------------------------------------------------------------------
fred = get_fred_client()
START = "1990-01-01"
dur = fred.get_series("DDURRA3M086SBEA", observation_start=START).dropna()
psavert = fred.get_series("PSAVERT", observation_start=START).dropna()
dspi = fred.get_series("DSPI", observation_start=START).dropna()
altsales = fred.get_series("ALTSALES", observation_start=START).dropna()
net_flow = fred.get_series("BOGZ1FA155111005Q", observation_start=START).dropna()
dur_stock = fred.get_series("BOGZ1LM155111005Q", observation_start=START).dropna()
fa_stock = fred.get_series("K1CTOTL1CD000", observation_start=START).dropna()
fa_dep = fred.get_series("M1CTOTL1CD000", observation_start=START).dropna()
pcedg = fred.get_series("PCEDG", observation_start=START).dropna()
dpi_q = fred.get_series("DPI", observation_start=START).dropna()

# ---------------------------------------------------------------------------
# 3. Durables trend gaps — refit frozen 2011-01..2019-12 OLS myself
# ---------------------------------------------------------------------------
# Own construction: fit on the raw monthly series directly (DatetimeIndex,
# not the analysis's positional dataframe), t in months since Jan-2011.
fitwin = dur[(dur.index >= "2011-01-01") & (dur.index <= "2019-12-01")]
t_fit = (fitwin.index.year - 2011) * 12 + (fitwin.index.month - 1)
slope, intercept = np.polyfit(t_fit, np.log(fitwin.values), 1)
slope_yr_pct = slope * 12 * 100
if 5.7 <= slope_yr_pct <= 6.7:
    check("PASS", "frozen log-trend slope",
          f"refit 2011-01..2019-12 slope {slope_yr_pct:.2f}%/yr "
          f"(expected ~6.2%/yr, n={len(fitwin)} months)")
else:
    check("FAIL", "frozen log-trend slope",
          f"refit slope {slope_yr_pct:.2f}%/yr — expected ~6.2%/yr; "
          "wrong fit window or series")

t_all = (dur.index.year - 2011) * 12 + (dur.index.month - 1)
trend_log = np.exp(intercept + slope * t_all)
my_gap = pd.Series((np.log(dur.values) - np.log(trend_log)) * 100,
                   index=dur.index)
slope_lin, int_lin = np.polyfit(t_fit, fitwin.values, 1)
trend_lin = int_lin + slope_lin * t_all
my_gap_lin = pd.Series((dur.values / trend_lin - 1) * 100, index=dur.index)

GAP_TOL = 0.3  # pp vs xlsx
spot = pd.Timestamp("2021-03-01")
latest_m = my_gap.index[-1]
for name, mine_s, col, anchors in (
    ("log-linear gap", my_gap, "DUR_GAP",
     {spot: 17.3, latest_m: -11.8}),
    ("linear gap", my_gap_lin, "DUR_GAP_LIN",
     {latest_m: 3.3}),
):
    bad, msgs = [], []
    for d, anchor in anchors.items():
        mv = mine_s.loc[d]
        if d not in dmx.index or pd.isna(dmx.loc[d, col]):
            bad.append(f"{d.date()} missing from xlsx {col}")
            continue
        xv = dmx.loc[d, col]
        if abs(xv - mv) > GAP_TOL:
            bad.append(f"{d.date()}: mine {mv:+.2f}% vs xlsx {xv:+.2f}% "
                       f"(diff {abs(xv - mv):.2f}pp > {GAP_TOL}pp)")
        else:
            msgs.append(f"{d:%b-%Y} mine {mv:+.1f}% = xlsx {xv:+.1f}% "
                        f"(anchor ~{anchor:+.1f}%)")
    if bad:
        check("FAIL", f"durables {name}", "; ".join(bad))
    else:
        check("PASS", f"durables {name}", "; ".join(msgs))
        # anchor drift is vintage-dependent -> WARN only
        drift = [f"{d:%b-%Y} {mine_s.loc[d]:+.1f}% vs anchor {a:+.1f}%"
                 for d, a in anchors.items()
                 if abs(mine_s.loc[d] - a) > 1.0]
        if drift:
            check("WARN", f"durables {name} anchor drift",
                  "; ".join(drift) + " (>1pp from stated anchor — data "
                  "revision, verify narrative still holds)")

# linear-gap peak in the 2020+ window: anchor ~+23.9% at Apr-2021
post20_lin = my_gap_lin[my_gap_lin.index >= "2020-01-01"]
pk_d, pk_v = post20_lin.idxmax(), post20_lin.max()
if abs(pk_v - 23.9) <= 1.0 and pd.Timestamp("2021-01-01") <= pk_d <= \
        pd.Timestamp("2021-12-01"):
    check("PASS", "linear gap 2020+ peak",
          f"{pk_v:+.1f}% at {pk_d:%b-%Y} (anchor ~+23.9% Apr-2021)")
elif abs(pk_v - 23.9) <= 3.0:
    check("WARN", "linear gap 2020+ peak",
          f"{pk_v:+.1f}% at {pk_d:%b-%Y} vs anchor +23.9% Apr-2021 "
          "(vintage drift)")
else:
    check("FAIL", "linear gap 2020+ peak",
          f"{pk_v:+.1f}% at {pk_d:%b-%Y}, expected ~+23.9% Apr-2021")

# ---------------------------------------------------------------------------
# 4. Excess savings — SF Fed replication with my own code
# ---------------------------------------------------------------------------
saving = (psavert * dspi / 100).dropna()  # $bn SAAR
es_fit = saving[(saving.index >= "2016-03-01") & (saving.index <= "2020-02-01")]
es_t_fit = (es_fit.index.year - 2016) * 12 + (es_fit.index.month - 3)
b, a = np.polyfit(es_t_fit, es_fit.values, 1)
t_s = (saving.index.year - 2016) * 12 + (saving.index.month - 3)
trend_sav = a + b * t_s
excess = ((saving.values - trend_sav) / 12)
my_es = pd.Series(excess, index=saving.index)
my_es = my_es[my_es.index >= "2020-03-01"].cumsum()

my_pk_d, my_pk_v = my_es.idxmax(), my_es.max()
below = my_es[my_es < 0]
my_cross = below.index[0] if len(below) else None
my_latest = my_es.iloc[-1]

xl_es = dmx["EXCESS_SAVINGS"].dropna()
xl_pk_d, xl_pk_v = xl_es.idxmax(), xl_es.max()
ES_XL_TOL = 5.0
diffs = []
if abs(xl_pk_v - my_pk_v) > ES_XL_TOL or xl_pk_d != my_pk_d:
    diffs.append(f"peak mine +${my_pk_v:,.0f}B {my_pk_d:%b-%Y} vs xlsx "
                 f"+${xl_pk_v:,.0f}B {xl_pk_d:%b-%Y}")
if abs(xl_es.iloc[-1] - my_latest) > ES_XL_TOL:
    diffs.append(f"latest mine {my_latest:+,.0f}B vs xlsx "
                 f"{xl_es.iloc[-1]:+,.0f}B")
if diffs:
    check("FAIL", "excess savings vs xlsx", "; ".join(diffs))
else:
    check("PASS", "excess savings vs xlsx",
          f"peak +${my_pk_v:,.0f}B ({my_pk_d:%b-%Y}), latest "
          f"{my_latest:+,.0f}B — xlsx matches within ${ES_XL_TOL:.0f}B")

# anchors: peak ~ +$2,184B Aug-2021; zero-cross Dec-2024; latest ~ -$1,545B
pk_err = abs(my_pk_v - 2184)
cross_txt = f"{my_cross:%b-%Y}" if my_cross is not None else "none"
if pk_err <= 30 and my_pk_d == pd.Timestamp("2021-08-01"):
    check("PASS", "excess savings anchors",
          f"peak +${my_pk_v:,.0f}B {my_pk_d:%b-%Y} within $30B of +$2,184B "
          f"Aug-2021; zero-cross {cross_txt}; "
          f"latest {my_latest:+,.0f}B (anchor -$1,545B)")
elif pk_err <= 100:
    check("WARN", "excess savings anchors",
          f"peak +${my_pk_v:,.0f}B ({my_pk_d:%b-%Y}) is ${pk_err:.0f}B from "
          "the +$2,184B anchor — within $100B, consistent with DSPI/PSAVERT "
          "vintage drift")
else:
    check("FAIL", "excess savings anchors",
          f"peak +${my_pk_v:,.0f}B ({my_pk_d:%b-%Y}) is ${pk_err:.0f}B from "
          "the +$2,184B Aug-2021 anchor (>$100B) — construction error "
          "(check SAAR /12, trend window Mar-2016..Feb-2020)")

if my_cross is not None and my_cross == pd.Timestamp("2024-12-01"):
    check("PASS", "excess savings zero-cross",
          f"first negative month {my_cross:%b-%Y} (anchor Dec-2024)")
elif my_cross is not None and \
        abs((my_cross - pd.Timestamp("2024-12-01")).days) <= 95:
    check("WARN", "excess savings zero-cross",
          f"first negative month {my_cross:%b-%Y} vs anchor Dec-2024 "
          "(within a quarter — vintage drift)")
else:
    check("FAIL", "excess savings zero-cross",
          f"first negative month {my_cross} vs anchor Dec-2024")

# ---------------------------------------------------------------------------
# 5. Stock-adjustment ratio — NET-flow handling is the critical test
# ---------------------------------------------------------------------------
# 5a. Identity: BOGZ1FA155111005Q must be the NET transactions flow. Over
# 1995Q1-1999Q4, cumulated flows (SAAR/4) should be of the same order as the
# Z.1 stock change and FAR below cumulated gross PCEDG.
win = (net_flow.index >= "1995-01-01") & (net_flow.index <= "1999-10-01")
cum_net = (net_flow[win] / 4).sum() / 1000            # $M SAAR -> $bn
stock_chg = (dur_stock.loc["1999-10-01"] - dur_stock.loc["1994-10-01"]) / 1000
pgw = (pcedg.index >= "1995-01-01") & (pcedg.index <= "1999-12-01")
cum_gross = (pcedg[pgw] / 12).sum()                   # $bn SAAR monthly -> $bn
ratio_ns = cum_net / stock_chg
if 0.5 <= ratio_ns <= 2.5 and cum_net < 0.4 * cum_gross:
    check("PASS", "Z.1 flow is NET (identity 1995-99)",
          f"cum flows ${cum_net:,.0f}bn ~ stock change ${stock_chg:,.0f}bn "
          f"(ratio {ratio_ns:.2f}) << cum gross PCEDG ${cum_gross:,.0f}bn — "
          "confirms BOGZ1FA155111005Q is net-of-depreciation, as the script "
          "assumes")
elif cum_net > 0.7 * cum_gross:
    check("FAIL", "Z.1 flow is NET (identity 1995-99)",
          f"cum flows ${cum_net:,.0f}bn ~ gross PCEDG ${cum_gross:,.0f}bn — "
          "series is GROSS purchases; the script's net+replacement "
          "reconstruction would double-count replacement")
else:
    check("WARN", "Z.1 flow is NET (identity 1995-99)",
          f"cum flows ${cum_net:,.0f}bn vs stock change ${stock_chg:,.0f}bn "
          f"(ratio {ratio_ns:.2f}) vs gross ${cum_gross:,.0f}bn — identity "
          "looser than expected (revaluation effects?)")

# 5b. Re-derive the ratio with my own interpolation of the annual dep rate.
delta_a = (fa_dep / fa_stock).dropna()
q_idx = pd.DatetimeIndex(dur_stock.index)
delta_q = (delta_a.reindex(delta_a.index.union(q_idx))
           .interpolate(method="time").ffill().reindex(q_idx))
repl = delta_q * dur_stock.shift(1)
my_ratio = ((net_flow.reindex(q_idx) + repl) / repl).dropna()
my_r_latest = my_ratio.iloc[-1]
xl_ratio = dq.set_index("Date")["STOCKADJ_RATIO"].dropna()
xl_r_latest = xl_ratio.iloc[-1]
net_only = (net_flow.reindex(q_idx) / repl).dropna().iloc[-1]
if abs(xl_r_latest - net_only) < 0.05:
    check("FAIL", "stock-adjustment ratio construction",
          f"xlsx latest {xl_r_latest:.2f} matches net/replacement "
          f"({net_only:.2f}) — script treats the NET flow as gross purchases")
elif abs(xl_r_latest - my_r_latest) <= 0.02 and \
        xl_ratio.index[-1] == my_ratio.index[-1]:
    check("PASS", "stock-adjustment ratio construction",
          f"latest mine {my_r_latest:.3f} vs xlsx {xl_r_latest:.3f} at "
          f"{xl_ratio.index[-1]:%b-%Y} (anchor ~1.22; net/replacement would "
          f"be {net_only:.2f})")
else:
    check("FAIL", "stock-adjustment ratio construction",
          f"latest mine {my_r_latest:.3f} ({my_ratio.index[-1]:%b-%Y}) vs "
          f"xlsx {xl_r_latest:.3f} ({xl_ratio.index[-1]:%b-%Y}) — diff "
          "> 0.02")
if abs(my_r_latest - 1.22) > 0.05:
    check("WARN", "stock-adjustment ratio anchor",
          f"my latest {my_r_latest:.2f} vs stated anchor ~1.22 "
          "(vintage drift)")

# ---------------------------------------------------------------------------
# 6. UMich intent leg — independent re-parse of the CACHED csvs
# ---------------------------------------------------------------------------
def parse_umich(path: Path) -> pd.Series:
    lines = path.read_text(encoding="utf-8").splitlines()
    hdr = [c.strip() for c in re.sub("<br>", " ", lines[1], flags=re.I).split(",")]
    idx = next(i for i, c in enumerate(hdr)
               if "good time" in c.lower() and "prices will increase" in c.lower())
    rows = [ln.split(",") for ln in lines[2:] if ln.strip()]
    dates = pd.to_datetime(
        {"year": [r[1] for r in rows], "month": [r[0] for r in rows],
         "day": [1] * len(rows)}, errors="coerce")
    vals = pd.to_numeric([r[idx] for r in rows], errors="coerce")
    return pd.Series(vals.values if hasattr(vals, "values") else vals,
                     index=dates).dropna().sort_index()


t36 = parse_umich(OUT_DIR / "umich_table36.csv")
bench = {
    "Dec-2024 tariff-anticipation": (pd.Timestamp("2024-12-01"), 22),
    "1979-80 episode max": (None, 53),
    "latest (May-2026)": (pd.Timestamp("2026-05-01"), 9),
}
bad = []
v_dec = t36.get(pd.Timestamp("2024-12-01"))
v7980 = t36[(t36.index >= "1979-01-01") & (t36.index <= "1980-12-01")].max()
v_last, d_last = t36.iloc[-1], t36.index[-1]
if v_dec != 22:
    bad.append(f"Dec-2024 = {v_dec}, expected 22")
if v7980 != 53:
    bad.append(f"1979-80 max = {v7980}, expected 53")
if d_last != pd.Timestamp("2026-05-01") or v_last != 9:
    bad.append(f"latest = {v_last} at {d_last:%b-%Y}, expected 9 at May-2026")
if bad:
    check("FAIL", "UMich cached csv benchmarks", "; ".join(bad))
else:
    check("PASS", "UMich cached csv benchmarks",
          "durables advance-buying share: Dec-2024 = 22, 1979-80 max = 53, "
          "May-2026 = 9 (all match)")

# xlsx Data_Intent must equal the re-parse, and INTENT_AVG = row mean of the
# three tables.
t38 = parse_umich(OUT_DIR / "umich_table38.csv")
t42 = parse_umich(OUT_DIR / "umich_table42.csv")
dix = di.set_index("Date")
merged = pd.DataFrame({"D": t36, "V": t38, "H": t42})
my_avg = merged.mean(axis=1)
common = dix.index.intersection(my_avg.index)
d_dur = (dix.loc[common, "INTENT_DURABLES"] - t36.reindex(common)).abs().max()
d_avg = (dix.loc[common, "INTENT_AVG"] - my_avg.reindex(common)).abs().max()
if d_dur < 1e-9 and d_avg < 1e-9:
    check("PASS", "Data_Intent vs cached csvs",
          f"{len(common)} months: INTENT_DURABLES exact, INTENT_AVG = mean "
          f"of 3 tables (max diff {d_avg:.1e}); latest INTENT_AVG "
          f"{dix['INTENT_AVG'].dropna().iloc[-1]:.2f}")
else:
    check("FAIL", "Data_Intent vs cached csvs",
          f"max diffs: INTENT_DURABLES {d_dur:.3g}, INTENT_AVG {d_avg:.3g}")

# ---------------------------------------------------------------------------
# 7. Composite arithmetic at the latest month + z-window verification
# ---------------------------------------------------------------------------
last = dm.dropna(subset=["GAUGE"]).iloc[-1]
pull_cols = [c for c in ["Z_DUR_GAP", "Z_AUTO_GAP", "Z_STOCKADJ",
                         "Z_FRONTRUN", "Z_INTENT"] if c in dm.columns]
strain_cols = ["Z_SAVERT", "Z_CREDIT_GAP", "Z_DELINQ", "Z_CICC"]
my_pull = np.nanmean([last[c] for c in pull_cols])
my_strain = np.nanmean([last[c] for c in strain_cols])
my_gauge = (my_pull + my_strain) / 2
errs = []
if abs(my_pull - last["PULLFORWARD"]) > 0.01:
    errs.append(f"pull-forward recomputed {my_pull:+.3f} vs stored "
                f"{last['PULLFORWARD']:+.3f}")
if abs(my_strain - last["STRAIN"]) > 0.01:
    errs.append(f"strain recomputed {my_strain:+.3f} vs stored "
                f"{last['STRAIN']:+.3f}")
if abs(my_gauge - last["GAUGE"]) > 0.01:
    errs.append(f"gauge recomputed {my_gauge:+.3f} vs stored "
                f"{last['GAUGE']:+.3f}")
if errs:
    check("FAIL", "composite arithmetic", "; ".join(errs))
else:
    check("PASS", "composite arithmetic",
          f"at {last['Date']:%b-%Y}: pull-forward {my_pull:+.2f} "
          f"({len(pull_cols)} legs), strain {my_strain:+.2f}, overall "
          f"{my_gauge:+.2f} — all equal stored values")
anchor_dev = max(abs(my_pull - (-1.39)), abs(my_strain - 0.02),
                 abs(my_gauge - (-0.68)))
if anchor_dev > 0.05:
    check("WARN", "composite anchors",
          f"latest pull {my_pull:+.2f}/strain {my_strain:+.2f}/gauge "
          f"{my_gauge:+.2f} vs anchors -1.39/+0.02/-0.68 (max dev "
          f"{anchor_dev:.2f} — vintage drift)")

# z-scores must be standardized over the 2006+ charted window (repo
# convention): recompute Z_DUR_GAP from the stored 2006+ DUR_GAP column.
g06 = dm["DUR_GAP"]
z_mine = (g06 - g06.mean()) / g06.std()
zdiff = (z_mine - dm["Z_DUR_GAP"]).abs().max()
if zdiff < 1e-9:
    check("PASS", "z-score window",
          f"Z_DUR_GAP == z over the 2006+ sheet window (mean "
          f"{g06.mean():+.2f}, sd {g06.std():.2f}, max diff {zdiff:.1e})")
else:
    # would it instead match a z over the full 1990+ pull?
    gfull = my_gap
    zfull = ((g06 - gfull.mean()) / gfull.std())
    alt = (zfull - dm["Z_DUR_GAP"]).abs().max()
    check("FAIL", "z-score window",
          f"Z_DUR_GAP does not match a 2006+ z (max diff {zdiff:.3g}); "
          f"full-history z max diff {alt:.3g}")

# ---------------------------------------------------------------------------
# 8. Internal consistency: date axis, interior NaNs, staleness
# ---------------------------------------------------------------------------
dts = dm["Date"]
mono = dts.is_monotonic_increasing
dupes = int(dts.duplicated().sum())
per = pd.PeriodIndex(dts, freq="M")
missing = len(pd.period_range(per[0], per[-1], freq="M")) - len(per)
starts_2006 = dts.iloc[0] == pd.Timestamp("2006-01-01")
if mono and dupes == 0 and missing == 0 and starts_2006:
    check("PASS", "Data_Monthly date axis",
          f"{len(dts)} rows, contiguous monthly "
          f"{dts.iloc[0].date()}..{dts.iloc[-1].date()}, no dupes/gaps, "
          "starts 2006-01 per convention")
else:
    check("FAIL", "Data_Monthly date axis",
          f"monotonic={mono}, dupes={dupes}, missing months={missing}, "
          f"starts {dts.iloc[0].date()}")

holes = []
for col in ["DUR_REAL", "DUR_GAP", "ALTSALES", "PSAVERT", "DSPI",
            "STOCKADJ_RATIO", "PULLFORWARD", "STRAIN", "GAUGE"]:
    s = dmx[col]
    valid = s.dropna()
    if valid.empty:
        holes.append(f"{col}: all NaN")
        continue
    inner = s.loc[valid.index[0]:valid.index[-1]]
    if inner.isna().sum():
        holes.append(f"{col}: {int(inner.isna().sum())} interior NaN(s)")
if holes:
    check("FAIL", "interior NaNs", "; ".join(holes))
else:
    check("PASS", "interior NaNs",
          "no interior NaN runs in key monthly columns (trailing "
          "publication-lag NaNs allowed)")

age = (TODAY - dmx["DUR_GAP"].dropna().index[-1]).days
if age <= 75:
    check("PASS", "staleness",
          f"latest durables month {dmx['DUR_GAP'].dropna().index[-1].date()} "
          f"is {age} days old (<= 75; PCE publication lag ~1 month)")
else:
    check("FAIL", "staleness",
          f"latest durables month is {age} days old (> 75) — output stale")

# unit sanity: no series jumps by ~1000x anywhere
jumps = []
for col in ["DUR_REAL", "ALTSALES", "DSPI", "REVOLSL"]:
    s = dmx[col].dropna()
    if (s <= 0).any():
        continue
    lg = np.abs(np.log(s / s.shift(1))).dropna()
    if len(lg) and lg.max() >= 1.0:
        jumps.append(f"{col}: |m/m log change| {lg.max():.2f} at "
                     f"{lg.idxmax().date()}")
if jumps:
    check("FAIL", "unit discontinuities", "; ".join(jumps))
else:
    check("PASS", "unit discontinuities",
          "no ~1000x unit breaks in DUR_REAL/ALTSALES/DSPI/REVOLSL")

# ---------------------------------------------------------------------------
# 9. Event vlines are chart-only (static scan + data spot-check)
# ---------------------------------------------------------------------------
src = SCRIPT.read_text(encoding="utf-8")
ev_uses = [ln.strip() for ln in src.splitlines()
           if "EVENTS" in ln and not ln.strip().startswith("#")]
mutating = [ln for ln in ev_uses
            if re.search(r"(df_m|df_q|dfm06|df_intent)\s*\[", ln)]
raw_match = True
for ts, _ in [(pd.Timestamp("2009-07-01"), ""), (pd.Timestamp("2021-03-01"), ""),
              (pd.Timestamp("2025-03-01"), "")]:
    if ts in dmx.index and ts in dur.index:
        if abs(dmx.loc[ts, "DUR_REAL"] - dur.loc[ts]) > 1e-6 * dur.loc[ts]:
            raw_match = False
if not mutating and raw_match:
    check("PASS", "event vlines chart-only",
          f"EVENTS referenced on {len(ev_uses)} non-comment lines, none "
          "index a dataframe; DUR_REAL in xlsx equals the raw FRED pull at "
          "event months — no data mutation")
else:
    check("FAIL", "event vlines chart-only",
          f"mutating lines: {mutating}; raw DUR_REAL match at event months: "
          f"{raw_match}")

# ---------------------------------------------------------------------------
# 10. EXTERNAL cross-check: monthly DSPI aggregates to quarterly DPI (two
# independent FRED series), validating the excess-savings income leg
# ---------------------------------------------------------------------------
dspi_q = dspi.resample("QS").mean()
common_q = dspi_q.index.intersection(dpi_q.index)[-8:]
rel = ((dspi_q.loc[common_q] - dpi_q.loc[common_q]).abs()
       / dpi_q.loc[common_q]).max()
if rel <= 0.005:
    check("PASS", "external DSPI vs quarterly DPI",
          f"quarterly mean of monthly DSPI matches the independent DPI "
          f"series within {rel:.2%} over the last {len(common_q)} quarters "
          f"(latest DPI ${dpi_q.iloc[-1]/1000:.1f}T SAAR) — income leg units "
          "confirmed")
else:
    check("FAIL", "external DSPI vs quarterly DPI",
          f"max relative gap {rel:.2%} over last {len(common_q)} quarters — "
          "DSPI units/vintage problem")

# ---------------------------------------------------------------------------
# 11. Documentation accuracy (static scan of the script header/comments)
# ---------------------------------------------------------------------------
# 11a. Header anchor for the excess-savings zero-cross vs actual output
# (the comment wraps across lines, so allow a "\n#" between the words).
m = re.search(r"crosses zero\s*(?:\n#\s*)?~?\s*([A-Za-z]{3})-(\d{4})", src)
if m and my_cross is not None:
    stated = pd.Timestamp(f"{m.group(1)} 1 {m.group(2)}")
    if stated == my_cross:
        check("PASS", "header zero-cross anchor",
              f"script header states {stated:%b-%Y}, matches actual "
              f"{my_cross:%b-%Y}")
    else:
        check("WARN", "header zero-cross anchor",
              f"script header (line ~38) states excess savings 'crosses "
              f"zero ~ {stated:%b-%Y}' but the actual output (and my "
              f"re-derivation) crosses {my_cross:%b-%Y} — stale comment, "
              "numbers themselves are correct")
else:
    check("WARN", "header zero-cross anchor",
          "could not locate/parse the zero-cross anchor in the script "
          "header for comparison")

# 11b. REVOLSL units: FRED serves REVOLSL in $MILLIONS (revolving consumer
# credit is ~$1.3T, so the level should be ~1.3e6). The script comments it
# as "$M" and only uses the YoY % — verify level magnitude matches the
# commented unit.
rev_latest = dmx["REVOLSL"].dropna().iloc[-1]
rev_comment_m = bool(re.search(r'"REVOLSL",\s*#[^\n]*\$M', src))
if 300_000 <= rev_latest <= 5_000_000 and rev_comment_m:
    check("PASS", "REVOLSL units comment",
          f"latest {rev_latest:,.0f} $M = ${rev_latest/1e6:.2f}T revolving "
          "credit (plausible), consistent with the script's '$M' comment; "
          "only YoY % enters the composite so units cancel anyway")
else:
    check("FAIL", "REVOLSL units comment",
          f"latest REVOLSL {rev_latest:,.0f} vs '$M' comment present="
          f"{rev_comment_m} — magnitude inconsistent with commented unit")

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
n_pass = RESULTS.count("PASS")
n_warn = RESULTS.count("WARN")
n_fail = RESULTS.count("FAIL")
print(f"\nSummary: {n_pass} PASS, {n_warn} WARN, {n_fail} FAIL "
      f"({len(RESULTS)} checks)")
sys.exit(1 if n_fail else 0)
