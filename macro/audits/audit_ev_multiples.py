# Independent audit of macro/ev_multiples_pull.py and macro/output/ev_multiples/.
#
# Re-pulls all raw FRED series itself (Z.1 stocks: NCBCEL, BCNSDODNS,
# BOGZ1FL104001005Q; NIPA 1.14 flows: A463RC1Q027SBEA, B471RC1Q027SBEA,
# B456RC1Q027SBEA, B465RC1Q027SBEA, W328RC1Q027SBEA, A455RC1Q027SBEA; Z.1
# flows: BOGZ1FA105050005Q, BOGZ1FA106121075Q, NCBCEBQ027S; plus BAA/AAA)
# and re-derives EV, EBIT/EBITDA multiples, FCF yields, the distribution leg
# and the near-zero mask with its own code. Does NOT import or exec
# ev_multiples_pull.py; reads ev_multiples.xlsx only to compare.
#
# Adversarial focus:
#   - SAAR handling: flows are SAAR, so trailing-year smoothing must be a
#     rolling MEAN. A rolling SUM would inflate the denominator 4x and put
#     EV/EBITDA at ~3.4x instead of ~13.8x — a discriminating test is run.
#   - Window length: the FCF legs are documented as 8q means; an alternative
#     4q construction is computed to prove the xlsx used 8q, not 4q.
#   - Units: Z.1 is $MM, NIPA is $B SAAR; a mixed-unit EV would be off 1000x.
#   - Sign convention: LFCF_DIST = dividends MINUS net issuance (issuance
#     negative when buybacks dominate); a flipped sign would go negative.
#   - Mask rule: EV_UFCF / MC_LFCF must be NaN exactly where the smoothed
#     numerator flow < 1% of EV / MktCap (or is warm-up NaN), nowhere else.
#
# Tolerances:
#   - Value re-derivations: 0.1% relative (same-day FRED vintage; only float
#     round-trip through xlsx expected; real defects are >=4x or 1000x).
#   - External benchmarks (research probe): current EV/EBITDA percentile
#     ~98th (accept 95-99.5), full-history median ~6.8x (accept +/-0.3),
#     latest Baa-Aaa ~0.48-0.54pp, MC/dist latest ~44.5x (accept +/-1%).
#   - Staleness: latest ratio quarter within 200 days of today (Z.1 lags the
#     quarter end by ~10 weeks).
#
# Prints one line per check: PASS|FAIL|WARN <name>: <detail>. Exits 1 on FAIL.

from datetime import datetime
from pathlib import Path
import sys

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent      # macro/audits
MACRO_DIR = HERE.parent                     # macro/
REPO_ROOT = MACRO_DIR.parent
sys.path.insert(0, str(REPO_ROOT))

from shared.fred_helpers import get_fred_client  # noqa: E402

OUT_DIR = MACRO_DIR / "output" / "ev_multiples"
XLSX = OUT_DIR / "ev_multiples.xlsx"

TODAY = pd.Timestamp(datetime.today().date())
REL_TOL = 0.001

RESULTS = []


def check(status: str, name: str, detail: str) -> None:
    RESULTS.append(status)
    print(f"{status} {name}: {detail}")


def rel(a: float, b: float) -> float:
    return abs(a - b) / abs(b)


# ---------------------------------------------------------------------------
# 1. Output file existence + non-trivial size
# ---------------------------------------------------------------------------
EXPECTED_FILES = {
    "ev_multiples.xlsx": 20_000,
    "ev_ebit_ebitda.png": 20_000,
    "fcf_yields.png": 20_000,
    "marketcap_gva.png": 20_000,
    "ev_multiples_2006.png": 20_000,
    "valuation_vs_credit.png": 20_000,
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
    print("\nSummary: cannot continue without ev_multiples.xlsx")
    sys.exit(1)

data = pd.read_excel(XLSX, sheet_name="Data")
data["Date"] = pd.to_datetime(data["Date"])
summary = pd.read_excel(XLSX, sheet_name="Summary", index_col=0)
xl = data.set_index("Date")

RATIO_COLS = [
    "EV_EBIT", "EV_EBITDA", "EV_UFCF", "MC_LFCF", "MC_LFCF_DIST",
    "UFCF_YIELD", "LFCF_YIELD", "LFCF_DIST_YIELD",
    "MC_GVA", "EBITDA_MARGIN_GVA", "BAA_AAA_SPREAD",
]

# ---------------------------------------------------------------------------
# 2. Independent raw pulls from FRED (own code; buffered from 1945 so the
#    rolling windows at 1952Q4 have the same lookback the analysis had)
# ---------------------------------------------------------------------------
fred = get_fred_client()
RAW_IDS = {
    "EQ_MM": "NCBCEL",
    "DEBT_MM": "BCNSDODNS",
    "LIQ_MM": "BOGZ1FL104001005Q",
    "PROFIT_PRETAX": "A463RC1Q027SBEA",
    "NET_INTEREST": "B471RC1Q027SBEA",
    "CFC": "B456RC1Q027SBEA",
    "TAXES": "B465RC1Q027SBEA",
    "PROFIT_AT": "W328RC1Q027SBEA",
    "GVA": "A455RC1Q027SBEA",
    "CAPEX_MM": "BOGZ1FA105050005Q",
    "DIV_MM": "BOGZ1FA106121075Q",
    "EQ_ISS_MM": "NCBCEBQ027S",
}
raw = {name: fred.get_series(sid, observation_start="1945-01-01")
       for name, sid in RAW_IDS.items()}

mine = pd.DataFrame({
    "MKTCAP": raw["EQ_MM"] / 1000.0,      # $MM -> $B
    "DEBT": raw["DEBT_MM"] / 1000.0,
    "LIQ": raw["LIQ_MM"] / 1000.0,
    "CAPEX": raw["CAPEX_MM"] / 1000.0,
    "DIV": raw["DIV_MM"] / 1000.0,
    "EQ_ISS": raw["EQ_ISS_MM"] / 1000.0,
    "PROFIT_PRETAX": raw["PROFIT_PRETAX"],  # already $B SAAR
    "NET_INTEREST": raw["NET_INTEREST"],
    "CFC": raw["CFC"],
    "TAXES": raw["TAXES"],
    "PROFIT_AT": raw["PROFIT_AT"],
    "GVA": raw["GVA"],
})
mine = mine.sort_index()
mine["EV"] = mine["MKTCAP"] + mine["DEBT"] - mine["LIQ"]
mine["EBIT"] = mine["PROFIT_PRETAX"] + mine["NET_INTEREST"]
mine["EBITDA"] = mine["EBIT"] + mine["CFC"]
mine["UFCF"] = mine["EBIT"] - mine["TAXES"] + mine["CFC"] - mine["CAPEX"]
mine["LFCF"] = mine["PROFIT_AT"] + mine["CFC"] - mine["CAPEX"]
mine["LFCF_DIST"] = mine["DIV"] - mine["EQ_ISS"]

# SAAR flows: trailing-year smoothing = rolling MEAN (4q earnings/dist,
# 8q for the noisier FCF legs).
mine["EBIT_4Q"] = mine["EBIT"].rolling(4).mean()
mine["EBITDA_4Q"] = mine["EBITDA"].rolling(4).mean()
mine["LFCF_DIST_4Q"] = mine["LFCF_DIST"].rolling(4).mean()
mine["UFCF_8Q"] = mine["UFCF"].rolling(8).mean()
mine["LFCF_8Q"] = mine["LFCF"].rolling(8).mean()

mine["EV_EBIT"] = mine["EV"] / mine["EBIT_4Q"]
mine["EV_EBITDA"] = mine["EV"] / mine["EBITDA_4Q"]
mine["MC_LFCF_DIST"] = mine["MKTCAP"] / mine["LFCF_DIST_4Q"]
mine["UFCF_YIELD"] = mine["UFCF_8Q"] / mine["EV"] * 100
mine["LFCF_YIELD"] = mine["LFCF_8Q"] / mine["MKTCAP"] * 100
mine["LFCF_DIST_YIELD"] = mine["LFCF_DIST_4Q"] / mine["MKTCAP"] * 100
mine["MC_GVA"] = mine["MKTCAP"] / mine["GVA"]
mine["EBITDA_MARGIN_GVA"] = mine["EBITDA"] / mine["GVA"] * 100

# Deliberately DIFFERENT (wrong) constructions, used only to prove the
# discriminating tests below can actually tell the difference:
mine["EV_EBITDA_SUMWIN"] = mine["EV"] / mine["EBITDA"].rolling(4).sum()
mine["UFCF_YIELD_4Q"] = mine["UFCF"].rolling(4).mean() / mine["EV"] * 100

# ---------------------------------------------------------------------------
# 3. EV identity at 2026Q1 (units test: Z.1 $MM -> $B; sign of liquid assets)
# ---------------------------------------------------------------------------
Q1_26 = pd.Timestamp("2026-01-01")
EXP_EV = {"EV": 75_459, "MKTCAP": 69_512, "DEBT": 14_454, "LIQ": 8_506}
if Q1_26 not in mine.index or pd.isna(mine.loc[Q1_26, "EV"]):
    check("WARN", "EV identity 2026Q1",
          "2026Q1 absent/NaN in independent pull (revision since analysis run?)")
elif Q1_26 not in xl.index:
    check("FAIL", "EV identity 2026Q1", "2026Q1 row missing from xlsx")
else:
    parts_bad = []
    for col in ["MKTCAP", "DEBT", "LIQ", "EV"]:
        mv, xv = mine.loc[Q1_26, col], xl.loc[Q1_26, col]
        if rel(xv, mv) > REL_TOL:
            parts_bad.append(f"{col}: xlsx {xv:,.0f}B vs re-pull {mv:,.0f}B")
        if rel(mv, EXP_EV[col]) > 0.005:   # 0.5% headroom for a re-vintage
            parts_bad.append(
                f"{col}: re-pull {mv:,.0f}B vs expected ~{EXP_EV[col]:,}B")
    ev_id = (xl.loc[Q1_26, "MKTCAP"] + xl.loc[Q1_26, "DEBT"]
             - xl.loc[Q1_26, "LIQ"])
    if rel(xl.loc[Q1_26, "EV"], ev_id) > 1e-9:
        parts_bad.append(
            f"xlsx EV {xl.loc[Q1_26, 'EV']:,.1f} != MKTCAP+DEBT-LIQ {ev_id:,.1f}")
    if parts_bad:
        check("FAIL", "EV identity 2026Q1", "; ".join(parts_bad))
    else:
        check("PASS", "EV identity 2026Q1",
              f"EV ${xl.loc[Q1_26, 'EV']:,.0f}B = "
              f"{xl.loc[Q1_26, 'MKTCAP']:,.0f} + {xl.loc[Q1_26, 'DEBT']:,.0f} "
              f"- {xl.loc[Q1_26, 'LIQ']:,.0f} (matches independent re-pull "
              f"and expected ~$75,459B; units $MM->$B confirmed)")

# ---------------------------------------------------------------------------
# 4. Spot re-derivations: EV/EBITDA(4q) and EV/EBIT(4q) at 4 quarters
# ---------------------------------------------------------------------------
latest_ratio = xl["EV_EBITDA"].dropna().index[-1]
SPOTS = [pd.Timestamp("1952-10-01"), pd.Timestamp("2000-01-01"),
         pd.Timestamp("2021-10-01"), latest_ratio]
for spot in SPOTS:
    label = f"spot re-derivation {spot.date()}"
    if spot not in xl.index:
        check("FAIL", label, "date missing from xlsx Data sheet")
        continue
    if spot not in mine.index:
        check("FAIL", label, "date missing from independent FRED pull")
        continue
    bad, worst = [], ("", 0.0)
    for col in ["EV_EBIT", "EV_EBITDA", "MC_GVA", "EBITDA_MARGIN_GVA"]:
        mv, xv = mine.loc[spot, col], xl.loc[spot, col]
        if pd.isna(mv) or pd.isna(xv):
            bad.append(f"{col}: mine={mv}, xlsx={xv} (NaN)")
            continue
        r = rel(xv, mv)
        if r > worst[1]:
            worst = (col, r)
        if r > REL_TOL:
            bad.append(f"{col}: mine={mv:.4f} xlsx={xv:.4f} rel={r:.2%}")
    if bad:
        check("FAIL", label, "; ".join(bad))
    else:
        check("PASS", label,
              f"EV/EBIT {xl.loc[spot, 'EV_EBIT']:.2f}x, "
              f"EV/EBITDA {xl.loc[spot, 'EV_EBITDA']:.2f}x re-derived within "
              f"{REL_TOL:.1%} (worst {worst[0]} rel {worst[1]:.2e})")

# ---------------------------------------------------------------------------
# 5. SAAR discriminating window test: 4q MEAN of SAAR EBITDA, not 4q SUM
# ---------------------------------------------------------------------------
if latest_ratio in mine.index and not pd.isna(mine.loc[latest_ratio, "EV_EBITDA_SUMWIN"]):
    xv = xl.loc[latest_ratio, "EV_EBITDA"]
    v_mean = mine.loc[latest_ratio, "EV_EBITDA"]
    v_sum = mine.loc[latest_ratio, "EV_EBITDA_SUMWIN"]
    if rel(v_mean, v_sum) < 0.5:
        check("WARN", "SAAR mean-vs-sum",
              "mean and sum constructions unexpectedly close; test not "
              "discriminating")
    elif rel(xv, v_mean) <= REL_TOL and rel(xv, v_sum) > 0.5:
        check("PASS", "SAAR mean-vs-sum",
              f"xlsx EV/EBITDA {xv:.2f}x = EV / 4q MEAN of SAAR EBITDA "
              f"({v_mean:.2f}x), not EV / 4q SUM ({v_sum:.2f}x ~= mean/4) — "
              "SAAR handled correctly")
    elif rel(xv, v_sum) <= REL_TOL:
        check("FAIL", "SAAR mean-vs-sum",
              f"xlsx EV/EBITDA {xv:.2f}x matches the 4q SUM construction "
              f"({v_sum:.2f}x): SAAR flows were summed, denominator 4x too "
              f"big (correct mean gives {v_mean:.2f}x)")
    else:
        check("FAIL", "SAAR mean-vs-sum",
              f"xlsx EV/EBITDA {xv:.2f}x matches neither mean ({v_mean:.2f}x) "
              f"nor sum ({v_sum:.2f}x) construction")
else:
    check("FAIL", "SAAR mean-vs-sum", f"{latest_ratio.date()} not derivable")

# ---------------------------------------------------------------------------
# 6. FCF window test: UFCF yield uses an 8q mean (4q would differ)
# ---------------------------------------------------------------------------
u_latest = xl["UFCF_YIELD"].dropna().index[-1]
if u_latest in mine.index and not pd.isna(mine.loc[u_latest, "UFCF_YIELD"]):
    xv = xl.loc[u_latest, "UFCF_YIELD"]
    v8 = mine.loc[u_latest, "UFCF_YIELD"]
    v4 = mine.loc[u_latest, "UFCF_YIELD_4Q"]
    d8, d4 = abs(xv - v8), abs(xv - v4)
    if abs(v8 - v4) < 0.01:
        check("WARN", "UFCF 8q window",
              f"8q ({v8:.3f}%) and 4q ({v4:.3f}%) yields differ <1bp at "
              f"{u_latest.date()}; window test not discriminating")
    elif rel(xv, v8) <= REL_TOL and d4 > 10 * d8 + 0.005:
        check("PASS", "UFCF 8q window",
              f"xlsx UFCF yield {xv:.2f}% at {u_latest.date()} = 8q mean "
              f"({v8:.2f}%, expected ~2.28%); a 4q mean would give {v4:.2f}% "
              "— 8q window confirmed")
    elif rel(xv, v4) <= REL_TOL:
        check("FAIL", "UFCF 8q window",
              f"xlsx UFCF yield {xv:.2f}% matches the 4q construction "
              f"({v4:.2f}%), not the documented 8q ({v8:.2f}%)")
    else:
        check("FAIL", "UFCF 8q window",
              f"xlsx UFCF yield {xv:.3f}% matches neither 8q ({v8:.3f}%) nor "
              f"4q ({v4:.3f}%) construction")
else:
    check("FAIL", "UFCF 8q window", f"{u_latest.date()} not derivable")

# ---------------------------------------------------------------------------
# 7. Distribution leg: DIV - net issuance positive throughout; MC/dist ~44.5x
# ---------------------------------------------------------------------------
dist_mine = mine.loc[mine.index >= "1952-10-01", "LFCF_DIST"].dropna()
dist_xl = xl["LFCF_DIST"].dropna()
n_neg_mine = (dist_mine <= 0).sum()
n_neg_xl = (dist_xl <= 0).sum()
if n_neg_mine == 0 and n_neg_xl == 0:
    check("PASS", "LFCF_DIST sign convention",
          f"dividends - net issuance > 0 in all {len(dist_xl)} xlsx quarters "
          f"(min {dist_xl.min():.1f}B) and all {len(dist_mine)} re-derived "
          "quarters — buyback sign not flipped")
else:
    check("FAIL", "LFCF_DIST sign convention",
          f"non-positive distributions: {n_neg_xl} xlsx rows, {n_neg_mine} "
          "re-derived rows — sign convention on net issuance looks flipped")

mc_dist_x = xl["MC_LFCF_DIST"].dropna().iloc[-1]
mc_dist_dt = xl["MC_LFCF_DIST"].dropna().index[-1]
mc_dist_m = mine.loc[mc_dist_dt, "MC_LFCF_DIST"]
if pd.isna(mc_dist_m):
    check("FAIL", "MC/dist latest", f"{mc_dist_dt.date()} not derivable")
elif rel(mc_dist_x, mc_dist_m) > REL_TOL:
    check("FAIL", "MC/dist latest",
          f"xlsx {mc_dist_x:.2f}x vs re-derived {mc_dist_m:.2f}x at "
          f"{mc_dist_dt.date()} (rel {rel(mc_dist_x, mc_dist_m):.2%})")
elif rel(mc_dist_x, 44.5) > 0.01:
    check("WARN", "MC/dist latest",
          f"{mc_dist_x:.2f}x matches re-derivation but is >1% from the "
          "expected ~44.5x benchmark (data revision?)")
else:
    check("PASS", "MC/dist latest",
          f"{mc_dist_x:.1f}x at {mc_dist_dt.date()} matches re-derivation "
          f"(rel {rel(mc_dist_x, mc_dist_m):.2e}) and the ~44.5x benchmark")

# ---------------------------------------------------------------------------
# 8. Mask rule: EV_UFCF / MC_LFCF NaN exactly where smoothed flow < 1% of
#    EV / MktCap (or rolling warm-up); recount the 8 and 63 masked rows
# ---------------------------------------------------------------------------
for ratio_col, flow_col, base_col, exp_nan in [
        ("EV_UFCF", "UFCF_8Q", "EV", 8),
        ("MC_LFCF", "LFCF_8Q", "MKTCAP", 63)]:
    m = mine.loc[mine.index.isin(xl.index)]
    expected_nan = m[flow_col].isna() | (m[flow_col] < 0.01 * m[base_col])
    actual_nan = xl[ratio_col].isna()
    common = expected_nan.index.intersection(actual_nan.index)
    mism = (expected_nan.loc[common] != actual_nan.loc[common])
    n_mism = int(mism.sum())
    n_actual = int(actual_nan.sum())
    # value agreement where unmasked
    both = (~expected_nan.loc[common]) & (~actual_nan.loc[common])
    vals_m = m.loc[common[both], base_col] / m.loc[common[both], flow_col]
    vals_x = xl.loc[common[both], ratio_col]
    max_rel = float((abs(vals_x - vals_m) / abs(vals_m)).max())
    if n_mism == 0 and n_actual == exp_nan and max_rel <= REL_TOL:
        check("PASS", f"mask rule {ratio_col}",
              f"{n_actual} NaN rows (expected {exp_nan}), NaN exactly where "
              f"{flow_col} < 1% of {base_col} or warm-up; unmasked values "
              f"match re-derivation (max rel {max_rel:.2e})")
    elif n_mism == 0 and max_rel <= REL_TOL:
        check("WARN", f"mask rule {ratio_col}",
              f"mask pattern matches the rule cellwise but count is "
              f"{n_actual}, not the stated {exp_nan} (revision moved a "
              "quarter across the 1% threshold?)")
    else:
        bad_dates = [d.date() for d in common[mism][:5]]
        check("FAIL", f"mask rule {ratio_col}",
              f"{n_mism} cells where mask disagrees with the <1% rule "
              f"(e.g. {bad_dates}); NaN count {n_actual} vs expected "
              f"{exp_nan}; max unmasked rel diff {max_rel:.2e}")

# ---------------------------------------------------------------------------
# 9. Date axis: monotonic, unique, quarterly, no gaps, starts 1952Q4
# ---------------------------------------------------------------------------
dates = data["Date"]
mono = dates.is_monotonic_increasing
dupes = int(dates.duplicated().sum())
qstart = bool(dates.dt.is_quarter_start.all())
periods = pd.PeriodIndex(dates, freq="Q")
missing_q = len(pd.period_range(periods[0], periods[-1], freq="Q")) - len(periods)
if mono and dupes == 0 and qstart and missing_q == 0:
    check("PASS", "date axis",
          f"{len(dates)} rows, strictly quarterly {dates.iloc[0].date()}.."
          f"{dates.iloc[-1].date()}, no dupes/gaps")
else:
    check("FAIL", "date axis",
          f"monotonic={mono}, dupes={dupes}, all quarter-start={qstart}, "
          f"missing quarters={missing_q}")

if dates.iloc[0] == pd.Timestamp("1952-10-01"):
    check("PASS", "start date",
          "Data sheet starts 1952-10-01 (documented intentional exception to "
          "the 2006 convention; 2006+ companion chart present)")
else:
    check("FAIL", "start date",
          f"Data sheet starts {dates.iloc[0].date()}, expected 1952-10-01")

age_days = (TODAY - latest_ratio).days
if age_days <= 200:
    check("PASS", "staleness",
          f"latest ratio quarter {latest_ratio.date()} is {age_days} days old "
          "(<= 200; Z.1/NIPA publication lag)")
else:
    check("FAIL", "staleness",
          f"latest ratio quarter {latest_ratio.date()} is {age_days} days old "
          "(> 200) — output looks stale")

# ---------------------------------------------------------------------------
# 10. Internal consistency: no interior NaNs in the non-masked series
# ---------------------------------------------------------------------------
NONMASK = [c for c in RATIO_COLS if c not in ("EV_UFCF", "MC_LFCF")]
holes = []
for col in NONMASK:
    s = xl[col]
    valid = s.dropna()
    if valid.empty:
        holes.append(f"{col}: entirely NaN")
        continue
    inner = s.loc[valid.index[0]:valid.index[-1]]
    if inner.isna().sum():
        holes.append(f"{col}: {int(inner.isna().sum())} interior NaN(s)")
if holes:
    check("FAIL", "interior NaNs", "; ".join(holes))
else:
    check("PASS", "interior NaNs",
          "no NaNs inside any non-masked ratio/yield series (EV_UFCF and "
          "MC_LFCF excluded — their NaNs are the documented mask)")

# ---------------------------------------------------------------------------
# 11. Internal consistency: no unit-jump discontinuities in levels/ratios
# ---------------------------------------------------------------------------
jump_fails = []
worst = ("", 0.0)
for col in ["EV", "MKTCAP", "DEBT", "LIQ", "EBITDA_4Q", "EV_EBIT",
            "EV_EBITDA", "MC_GVA"]:
    s = xl[col].dropna()
    if (s <= 0).any():
        jump_fails.append(f"{col}: non-positive values")
        continue
    lg = np.abs(np.log(s / s.shift(1))).dropna()
    if len(lg) and lg.max() >= 1.0:
        jump_fails.append(
            f"{col}: |q/q log change| {lg.max():.2f} at {lg.idxmax().date()}")
    if len(lg) and lg.max() > worst[1]:
        worst = (f"{col} @ {lg.idxmax().date()}", lg.max())
if jump_fails:
    check("FAIL", "unit discontinuities", "; ".join(jump_fails))
else:
    check("PASS", "unit discontinuities",
          f"max |q/q log change| across levels and multiples {worst[1]:.2f} "
          f"({worst[0]}) < 1.0 — no $MM/$B unit breaks")

# ---------------------------------------------------------------------------
# 12. Summary sheet consistent with Data sheet
# ---------------------------------------------------------------------------
sum_bad = []
for col in RATIO_COLS:
    s = xl[col].dropna()
    for stat, val in (("min", s.min()), ("max", s.max()),
                      ("median", s.median()), ("current", s.iloc[-1])):
        sv = summary.loc[col, stat]
        if abs(sv - val) / abs(val) > REL_TOL:
            sum_bad.append(f"{col}.{stat}: summary {sv:.4f} vs data {val:.4f}")
    pct = summary.loc[col, "current_pctile"]
    my_pct = (s <= s.iloc[-1]).mean() * 100
    if abs(pct - my_pct) > 0.5:
        sum_bad.append(f"{col}.pctile: summary {pct:.1f} vs data {my_pct:.1f}")
if sum_bad:
    check("FAIL", "Summary sheet consistency", "; ".join(sum_bad))
else:
    check("PASS", "Summary sheet consistency",
          "min/max/median/current/pctile for all 11 series match Data sheet")

# ---------------------------------------------------------------------------
# 13. EXTERNAL: research-probe benchmarks (EV/EBITDA ~98th pctile, median
#     ~6.8x) and equity leg vs the alternative FRED series NCBEILQ027S
# ---------------------------------------------------------------------------
pctile = summary.loc["EV_EBITDA", "current_pctile"]
med = summary.loc["EV_EBITDA", "median"]
if 95.0 <= pctile <= 99.5 and abs(med - 6.8) <= 0.3:
    check("PASS", "external EV/EBITDA benchmark",
          f"current pctile {pctile:.1f} (expected ~98th) and 1952Q4+ median "
          f"{med:.2f}x (expected ~6.8x) match the independent research probe")
else:
    check("FAIL", "external EV/EBITDA benchmark",
          f"current pctile {pctile:.1f} (expected 95-99.5) / median "
          f"{med:.2f}x (expected 6.8 +/- 0.3) off the research-probe values")

ncbeil = fred.get_series("NCBEILQ027S", observation_start="2000-01-01") / 1000.0
common_eq = ncbeil.dropna().index.intersection(xl["MKTCAP"].dropna().index)
if len(common_eq) == 0:
    check("WARN", "external equity cross-series", "no overlap with NCBEILQ027S")
else:
    diff = (xl.loc[common_eq, "MKTCAP"] - ncbeil.loc[common_eq]).abs() \
        / ncbeil.loc[common_eq]
    q = common_eq[-1]
    if diff.max() <= 0.005:
        check("PASS", "external equity cross-series",
              f"MKTCAP matches independent series NCBEILQ027S within "
              f"{diff.max():.2%} across {len(common_eq)} quarters "
              f"(latest ${xl.loc[q, 'MKTCAP']/1000:.1f}T at {q.date()})")
    else:
        check("FAIL", "external equity cross-series",
              f"MKTCAP diverges from NCBEILQ027S by up to {diff.max():.2%} "
              f"(worst at {diff.idxmax().date()})")

# ---------------------------------------------------------------------------
# 14. EXTERNAL: own BAA/AAA pull -> quarterly averages; latest ~0.48-0.54pp
# ---------------------------------------------------------------------------
baa = fred.get_series("BAA", observation_start="1945-01-01")
aaa = fred.get_series("AAA", observation_start="1945-01-01")
spread_m = (baa - aaa).dropna()
# Quarterly mean via period groupby — a different mechanism than the
# analysis' resample("QS"), so a quarter-labeling bug would show up here.
spread_q = spread_m.groupby(pd.PeriodIndex(spread_m.index, freq="Q")).mean()
spread_q.index = spread_q.index.to_timestamp(how="start")
common_sp = spread_q.index.intersection(xl["BAA_AAA_SPREAD"].dropna().index)
sp_diff = (xl.loc[common_sp, "BAA_AAA_SPREAD"] - spread_q.loc[common_sp]).abs()
sp_latest = xl["BAA_AAA_SPREAD"].dropna().iloc[-1]
sp_latest_dt = xl["BAA_AAA_SPREAD"].dropna().index[-1]
if sp_diff.max() > 0.005:
    check("FAIL", "external Baa-Aaa spread",
          f"xlsx quarterly spread deviates from own monthly-avg pull by up to "
          f"{sp_diff.max():.3f}pp (worst {sp_diff.idxmax().date()}) — "
          "quarter labeling/averaging error")
elif not (0.48 <= sp_latest <= 0.54):
    check("WARN", "external Baa-Aaa spread",
          f"latest {sp_latest:.2f}pp at {sp_latest_dt.date()} matches own "
          f"pull ({spread_q.loc[sp_latest_dt]:.2f}pp) but sits outside the "
          "expected 0.48-0.54pp band")
else:
    check("PASS", "external Baa-Aaa spread",
          f"latest {sp_latest:.2f}pp at {sp_latest_dt.date()} within expected "
          f"0.48-0.54pp; all {len(common_sp)} quarters match own BAA-AAA pull "
          f"(max diff {sp_diff.max():.4f}pp)")

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
n_pass = RESULTS.count("PASS")
n_warn = RESULTS.count("WARN")
n_fail = RESULTS.count("FAIL")
print(f"\nSummary: {n_pass} PASS, {n_warn} WARN, {n_fail} FAIL "
      f"({len(RESULTS)} checks)")
sys.exit(1 if n_fail else 0)
