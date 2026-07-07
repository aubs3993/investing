# Independent audit of macro/valuation_multiple_drivers_pull.py and
# macro/output/valuation_multiple_drivers/.
#
# Re-pulls all raw FRED series itself (A463/B471/B456/A455/B075/A053/Y001/PNFI
# NIPA flows, NCBCEL/BCNSDODNS/BOGZ1FL104001005Q Z.1 levels, GS10/CPIAUCSL/
# REAINTRATREARAT10Y monthly rates) and re-derives the driver series with its
# own code — time-aware lookbacks instead of positional shifts, so positional
# off-by-one bugs in the analysis would surface as mismatches. Does NOT import
# or exec valuation_multiple_drivers_pull.py; reads the xlsx only to compare.
#
# Tolerances:
#   - Value re-derivations: 0.1% relative (same-day FRED vintage; only float
#     round-trip through xlsx expected; unit errors are 1000x, series mix-ups
#     several %, off-by-one rolling windows typically >0.5% in volatile spans).
#   - Benchmark bands (2016Q4 tax ~13.9%, 2018-19 ~9-10%, TRA86 rise, margin
#     ~35% vs ~27.8% median, intangible share 40%+): generous bands, since
#     NIPA revisions can move decades-old levels slightly.
#   - Splice boundary: WARN (not FAIL) if Dec-1981 proxy vs Jan-1982 Cleveland
#     model differ by >150bp — a known series-construction difference; the
#     chart labels the two segments separately, so it is a labeling judgment.
#   - Staleness: latest NIPA quarter within 200 days of today.
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

OUT_DIR = MACRO_DIR / "output" / "valuation_multiple_drivers"
XLSX = OUT_DIR / "valuation_multiple_drivers.xlsx"
EV_MULT_XLSX = MACRO_DIR / "output" / "ev_multiples" / "ev_multiples.xlsx"

TODAY = pd.Timestamp(datetime.today().date())
REL_TOL = 0.001

RESULTS = []


def check(status: str, name: str, detail: str) -> None:
    RESULTS.append(status)
    print(f"{status} {name}: {detail}")


# ---------------------------------------------------------------------------
# 1. Output file existence + non-trivial size (+ .gitkeep repo convention)
# ---------------------------------------------------------------------------
EXPECTED_FILES = {
    "valuation_multiple_drivers.xlsx": 10_000,
    "da_share.png": 20_000,
    "effective_tax_rate.png": 20_000,
    "earnings_yield_vs_real10y.png": 20_000,
    "ebitda_margin_ev_gva.png": 20_000,
    "intangible_share.png": 20_000,
    "drivers_overview.png": 20_000,
    "drivers_overview_2006.png": 20_000,
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

if (OUT_DIR / ".gitkeep").exists():
    check("PASS", "output .gitkeep", "present (repo convention)")
else:
    check("WARN", "output .gitkeep", "missing from output subfolder")

if not XLSX.exists():
    print("\nSummary: cannot continue without valuation_multiple_drivers.xlsx")
    sys.exit(1)

data = pd.read_excel(XLSX, sheet_name="Data")
data["Date"] = pd.to_datetime(data["Date"])
summary = pd.read_excel(XLSX, sheet_name="Summary", index_col=0)
xl = data.set_index("Date")

DERIVED_COLS = [
    "DA_SHARE", "DA_SHARE_4Q",
    "EFF_TAX", "EFF_TAX_4Q",
    "EBIT_YIELD", "EBITDA_YIELD", "REAL10Y", "EBIT_SPREAD",
    "EBITDA_MARGIN", "EV_GVA",
    "INTANG_SHARE", "INTANG_SHARE_4Q",
]

# ---------------------------------------------------------------------------
# 2. Independent raw pulls from FRED
# ---------------------------------------------------------------------------
fred = get_fred_client()
RAW_Q = {
    "PROFITS_IVA": "A463RC1Q027SBEA",
    "NET_INTEREST": "B471RC1Q027SBEA",
    "DA": "B456RC1Q027SBEA",
    "GVA": "A455RC1Q027SBEA",
    "FED_TAX": "B075RC1Q027SBEA",
    "PBT": "A053RC1Q027SBEA",
    "IPP": "Y001RC1Q027SBEA",
    "PNFI": "PNFI",
    "EQ_MM": "NCBCEL",
    "DEBT_MM": "BCNSDODNS",
    "LIQ_MM": "BOGZ1FL104001005Q",
}
RAW_M = {
    "REAL10Y_CLEV": "REAINTRATREARAT10Y",
    "GS10": "GS10",
    "CPI": "CPIAUCSL",
}
rq = {n: fred.get_series(s, observation_start="1945-01-01")
      for n, s in RAW_Q.items()}
rm = {n: fred.get_series(s, observation_start="1947-01-01")
      for n, s in RAW_M.items()}

# Own derivation: each series' 4q trailing mean on its OWN native quarterly
# index, then join — a positional-window bug across merge-induced gap rows in
# the analysis would show up as a mismatch rather than being replicated.
mine = pd.DataFrame(rq)
mine["EBIT"] = mine["PROFITS_IVA"] + mine["NET_INTEREST"]
mine["EBITDA"] = mine["EBIT"] + mine["DA"]
mine["EBIT_4Q"] = mine["EBIT"].rolling(4).mean()
mine["EBITDA_4Q"] = mine["EBITDA"].rolling(4).mean()
mine["GVA_4Q"] = mine["GVA"].rolling(4).mean()
mine["DA_SHARE"] = mine["DA"] / mine["EBITDA"] * 100
mine["DA_SHARE_4Q"] = mine["DA_SHARE"].rolling(4).mean()
mine["EFF_TAX"] = mine["FED_TAX"] / mine["PBT"] * 100
mine["EFF_TAX_4Q"] = mine["EFF_TAX"].rolling(4).mean()
mine["EV"] = (mine["EQ_MM"] + mine["DEBT_MM"] - mine["LIQ_MM"]) / 1000.0
mine["EBIT_YIELD"] = mine["EBIT_4Q"] / mine["EV"] * 100
mine["EBITDA_YIELD"] = mine["EBITDA_4Q"] / mine["EV"] * 100
mine["EBITDA_MARGIN"] = mine["EBITDA_4Q"] / mine["GVA_4Q"] * 100
mine["EV_GVA"] = mine["EV"] / mine["GVA_4Q"]
mine["INTANG_SHARE"] = mine["IPP"] / mine["PNFI"] * 100
mine["INTANG_SHARE_4Q"] = mine["INTANG_SHARE"].rolling(4).mean()


def rel_cmp(name, xlsx_val, my_val, tol=REL_TOL, extra=""):
    """PASS/FAIL a relative comparison between xlsx and independent value."""
    if pd.isna(xlsx_val) or pd.isna(my_val):
        check("FAIL", name, f"NaN encountered (xlsx={xlsx_val}, mine={my_val})")
        return False
    rel = abs(xlsx_val - my_val) / abs(my_val)
    if rel <= tol:
        check("PASS", name,
              f"xlsx {xlsx_val:.4f} vs independent {my_val:.4f} "
              f"(rel {rel:.2e}){extra}")
        return True
    check("FAIL", name,
          f"xlsx {xlsx_val:.4f} vs independent {my_val:.4f} (rel {rel:.2%})")
    return False


# ---------------------------------------------------------------------------
# 3. Re-derivation 1: effective federal tax rate (B075/A053)
# ---------------------------------------------------------------------------
q2016 = pd.Timestamp("2016-10-01")
rel_cmp("eff tax 2016Q4 (4q) re-derivation",
        xl.loc[q2016, "EFF_TAX_4Q"], mine.loc[q2016, "EFF_TAX_4Q"])
v2016 = xl.loc[q2016, "EFF_TAX_4Q"]
if 13.0 <= v2016 <= 15.0:
    check("PASS", "eff tax 2016Q4 benchmark",
          f"{v2016:.2f}% within expected ~13.9% band (13-15%)")
else:
    check("FAIL", "eff tax 2016Q4 benchmark",
          f"{v2016:.2f}% outside expected ~13.9% band (13-15%)")

tcja = xl.loc["2018-01-01":"2019-12-31", "EFF_TAX_4Q"]
my_tcja = mine.loc["2018-01-01":"2019-12-31", "EFF_TAX_4Q"]
rel_cmp("eff tax 2018-19 avg re-derivation", tcja.mean(), my_tcja.mean())
if 8.5 <= tcja.mean() <= 11.0:
    check("PASS", "eff tax TCJA benchmark",
          f"2018-19 avg {tcja.mean():.2f}% within expected 9-10% band "
          f"(vs {v2016:.1f}% in 2016Q4 — clean TCJA step-down confirmed)")
else:
    check("FAIL", "eff tax TCJA benchmark",
          f"2018-19 avg {tcja.mean():.2f}% outside expected 8.5-11% band")

# TRA86 direction: the rate must RISE ~20% -> ~25% across 1985 -> 1987-88
# (base-broadening beat the statutory cut). A step-DOWN claim would be a
# defect; the script's footnote asserts the rise — verify the data agrees.
pre = xl.loc["1985-01-01":"1985-12-31", "EFF_TAX_4Q"].mean()
post = xl.loc["1987-01-01":"1988-12-31", "EFF_TAX_4Q"].mean()
my_pre = mine.loc["1985-01-01":"1985-12-31", "EFF_TAX_4Q"].mean()
my_post = mine.loc["1987-01-01":"1988-12-31", "EFF_TAX_4Q"].mean()
if post > pre + 2.0 and 17.0 <= pre <= 23.0 and 23.0 <= post <= 28.0:
    check("PASS", "TRA86 direction",
          f"eff tax ROSE {pre:.1f}% (1985) -> {post:.1f}% (1987-88) as the "
          f"script's footnote claims (independent: {my_pre:.1f} -> "
          f"{my_post:.1f}); no false step-down")
elif post > pre:
    check("WARN", "TRA86 direction",
          f"rate rose {pre:.1f}% -> {post:.1f}% but outside the expected "
          f"~20%->~25% bands")
else:
    check("FAIL", "TRA86 direction",
          f"eff tax FELL {pre:.1f}% (1985) -> {post:.1f}% (1987-88); the "
          f"chart footnote's claimed TRA86 rise is contradicted by the data")

# ---------------------------------------------------------------------------
# 4. Re-derivation 2: D&A share of EBITDA — latest + 1950Q1 + long-run uptrend
# ---------------------------------------------------------------------------
da_latest_q = xl["DA_SHARE_4Q"].dropna().index[-1]
rel_cmp(f"D&A share latest ({da_latest_q.date()}) re-derivation",
        xl.loc[da_latest_q, "DA_SHARE_4Q"], mine.loc[da_latest_q, "DA_SHARE_4Q"])
q1950 = pd.Timestamp("1950-01-01")
rel_cmp("D&A share 1950Q1 re-derivation",
        xl.loc[q1950, "DA_SHARE"], mine.loc[q1950, "DA_SHARE"])
fifties_da = xl.loc["1950-01-01":"1959-12-31", "DA_SHARE_4Q"].mean()
latest_da = xl.loc[da_latest_q, "DA_SHARE_4Q"]
if latest_da > fifties_da + 5.0:
    check("PASS", "D&A share long-run uptrend",
          f"latest {latest_da:.1f}% vs 1950s avg {fifties_da:.1f}% — "
          f"uptrend confirmed")
else:
    check("FAIL", "D&A share long-run uptrend",
          f"latest {latest_da:.1f}% not meaningfully above 1950s avg "
          f"{fifties_da:.1f}%")

# ---------------------------------------------------------------------------
# 5. Re-derivation 3: real-10y splice — proxy at 1975Q1 + 1982 boundary
# ---------------------------------------------------------------------------
# Time-aware trailing-10y CPI inflation (vs the analysis' positional
# shift(120)): for each month m, infl = (CPI[m]/CPI[m-10y])^(1/10)-1.
cpi = rm["CPI"].dropna()
gs10 = rm["GS10"].dropna()
clev = rm["REAL10Y_CLEV"].dropna()

def proxy_month(m: pd.Timestamp) -> float:
    m0 = m - pd.DateOffset(years=10)
    if m not in cpi.index or m0 not in cpi.index or m not in gs10.index:
        return np.nan
    infl = ((cpi.loc[m] / cpi.loc[m0]) ** (1 / 10) - 1) * 100
    return gs10.loc[m] - infl

q1975 = pd.Timestamp("1975-01-01")
my_proxy_75 = np.mean([proxy_month(pd.Timestamp(f"1975-{mm:02d}-01"))
                       for mm in (1, 2, 3)])
xl_75 = xl.loc[q1975, "REAL10Y"]
if pd.isna(xl_75) or pd.isna(my_proxy_75):
    check("FAIL", "real-10y proxy 1975Q1 re-derivation",
          f"NaN (xlsx={xl_75}, mine={my_proxy_75})")
elif abs(xl_75 - my_proxy_75) <= 0.02:
    check("PASS", "real-10y proxy 1975Q1 re-derivation",
          f"xlsx REAL10Y {xl_75:.3f}% vs time-aware re-derivation "
          f"{my_proxy_75:.3f}% (diff {abs(xl_75 - my_proxy_75)*100:.2f}bp) — "
          f"GS10 minus trailing-10y CPI, quarterly-averaged")
else:
    check("FAIL", "real-10y proxy 1975Q1 re-derivation",
          f"xlsx REAL10Y {xl_75:.3f}% vs re-derivation {my_proxy_75:.3f}% "
          f"(diff {abs(xl_75 - my_proxy_75)*100:.0f}bp) — possible positional "
          f"shift(120) misalignment or wrong lookback")

# xlsx must use the PROXY (not Cleveland) at 1975Q1 and Cleveland from 1982.
p75 = xl.loc[q1975, "REAL10Y_PROXY"]
c82 = xl.loc[pd.Timestamp("1982-01-01"), "REAL10Y_CLEV"]
p82 = xl.loc[pd.Timestamp("1982-01-01"), "REAL10Y_PROXY"]
if pd.notna(p75) and pd.notna(c82) and pd.isna(p82):
    check("PASS", "splice segment assignment",
          "1975Q1 carried as proxy segment; 1982Q1 carried as Cleveland "
          "segment with proxy masked — splice boundary where documented")
else:
    check("FAIL", "splice segment assignment",
          f"unexpected segment layout: proxy@1975Q1={p75}, "
          f"clev@1982Q1={c82}, proxy@1982Q1={p82}")

dec81 = proxy_month(pd.Timestamp("1981-12-01"))
jan82_clev = clev.loc[pd.Timestamp("1982-01-01")] \
    if pd.Timestamp("1982-01-01") in clev.index else np.nan
disc = abs(dec81 - jan82_clev)
if pd.isna(disc):
    check("FAIL", "1982 splice discontinuity",
          f"could not compute (proxy Dec-81={dec81}, Clev Jan-82={jan82_clev})")
elif disc <= 1.5:
    check("PASS", "1982 splice discontinuity",
          f"Dec-1981 proxy {dec81:.2f}% vs Jan-1982 Cleveland "
          f"{jan82_clev:.2f}% — {disc*100:.0f}bp gap (<= 150bp)")
else:
    check("WARN", "1982 splice discontinuity",
          f"Dec-1981 proxy {dec81:.2f}% vs Jan-1982 Cleveland "
          f"{jan82_clev:.2f}% — {disc*100:.0f}bp gap (> 150bp). Known "
          f"series-construction difference (backward-looking CPI proxy vs "
          f"Cleveland expectations model); the chart does plot the segments "
          f"as separate labeled lines (dashed 'proxy' vs 'Cleveland Fed'), "
          f"but the spliced EBIT_SPREAD subpanel crosses this jump silently")

# ---------------------------------------------------------------------------
# 6. Re-derivation 4: intangible share of PNFI — latest 40%+, 1950s single-digit
# ---------------------------------------------------------------------------
int_latest_q = xl["INTANG_SHARE_4Q"].dropna().index[-1]
rel_cmp(f"intangible share latest ({int_latest_q.date()}) re-derivation",
        xl.loc[int_latest_q, "INTANG_SHARE_4Q"],
        mine.loc[int_latest_q, "INTANG_SHARE_4Q"])
int_latest = xl.loc[int_latest_q, "INTANG_SHARE_4Q"]
fifties_int = xl.loc["1950-01-01":"1959-12-31", "INTANG_SHARE_4Q"].mean()
if int_latest >= 38.0 and fifties_int < 10.0:
    check("PASS", "intangible share benchmark",
          f"latest {int_latest:.1f}% (>= 38%), 1950s avg {fifties_int:.1f}% "
          f"(single digits) — structural rise confirmed")
else:
    check("FAIL", "intangible share benchmark",
          f"latest {int_latest:.1f}% / 1950s avg {fifties_int:.1f}% outside "
          f"expected (>=38% now, <10% in the 1950s)")

# ---------------------------------------------------------------------------
# 7. Re-derivation 5: EBITDA margin on GVA — latest ~35% vs ~27.8% median
# ---------------------------------------------------------------------------
mar_latest_q = xl["EBITDA_MARGIN"].dropna().index[-1]
rel_cmp(f"EBITDA margin latest ({mar_latest_q.date()}) re-derivation",
        xl.loc[mar_latest_q, "EBITDA_MARGIN"],
        mine.loc[mar_latest_q, "EBITDA_MARGIN"])
mar_latest = xl.loc[mar_latest_q, "EBITDA_MARGIN"]
mar_median = xl["EBITDA_MARGIN"].dropna().median()
if 33.0 <= mar_latest <= 37.0 and 26.5 <= mar_median <= 29.0:
    check("PASS", "EBITDA margin benchmark",
          f"latest {mar_latest:.1f}% (expected ~35%), 1947+ median "
          f"{mar_median:.1f}% (expected ~27.8%)")
else:
    check("FAIL", "EBITDA margin benchmark",
          f"latest {mar_latest:.1f}% / median {mar_median:.1f}% outside "
          f"expected (~35% / ~27.8%)")

# ---------------------------------------------------------------------------
# 8. EV, yields and the margin/multiple identity at the latest EV quarter
# ---------------------------------------------------------------------------
ev_latest_q = xl["EV"].dropna().index[-1]
rel_cmp(f"EV latest ({ev_latest_q.date()}) re-derivation ($MM->$BN)",
        xl.loc[ev_latest_q, "EV"], mine.loc[ev_latest_q, "EV"],
        extra=" — NCBCEL + BCNSDODNS - liquid assets, /1000")
rel_cmp(f"EBIT yield latest ({ev_latest_q.date()}) re-derivation",
        xl.loc[ev_latest_q, "EBIT_YIELD"], mine.loc[ev_latest_q, "EBIT_YIELD"])
rel_cmp(f"EV/GVA latest ({ev_latest_q.date()}) re-derivation",
        xl.loc[ev_latest_q, "EV_GVA"], mine.loc[ev_latest_q, "EV_GVA"])

# Identity check inside the xlsx: EV/EBITDA(4q) == EV_GVA / (margin/100).
sub = xl[["EV", "EBITDA_4Q", "EV_GVA", "EBITDA_MARGIN"]].dropna()
lhs = sub["EV"] / sub["EBITDA_4Q"]
rhs = sub["EV_GVA"] / (sub["EBITDA_MARGIN"] / 100)
ident_max = (lhs - rhs).abs().max()
if ident_max < 1e-9:
    check("PASS", "margin-vs-multiple identity",
          f"EV/EBITDA(4q) == (EV/GVA)/(EBITDA margin) holds exactly on all "
          f"{len(sub)} rows (max abs diff {ident_max:.1e})")
else:
    check("FAIL", "margin-vs-multiple identity",
          f"identity broken: max abs diff {ident_max:.2e} — the two legs are "
          f"not on consistent 4q averaging")

# ---------------------------------------------------------------------------
# 9. EXTERNAL cross-check: EV series equals ev_multiples module's EV
# ---------------------------------------------------------------------------
if not EV_MULT_XLSX.exists():
    check("WARN", "external EV vs ev_multiples",
          f"{EV_MULT_XLSX} not found — cannot cross-check")
else:
    ev2 = pd.read_excel(EV_MULT_XLSX, sheet_name="Data")
    ev2["Date"] = pd.to_datetime(ev2["Date"])
    ev2 = ev2.set_index("Date")["EV"].dropna()
    ev1 = xl["EV"].dropna()
    common = ev1.index.intersection(ev2.index)
    if len(common) < 100:
        check("WARN", "external EV vs ev_multiples",
              f"only {len(common)} common quarters")
    else:
        rel = ((ev1.loc[common] - ev2.loc[common]).abs()
               / ev2.loc[common].abs()).max()
        latest_common = common[-1]
        if rel <= 1e-9:
            check("PASS", "external EV vs ev_multiples",
                  f"EV identical on all {len(common)} common quarters "
                  f"(max rel diff {rel:.1e}); latest {latest_common.date()} "
                  f"${ev1.loc[latest_common]:,.0f}B both modules — same "
                  f"NCBCEL/BCNSDODNS/BOGZ1FL104001005Q construction")
        elif rel <= REL_TOL:
            check("WARN", "external EV vs ev_multiples",
                  f"EV agrees within {rel:.2e} but not bit-identical — "
                  f"different pull vintages?")
        else:
            check("FAIL", "external EV vs ev_multiples",
                  f"EV diverges up to {rel:.2%} between the two modules — "
                  f"construction mismatch")

# ---------------------------------------------------------------------------
# 10. Data-range checks: start date + staleness
# ---------------------------------------------------------------------------
first_date = data["Date"].iloc[0]
if first_date == pd.Timestamp("1947-01-01"):
    check("PASS", "start date", "Data sheet starts 1947-01-01 (documented "
          "intentional exception to the 2006 convention)")
else:
    check("FAIL", "start date",
          f"Data sheet starts {first_date.date()}, expected 1947-01-01")

latest_nipa_q = xl["EFF_TAX"].dropna().index[-1]
age_days = (TODAY - latest_nipa_q).days
if age_days <= 200:
    check("PASS", "staleness",
          f"latest NIPA quarter {latest_nipa_q.date()} is {age_days} days old "
          f"(<= 200; publication lag)")
else:
    check("FAIL", "staleness",
          f"latest NIPA quarter {latest_nipa_q.date()} is {age_days} days old "
          f"(> 200) — output looks stale")

# ---------------------------------------------------------------------------
# 11. Internal consistency: dates monotonic, unique, quarterly, gap-free
# ---------------------------------------------------------------------------
dates = data["Date"]
diffs_ok = dates.is_monotonic_increasing
dupes = dates.duplicated().sum()
qstart = dates.dt.is_quarter_start.all()
periods = pd.PeriodIndex(dates, freq="Q")
full = pd.period_range(periods[0], periods[-1], freq="Q")
missing_q = len(full) - len(periods)
if diffs_ok and dupes == 0 and qstart and missing_q == 0:
    check("PASS", "date axis", f"{len(dates)} rows, strictly quarterly "
          f"{dates.iloc[0].date()}..{dates.iloc[-1].date()}, no dupes/gaps")
else:
    check("FAIL", "date axis",
          f"monotonic={diffs_ok}, dupes={dupes}, all quarter-start={qstart}, "
          f"missing quarters={missing_q}")

# ---------------------------------------------------------------------------
# 12. Internal consistency: no interior NaN runs (per-column valid start)
# ---------------------------------------------------------------------------
# NIPA-only columns are dense from 1947/1948; EV-based columns are Q4-only
# pre-1952 by design (Z.1 annual); REAL10Y starts 1957 (proxy lookback).
COL_START = {
    "DA_SHARE": "1947-01-01", "DA_SHARE_4Q": "1947-10-01",
    "EFF_TAX": "1947-01-01", "EFF_TAX_4Q": "1947-10-01",
    "INTANG_SHARE": "1947-01-01", "INTANG_SHARE_4Q": "1947-10-01",
    "EBITDA_MARGIN": "1947-10-01",
    "EBIT_YIELD": "1952-10-01", "EBITDA_YIELD": "1952-10-01",
    "EV_GVA": "1952-10-01", "EBIT_SPREAD": "1957-01-01",
    "REAL10Y": "1957-01-01",
}
holes = []
for col, start in COL_START.items():
    s = xl.loc[xl.index >= start, col]
    valid = s.dropna()
    if valid.empty:
        holes.append(f"{col}: entirely NaN after {start}")
        continue
    inner = s.loc[valid.index[0]:valid.index[-1]]
    n_holes = int(inner.isna().sum())
    if n_holes:
        holes.append(f"{col}: {n_holes} interior NaN(s) after {start}")
if holes:
    check("FAIL", "interior NaNs", "; ".join(holes))
else:
    check("PASS", "interior NaNs",
          "no NaNs inside any derived series after its documented start "
          "(pre-1952 Q4-only EV rows and trailing publication-lag NaNs allowed)")

# ---------------------------------------------------------------------------
# 13. Internal consistency: no unit-jump discontinuities
# ---------------------------------------------------------------------------
jump_fails = []
worst = ("", 0.0)
for col in ["EV", "EBITDA_4Q", "GVA_4Q", "EBIT_YIELD", "EBITDA_YIELD",
            "EV_GVA", "EBITDA_MARGIN", "DA_SHARE_4Q"]:
    s = xl.loc[xl.index >= "1952-10-01", col].dropna()
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
          f"max |q/q log change| across level/ratio series {worst[1]:.2f} "
          f"({worst[0]}) < 1.0 — no 1000x unit breaks")

# ---------------------------------------------------------------------------
# 14. Summary sheet consistent with Data sheet
# ---------------------------------------------------------------------------
sum_bad = []
for col in DERIVED_COLS:
    s = xl[col].dropna()
    for stat, val in (("min", s.min()), ("max", s.max()),
                      ("mean", s.mean()), ("median", s.median()),
                      ("current", s.iloc[-1])):
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
          "min/max/mean/median/current/pctile for all 12 derived series "
          "match Data sheet")

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
n_pass = RESULTS.count("PASS")
n_warn = RESULTS.count("WARN")
n_fail = RESULTS.count("FAIL")
print(f"\nSummary: {n_pass} PASS, {n_warn} WARN, {n_fail} FAIL "
      f"({len(RESULTS)} checks)")
sys.exit(1 if n_fail else 0)
