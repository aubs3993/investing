# Independent audit of macro/nipa_pe_pull.py and macro/output/nipa_pe/.
#
# Re-pulls the six raw FRED series itself (BOGZ1LM883164105Q, NCBEILQ027S, CP,
# CPATAX, NFCPATAX, CPROFIT) and re-derives every ratio variant with its own
# code, then compares against the analysis outputs. Does NOT import or exec
# nipa_pe_pull.py; reads nipa_pe.xlsx only to compare.
#
# Tolerances:
#   - Value re-derivations: 0.1% relative. The audit runs against the same
#     same-day FRED vintage as the analysis, so the only expected divergence is
#     float round-trip through xlsx (~1e-15); 0.1% leaves headroom for an
#     intraday FRED revision without masking a real formula/unit error (unit
#     errors are 1000x, IVA/CCAdj series mix-ups are several %).
#   - Historical benchmark (1978Q4 trough ~5.96x): 2% relative, since decades-
#     old NIPA data can still move slightly in comprehensive revisions.
#   - Staleness: latest ratio quarter within 200 days of today (Z.1 + NIPA
#     third estimate lag the quarter end by ~ one quarter plus ~10 weeks).
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

OUT_DIR = MACRO_DIR / "output" / "nipa_pe"
XLSX = OUT_DIR / "nipa_pe.xlsx"

TODAY = pd.Timestamp(datetime.today().date())

RESULTS = []


def check(status: str, name: str, detail: str) -> None:
    RESULTS.append(status)
    print(f"{status} {name}: {detail}")


# ---------------------------------------------------------------------------
# 1. Output file existence + non-trivial size
# ---------------------------------------------------------------------------
EXPECTED_FILES = {
    "nipa_pe.xlsx": 10_000,
    "nipa_pe_full.png": 20_000,
    "nipa_pe_scope.png": 20_000,
    "nipa_pe_pretax.png": 20_000,
    "nipa_pe_2006.png": 20_000,
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
    print("\nSummary: cannot continue without nipa_pe.xlsx")
    sys.exit(1)

data = pd.read_excel(XLSX, sheet_name="Data")
data["Date"] = pd.to_datetime(data["Date"])
summary = pd.read_excel(XLSX, sheet_name="Summary", index_col=0)

RATIO_COLS = [
    "PE_ALLCORP_CPATAX",
    "PE_ALLCORP_CPATAX_4Q",
    "PE_ALLCORP_CP",
    "PE_ALLCORP_CP_4Q",
    "PE_NCB_NFCPATAX",
    "PE_ALLCORP_PRETAX",
]

# ---------------------------------------------------------------------------
# 2. Independent raw pulls from FRED
# ---------------------------------------------------------------------------
fred = get_fred_client()
RAW_IDS = {
    "EQ_ALLCORP_MM": "BOGZ1LM883164105Q",
    "EQ_NCB_MM": "NCBEILQ027S",
    "CP": "CP",
    "CPATAX": "CPATAX",
    "NFCPATAX": "NFCPATAX",
    "CPROFIT": "CPROFIT",
}
raw = {}
for name, sid in RAW_IDS.items():
    raw[name] = fred.get_series(sid, observation_start="1945-01-01")

# Own derivation, built deliberately differently from the analysis script:
# each profit series' 4q trailing mean is computed on its OWN native quarterly
# index (time-aware), then everything is joined — so a positional-window bug in
# the analysis (e.g. rolling across merge-induced gap rows) would show up as a
# mismatch rather than being replicated.
eq_allcorp = raw["EQ_ALLCORP_MM"] / 1000.0   # $MM -> $B
eq_ncb = raw["EQ_NCB_MM"] / 1000.0
cpatax_4q = raw["CPATAX"].rolling(4).mean()
cp_4q = raw["CP"].rolling(4).mean()

mine = pd.DataFrame({
    "EQ_ALLCORP": eq_allcorp,
    "EQ_NCB": eq_ncb,
    "CP": raw["CP"],
    "CPATAX": raw["CPATAX"],
    "NFCPATAX": raw["NFCPATAX"],
    "CPROFIT": raw["CPROFIT"],
    "CPATAX_4Q": cpatax_4q,
    "CP_4Q": cp_4q,
})
mine["PE_ALLCORP_CPATAX"] = mine["EQ_ALLCORP"] / mine["CPATAX"]
mine["PE_ALLCORP_CPATAX_4Q"] = mine["EQ_ALLCORP"] / mine["CPATAX_4Q"]
mine["PE_ALLCORP_CP"] = mine["EQ_ALLCORP"] / mine["CP"]
mine["PE_ALLCORP_CP_4Q"] = mine["EQ_ALLCORP"] / mine["CP_4Q"]
mine["PE_NCB_NFCPATAX"] = mine["EQ_NCB"] / mine["NFCPATAX"]
mine["PE_ALLCORP_PRETAX"] = mine["EQ_ALLCORP"] / mine["CPROFIT"]

xl = data.set_index("Date")

# ---------------------------------------------------------------------------
# 3. Spot re-derivations at 4 quarters, all 6 ratio variants, 0.1% tolerance
# ---------------------------------------------------------------------------
latest_common = mine.dropna(
    subset=["EQ_ALLCORP", "EQ_NCB", "CP", "CPATAX", "NFCPATAX", "CPROFIT"]
).index[-1]
SPOTS = [pd.Timestamp("1952-10-01"), pd.Timestamp("1978-10-01"),
         pd.Timestamp("2000-01-01"), latest_common]
REL_TOL = 0.001

for spot in SPOTS:
    label = f"spot re-derivation {spot.date()}"
    if spot not in xl.index:
        check("FAIL", label, "date missing from xlsx Data sheet")
        continue
    if spot not in mine.index:
        check("FAIL", label, "date missing from independent FRED pull")
        continue
    worst_col, worst_rel = None, 0.0
    bad = []
    for col in RATIO_COLS:
        mv, xv = mine.loc[spot, col], xl.loc[spot, col]
        if pd.isna(mv) or pd.isna(xv):
            bad.append(f"{col}: mine={mv}, xlsx={xv} (NaN)")
            continue
        rel = abs(xv - mv) / abs(mv)
        if rel > worst_rel:
            worst_col, worst_rel = col, rel
        if rel > REL_TOL:
            bad.append(f"{col}: mine={mv:.4f} xlsx={xv:.4f} rel={rel:.2%}")
    if bad:
        check("FAIL", label, "; ".join(bad))
    else:
        check("PASS", label,
              f"all 6 ratios within {REL_TOL:.1%} "
              f"(worst {worst_col} rel diff {worst_rel:.2e})")

# ---------------------------------------------------------------------------
# 4. Millions -> billions handling: latest AllCorp equities level
# ---------------------------------------------------------------------------
eq_latest_date = eq_allcorp.dropna().index[-1]
mine_eq_b = eq_allcorp.dropna().iloc[-1]                     # $B
raw_mm = raw["EQ_ALLCORP_MM"].dropna().iloc[-1]              # $MM as pulled
if eq_latest_date in xl.index and not pd.isna(xl.loc[eq_latest_date, "EQ_ALLCORP"]):
    xl_eq = xl.loc[eq_latest_date, "EQ_ALLCORP"]
    rel = abs(xl_eq - mine_eq_b) / mine_eq_b
    trillions = xl_eq / 1000.0
    if rel > REL_TOL:
        check("FAIL", "millions->billions conversion",
              f"xlsx EQ_ALLCORP {xl_eq:,.0f}B vs raw/1000 {mine_eq_b:,.0f}B "
              f"(rel {rel:.2%}) at {eq_latest_date.date()}")
    elif not (50_000 <= xl_eq <= 150_000):
        check("FAIL", "millions->billions conversion",
              f"latest AllCorp equities ${trillions:.1f}T outside plausible "
              f"$50-150T band — unit error (raw was {raw_mm:,.0f}MM)")
    else:
        check("PASS", "millions->billions conversion",
              f"latest AllCorp equities ${trillions:.1f}T at "
              f"{eq_latest_date.date()} (raw {raw_mm:,.0f}MM / 1000 matches "
              f"xlsx within {rel:.2e})")
else:
    check("FAIL", "millions->billions conversion",
          f"latest raw equity quarter {eq_latest_date.date()} absent/NaN in xlsx")

# ---------------------------------------------------------------------------
# 5. 4q-avg headline averages PROFITS, not the RATIO
# ---------------------------------------------------------------------------
# Find the post-1952 quarter where ratio-of-avg vs avg-of-ratio diverge most
# (profit swings, e.g. GFC), so the test actually discriminates.
avg_of_ratio = mine["PE_ALLCORP_CPATAX"].rolling(4).mean()
ratio_of_avg = mine["PE_ALLCORP_CPATAX_4Q"]
mask = (mine.index >= "1952-01-01") & ratio_of_avg.notna() & avg_of_ratio.notna()
gap = ((avg_of_ratio - ratio_of_avg).abs() / ratio_of_avg)[mask]
disc_q = gap.idxmax()
if gap.loc[disc_q] < 0.01:
    check("WARN", "4q-avg construction",
          "could not find a quarter where the two constructions differ >1%; "
          "test not discriminating")
elif disc_q not in xl.index:
    check("FAIL", "4q-avg construction", f"{disc_q.date()} missing from xlsx")
else:
    xv = xl.loc[disc_q, "PE_ALLCORP_CPATAX_4Q"]
    d_correct = abs(xv - ratio_of_avg.loc[disc_q]) / ratio_of_avg.loc[disc_q]
    d_wrong = abs(xv - avg_of_ratio.loc[disc_q]) / avg_of_ratio.loc[disc_q]
    if d_correct <= REL_TOL and d_wrong > 0.01:
        check("PASS", "4q-avg construction",
              f"at {disc_q.date()} xlsx {xv:.3f} = EQ / mean(4q CPATAX) "
              f"{ratio_of_avg.loc[disc_q]:.3f}, not mean of ratio "
              f"{avg_of_ratio.loc[disc_q]:.3f}")
    elif d_wrong <= REL_TOL:
        check("FAIL", "4q-avg construction",
              f"at {disc_q.date()} xlsx {xv:.3f} matches 4q mean of the RATIO "
              f"({avg_of_ratio.loc[disc_q]:.3f}), not ratio of 4q-mean profits "
              f"({ratio_of_avg.loc[disc_q]:.3f})")
    else:
        check("FAIL", "4q-avg construction",
              f"at {disc_q.date()} xlsx {xv:.3f} matches neither construction "
              f"(profit-avg {ratio_of_avg.loc[disc_q]:.3f}, "
              f"ratio-avg {avg_of_ratio.loc[disc_q]:.3f})")

# ---------------------------------------------------------------------------
# 6. Pre-1952 equities are annual (Q4-only) — raw FRED and xlsx handling
# ---------------------------------------------------------------------------
pre52 = raw["EQ_ALLCORP_MM"].dropna()
pre52 = pre52[pre52.index < "1952-01-01"]
non_q4 = [d for d in pre52.index if d.month != 10]
if len(pre52) == 0:
    check("FAIL", "pre-1952 raw frequency", "no pre-1952 equity observations")
elif non_q4:
    check("FAIL", "pre-1952 raw frequency",
          f"non-Q4 pre-1952 equity observations exist: "
          f"{[d.date() for d in non_q4]}")
else:
    check("PASS", "pre-1952 raw frequency",
          f"{len(pre52)} pre-1952 equity obs, all Q4 (Oct 1): "
          f"{pre52.index[0].date()}..{pre52.index[-1].date()}")

pre52_xl = xl[(xl.index >= "1947-01-01") & (xl.index < "1952-01-01")]
q4_rows = pre52_xl[pre52_xl.index.month == 10]
nq4_rows = pre52_xl[pre52_xl.index.month != 10]
q4_ok = q4_rows["PE_ALLCORP_CPATAX"].notna().all() and len(q4_rows) == 5
nq4_ok = nq4_rows["PE_ALLCORP_CPATAX"].isna().all() and len(nq4_rows) == 15
if q4_ok and nq4_ok:
    check("PASS", "pre-1952 xlsx gap handling",
          "1947-1951: ratio present on all 5 Q4 rows, NaN on all 15 non-Q4 "
          "rows (annual-frequency gaps preserved, as documented)")
else:
    check("FAIL", "pre-1952 xlsx gap handling",
          f"Q4 rows with ratio: {q4_rows['PE_ALLCORP_CPATAX'].notna().sum()}/"
          f"{len(q4_rows)} (want 5/5); non-Q4 NaN: "
          f"{nq4_rows['PE_ALLCORP_CPATAX'].isna().sum()}/{len(nq4_rows)} "
          f"(want 15/15)")

# ---------------------------------------------------------------------------
# 7. Historical benchmark: headline 4q trough ~5.96x at 1978Q4
# ---------------------------------------------------------------------------
xl_4q = xl["PE_ALLCORP_CPATAX_4Q"].dropna()
trough_date, trough_val = xl_4q.idxmin(), xl_4q.min()
my_4q = ratio_of_avg.dropna()
my_trough_date, my_trough_val = my_4q.idxmin(), my_4q.min()
EXP_TROUGH = 5.96
rel = abs(trough_val - EXP_TROUGH) / EXP_TROUGH
if trough_date == pd.Timestamp("1978-10-01") and rel <= 0.02:
    check("PASS", "1978Q4 trough benchmark",
          f"xlsx min {trough_val:.3f}x at {trough_date.date()} "
          f"(expected ~{EXP_TROUGH}x; independent re-derivation "
          f"{my_trough_val:.3f}x at {my_trough_date.date()})")
elif pd.Timestamp("1974-01-01") <= trough_date <= pd.Timestamp("1982-12-31") \
        and rel <= 0.05:
    check("WARN", "1978Q4 trough benchmark",
          f"xlsx min {trough_val:.3f}x at {trough_date.date()} — near but not "
          f"exactly the expected 5.96x @ 1978Q4 (possible data revision)")
else:
    check("FAIL", "1978Q4 trough benchmark",
          f"xlsx min {trough_val:.3f}x at {trough_date.date()}, expected "
          f"~{EXP_TROUGH}x at 1978-10-01 "
          f"(independent: {my_trough_val:.3f}x at {my_trough_date.date()})")

# ---------------------------------------------------------------------------
# 8. Data-range checks
# ---------------------------------------------------------------------------
first_date = data["Date"].iloc[0]
if first_date == pd.Timestamp("1947-01-01"):
    check("PASS", "start date", "Data sheet starts 1947-01-01 "
          "(documented intentional exception to the 2006 convention)")
else:
    check("FAIL", "start date", f"Data sheet starts {first_date.date()}, "
          "expected 1947-01-01")

latest_ratio_date = xl["PE_ALLCORP_CPATAX"].dropna().index[-1]
age_days = (TODAY - latest_ratio_date).days
if age_days <= 200:
    check("PASS", "staleness", f"latest ratio quarter {latest_ratio_date.date()} "
          f"is {age_days} days old (<= 200; Z.1/NIPA publication lag)")
else:
    check("FAIL", "staleness", f"latest ratio quarter {latest_ratio_date.date()} "
          f"is {age_days} days old (> 200) — output looks stale")

# ---------------------------------------------------------------------------
# 9. Internal consistency: dates monotonic, unique, quarterly
# ---------------------------------------------------------------------------
dates = data["Date"]
diffs_ok = dates.is_monotonic_increasing
dupes = dates.duplicated().sum()
qstart = (dates.dt.is_quarter_start).all()
gaps = pd.PeriodIndex(dates, freq="Q")
full = pd.period_range(gaps[0], gaps[-1], freq="Q")
missing_q = len(full) - len(gaps)
if diffs_ok and dupes == 0 and qstart and missing_q == 0:
    check("PASS", "date axis", f"{len(dates)} rows, strictly quarterly "
          f"{dates.iloc[0].date()}..{dates.iloc[-1].date()}, no dupes/gaps")
else:
    check("FAIL", "date axis",
          f"monotonic={diffs_ok}, dupes={dupes}, all quarter-start={qstart}, "
          f"missing quarters={missing_q}")

# ---------------------------------------------------------------------------
# 10. Internal consistency: no NaN runs mid-series (post-1952)
# ---------------------------------------------------------------------------
holes = []
for col in RATIO_COLS:
    s = xl.loc[xl.index >= "1952-01-01", col]
    valid = s.dropna()
    if valid.empty:
        holes.append(f"{col}: entirely NaN post-1952")
        continue
    inner = s.loc[valid.index[0]:valid.index[-1]]
    n_holes = inner.isna().sum()
    if n_holes:
        holes.append(f"{col}: {n_holes} interior NaN(s)")
if holes:
    check("FAIL", "interior NaNs", "; ".join(holes))
else:
    check("PASS", "interior NaNs",
          "no NaNs inside any ratio series post-1952 (trailing publication-lag "
          "NaNs allowed)")

# ---------------------------------------------------------------------------
# 11. Internal consistency: no unit-jump discontinuities
# ---------------------------------------------------------------------------
jump_fails = []
worst = ("", 0.0)
for col in RATIO_COLS + ["EQ_ALLCORP", "EQ_NCB"]:
    s = xl.loc[xl.index >= "1952-01-01", col].dropna()
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
          f"max |q/q log change| across all series {worst[1]:.2f} ({worst[0]}) "
          "< 1.0 — no 1000x unit breaks")

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
          "min/max/median/current/pctile for all 6 ratios match Data sheet")

# ---------------------------------------------------------------------------
# 13. EXTERNAL cross-check: NCB equities / GDP (Buffett-adjacent)
# ---------------------------------------------------------------------------
gdp = fred.get_series("GDP", observation_start="2020-01-01").dropna()
ncb_b = eq_ncb.dropna()
common = gdp.index.intersection(ncb_b.index)
if len(common) == 0:
    check("WARN", "external NCB-equities/GDP", "no common quarter with GDP")
else:
    q = common[-1]
    ratio = ncb_b.loc[q] / gdp.loc[q]
    if 2.5 <= ratio <= 3.5:
        check("PASS", "external NCB-equities/GDP",
              f"{ratio:.2f}x at {q.date()} (NCB equities "
              f"${ncb_b.loc[q]/1000:.1f}T / GDP ${gdp.loc[q]/1000:.1f}T) — "
              "within plausible 2.5-3.5x band, confirms equity units vs an "
              "independent series")
    else:
        check("WARN", "external NCB-equities/GDP",
              f"{ratio:.2f}x at {q.date()} outside the expected 2.5-3.5x band "
              f"(NCB ${ncb_b.loc[q]/1000:.1f}T, GDP ${gdp.loc[q]/1000:.1f}T)")

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
n_pass = RESULTS.count("PASS")
n_warn = RESULTS.count("WARN")
n_fail = RESULTS.count("FAIL")
print(f"\nSummary: {n_pass} PASS, {n_warn} WARN, {n_fail} FAIL "
      f"({len(RESULTS)} checks)")
sys.exit(1 if n_fail else 0)
