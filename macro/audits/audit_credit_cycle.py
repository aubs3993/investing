# Independent audit of macro/credit_cycle_pull.py and macro/output/credit_cycle/.
#
# Re-pulls the raw FRED series itself (BAA, AAA, BOGZ1FA154104005Q, HNOCCLQ027S,
# DPI, DRTSCILM) and re-downloads the Fed EBP csv, then re-derives the Baa-Aaa
# spread, the household borrowing rate B_HH, the consumer-credit impulse CI_CC,
# the z-scores and the STRESS/FROTH composites with its own code. Does NOT
# import or exec credit_cycle_pull.py; reads credit_cycle.xlsx (and the sibling
# consumer_pullforward.xlsx for the shared-bridge cross-check) only to COMPARE.
# The analysis source file is read as TEXT ONLY (never executed) to verify the
# froth chart's 24m forward shift / legend wording and the ffill(limit=8) guard,
# since neither is distinguishable from the stored data alone.
#
# Tolerances:
#   - Same-vintage FRED re-derivations (spread, B_HH, CI_CC, z, composites):
#     0.1% relative / 1e-6 absolute where noted. Audit runs against the same
#     same-day vintage, so only xlsx float round-trip (~1e-15) is expected;
#     0.1% leaves headroom for an intraday revision without masking a formula
#     or unit error (Mil$/Bil$ mix-ups are 1000x, wrong-denominator errors >5%).
#   - EBP: the Fed RE-ESTIMATES the full history every monthly update, so a
#     fresh download may legitimately differ from the archived run — that is a
#     WARN (after confirming xlsx == archived copy), not a FAIL.
#   - Historical anchors (latest spread ~0.48pp, max 5.64pp @ 1932-05, B_HH
#     13.9% @ 2006Q2 / -2.1% @ 2009Q3 / ~3.1% latest, EBP +3.41 @ 2008-10 /
#     -0.39 latest): looser bands, since these came from a prior session and
#     source data can revise.
#
# Prints one line per check: PASS|FAIL|WARN <name>: <detail>. Exits 1 on FAIL.

from datetime import datetime
from io import StringIO
from pathlib import Path
import re
import sys

import numpy as np
import pandas as pd
import requests

HERE = Path(__file__).resolve().parent      # macro/audits
MACRO_DIR = HERE.parent                     # macro/
REPO_ROOT = MACRO_DIR.parent
sys.path.insert(0, str(REPO_ROOT))

from shared.fred_helpers import get_fred_client  # noqa: E402

OUT_DIR = MACRO_DIR / "output" / "credit_cycle"
XLSX = OUT_DIR / "credit_cycle.xlsx"
EBP_ARCHIVE = OUT_DIR / "ebp_csv.csv"
CP_XLSX = MACRO_DIR / "output" / "consumer_pullforward" / "consumer_pullforward.xlsx"
SCRIPT = MACRO_DIR / "credit_cycle_pull.py"
EBP_URL = "https://www.federalreserve.gov/econres/notes/feds-notes/ebp_csv.csv"

TODAY = pd.Timestamp(datetime.today().date())

RESULTS = []


def check(status: str, name: str, detail: str) -> None:
    RESULTS.append(status)
    print(f"{status} {name}: {detail}")


# ---------------------------------------------------------------------------
# 1. Output file existence + non-trivial size
# ---------------------------------------------------------------------------
EXPECTED_FILES = {
    "credit_cycle.xlsx": 100_000,
    "ebp_csv.csv": 20_000,
    "credit_price_panel.png": 50_000,
    "credit_quantity_panel.png": 50_000,
    "credit_standards_performance.png": 50_000,
    "credit_cycle_composite.png": 50_000,
    "credit_composite_2006.png": 50_000,
    "credit_vs_spx.png": 50_000,
    "credit_vs_rates.png": 50_000,
    "credit_impulse_household.png": 50_000,
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
    print("\nSummary: cannot continue without credit_cycle.xlsx")
    sys.exit(1)

dm = pd.read_excel(XLSX, sheet_name="Data_Monthly")
dm["Date"] = pd.to_datetime(dm["Date"])
dm = dm.set_index("Date")
dq = pd.read_excel(XLSX, sheet_name="Data_Quarterly")
dq["Date"] = pd.to_datetime(dq["Date"])
dq = dq.set_index("Date")

# ---------------------------------------------------------------------------
# 2. Independent raw pulls from FRED
# ---------------------------------------------------------------------------
fred = get_fred_client()
baa = fred.get_series("BAA", observation_start="1900-01-01")
aaa = fred.get_series("AAA", observation_start="1900-01-01")
flow_hh_mm = fred.get_series("BOGZ1FA154104005Q", observation_start="1946-01-01")
flow_cc_mm = fred.get_series("HNOCCLQ027S", observation_start="1946-01-01")
dpi = fred.get_series("DPI", observation_start="1946-01-01")
drtscilm = fred.get_series("DRTSCILM", observation_start="1946-01-01")
revolsl = fred.get_series("REVOLSL", observation_start="1900-01-01")
busloans = fred.get_series("BUSLOANS", observation_start="1900-01-01")

# ---------------------------------------------------------------------------
# 3. Re-derivation: Baa-Aaa spread (full overlap + anchors)
# ---------------------------------------------------------------------------
my_spread = (baa - aaa).dropna()
common = my_spread.index.intersection(dm.index)
if len(common) < 1200:
    check("FAIL", "spread coverage",
          f"only {len(common)} overlapping months between my BAA-AAA pull and "
          f"Data_Monthly (expected ~1290)")
else:
    diff = (dm.loc[common, "SPREAD"] - my_spread.loc[common]).abs()
    if diff.max() <= 0.005:
        check("PASS", "spread re-derivation",
              f"BAA-AAA matches xlsx SPREAD on all {len(common)} months "
              f"(max abs diff {diff.max():.2e}pp)")
    else:
        check("FAIL", "spread re-derivation",
              f"max abs diff {diff.max():.3f}pp at {diff.idxmax().date()} "
              f"(mine {my_spread.loc[diff.idxmax()]:.2f}, "
              f"xlsx {dm.loc[diff.idxmax(), 'SPREAD']:.2f})")

sp_latest = my_spread.iloc[-1]
xl_sp_latest = dm["SPREAD"].dropna().iloc[-1]
if abs(sp_latest - xl_sp_latest) <= 0.005 and abs(sp_latest - 0.48) <= 0.15:
    check("PASS", "spread latest anchor",
          f"mine {sp_latest:.2f}pp ({my_spread.index[-1]:%Y-%m}) = xlsx "
          f"{xl_sp_latest:.2f}pp, near expected ~0.48pp")
elif abs(sp_latest - xl_sp_latest) <= 0.005:
    check("WARN", "spread latest anchor",
          f"mine {sp_latest:.2f}pp matches xlsx but is {sp_latest - 0.48:+.2f}pp "
          f"off the ~0.48pp anchor (new month or revision)")
else:
    check("FAIL", "spread latest anchor",
          f"mine {sp_latest:.2f}pp vs xlsx {xl_sp_latest:.2f}pp")

sp_max_date, sp_max = my_spread.idxmax(), my_spread.max()
xl_max_date, xl_max = dm["SPREAD"].idxmax(), dm["SPREAD"].max()
if (sp_max_date == pd.Timestamp("1932-05-01") and abs(sp_max - 5.64) <= 0.05
        and xl_max_date == sp_max_date and abs(xl_max - sp_max) <= 0.005):
    check("PASS", "spread max anchor (external benchmark)",
          f"max {sp_max:.2f}pp at {sp_max_date:%Y-%m} in both my pull and xlsx "
          f"— matches the known Great-Depression peak 5.64pp 1932-05")
else:
    check("FAIL", "spread max anchor (external benchmark)",
          f"mine max {sp_max:.2f}pp @ {sp_max_date:%Y-%m}, xlsx {xl_max:.2f}pp "
          f"@ {xl_max_date:%Y-%m}, expected 5.64pp @ 1932-05")

# ---------------------------------------------------------------------------
# 4. EXTERNAL cross-check: fresh Fed EBP download vs Data_Monthly (+ archive)
# ---------------------------------------------------------------------------
try:
    resp = requests.get(EBP_URL, timeout=120,
                        headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    ebp_fresh = pd.read_csv(StringIO(resp.text))
    ebp_fresh["Date"] = pd.to_datetime(ebp_fresh["date"])
    ebp_fresh = ebp_fresh.set_index("Date")["ebp"]
except Exception as exc:  # network failure should not kill the whole audit
    ebp_fresh = None
    check("WARN", "EBP fresh download", f"could not fetch {EBP_URL}: {exc}")

ebp_arch = pd.read_csv(EBP_ARCHIVE)
ebp_arch["Date"] = pd.to_datetime(ebp_arch["date"])
ebp_arch = ebp_arch.set_index("Date")["ebp"]

# 4a. xlsx must equal the archived copy it was built from (exact)
com_a = ebp_arch.index.intersection(dm.index)
d_arch = (dm.loc[com_a, "ebp"] - ebp_arch.loc[com_a]).abs().max()
if d_arch <= 1e-9:
    check("PASS", "EBP xlsx == archived csv",
          f"Data_Monthly ebp identical to archived ebp_csv.csv on all "
          f"{len(com_a)} months (max diff {d_arch:.1e})")
else:
    check("FAIL", "EBP xlsx == archived csv",
          f"max abs diff {d_arch:.4f} — xlsx was not built from the archived "
          f"copy in the output folder")

# 4b. fresh download vs xlsx (WARN on revision, per header note)
if ebp_fresh is not None:
    com_f = ebp_fresh.index.intersection(dm.index)
    d_fresh = (dm.loc[com_f, "ebp"] - ebp_fresh.loc[com_f]).abs().max()
    if d_fresh <= 0.02:
        check("PASS", "EBP fresh vs xlsx",
              f"fresh Fed download matches Data_Monthly ebp "
              f"(max abs diff {d_fresh:.2e} over {len(com_f)} months)")
    elif d_arch <= 1e-9:
        check("WARN", "EBP fresh vs xlsx",
              f"fresh download differs from xlsx by up to {d_fresh:.3f} but "
              f"xlsx == archived copy — Fed monthly full-history re-estimate "
              f"since the run, not a script error")
    else:
        check("FAIL", "EBP fresh vs xlsx",
              f"fresh differs by {d_fresh:.3f} AND xlsx differs from archive")

    src = ebp_fresh
else:
    src = ebp_arch

# 4c. anchors: 2008-10 ~ +3.41, latest ~ -0.39 (EBP revises -> WARN band)
e_gfc = src.get(pd.Timestamp("2008-10-01"), np.nan)
e_last = src.dropna().iloc[-1]
if abs(e_gfc - 3.41) <= 0.15 and abs(e_last - (-0.39)) <= 0.15:
    check("PASS", "EBP anchors",
          f"2008-10 EBP {e_gfc:+.2f} (~+3.41), latest {e_last:+.2f} "
          f"({src.dropna().index[-1]:%Y-%m}, ~-0.39)")
elif abs(e_gfc - 3.41) <= 0.5 and abs(e_last - (-0.39)) <= 0.5:
    check("WARN", "EBP anchors",
          f"2008-10 {e_gfc:+.2f} / latest {e_last:+.2f} within 0.5 of the "
          f"+3.41 / -0.39 anchors — consistent with a Fed re-estimate")
else:
    check("FAIL", "EBP anchors",
          f"2008-10 {e_gfc:+.2f} (expected ~+3.41), latest {e_last:+.2f} "
          f"(expected ~-0.39)")

# ---------------------------------------------------------------------------
# 5. Re-derivation: B_HH household borrowing rate (own 4q means)
# ---------------------------------------------------------------------------
# Own construction on a regular quarterly grid (time-aware, min_periods=4), so
# a positional-window bug in the analysis would surface as a mismatch.
qidx = pd.date_range("1946-01-01", max(flow_hh_mm.index.max(), dpi.index.max()),
                     freq="QS")
f_hh = flow_hh_mm.reindex(qidx) / 1000.0          # Mil$ SAAR -> Bil$ SAAR
f_cc = flow_cc_mm.reindex(qidx) / 1000.0
dpi_q = dpi.reindex(qidx)
f_hh_4q = f_hh.rolling(4, min_periods=4).mean()
f_cc_4q = f_cc.rolling(4, min_periods=4).mean()
dpi_4q = dpi_q.rolling(4, min_periods=4).mean()
my_bhh = 100 * f_hh_4q / dpi_4q
my_cicc = 100 * (f_cc_4q - f_cc_4q.shift(4)) / dpi_4q

BHH_SPOTS = [
    ("2006Q2 peak", pd.Timestamp("2006-04-01"), 13.9, 0.3),
    ("2009Q3 trough", pd.Timestamp("2009-07-01"), -2.1, 0.3),
]
for label, spot, anchor, band in BHH_SPOTS:
    mv = my_bhh.get(spot, np.nan)
    xv = dq.loc[spot, "B_HH"] if spot in dq.index else np.nan
    if pd.isna(mv) or pd.isna(xv):
        check("FAIL", f"B_HH {label}", f"missing value (mine={mv}, xlsx={xv})")
    elif abs(mv - xv) > 0.02:
        check("FAIL", f"B_HH {label}",
              f"mine {mv:.2f}% vs xlsx {xv:.2f}% — re-derivation mismatch "
              f"(unit or window error)")
    elif abs(mv - anchor) > band:
        check("WARN", f"B_HH {label}",
              f"mine {mv:.2f}% matches xlsx but is off the {anchor}% anchor "
              f"(Z.1 revision?)")
    else:
        check("PASS", f"B_HH {label}",
              f"mine {mv:.2f}% = xlsx {xv:.2f}% (anchor ~{anchor}%)")

bhh_mine_last = my_bhh.dropna().iloc[-1]
bhh_xl_last = dq["B_HH"].dropna().iloc[-1]
bhh_last_q = my_bhh.dropna().index[-1]
if abs(bhh_mine_last - bhh_xl_last) <= 0.02 and abs(bhh_mine_last - 3.1) <= 0.5:
    check("PASS", "B_HH latest",
          f"mine {bhh_mine_last:.2f}% ({bhh_last_q:%Y}Q{bhh_last_q.quarter}) "
          f"= xlsx {bhh_xl_last:.2f}% (anchor ~3.1%)")
elif abs(bhh_mine_last - bhh_xl_last) <= 0.02:
    check("WARN", "B_HH latest",
          f"mine {bhh_mine_last:.2f}% matches xlsx but off the ~3.1% anchor")
else:
    check("FAIL", "B_HH latest",
          f"mine {bhh_mine_last:.2f}% vs xlsx {bhh_xl_last:.2f}%")

# ---------------------------------------------------------------------------
# 6. Re-derivation: CI_CC at 2 spots + bridge cross-check vs
#    consumer_pullforward.xlsx (identical formula in two scripts)
# ---------------------------------------------------------------------------
for label, spot in [("2006Q2", pd.Timestamp("2006-04-01")),
                    ("2009Q3", pd.Timestamp("2009-07-01"))]:
    mv = my_cicc.get(spot, np.nan)
    xv = dq.loc[spot, "CI_CC"] if spot in dq.index else np.nan
    if pd.isna(mv) or pd.isna(xv):
        check("FAIL", f"CI_CC {label}", f"missing (mine={mv}, xlsx={xv})")
    elif abs(mv - xv) <= 0.02:
        check("PASS", f"CI_CC {label}",
              f"mine {mv:+.3f}% of DPI = xlsx {xv:+.3f}%")
    else:
        check("FAIL", f"CI_CC {label}",
              f"mine {mv:+.3f}% vs xlsx {xv:+.3f}% — re-derivation mismatch")

if not CP_XLSX.exists():
    check("WARN", "CI_CC bridge cross-check",
          f"{CP_XLSX} not found — cannot verify the shared bridge series")
else:
    cp = pd.read_excel(CP_XLSX, sheet_name="Data_Quarterly")
    cp["Date"] = pd.to_datetime(cp["Date"])
    cp = cp.set_index("Date")
    both = dq["CI_CC"].dropna().index.intersection(cp["CI_CC"].dropna().index)
    # consumer_pullforward is a 2006-focused module, so its quarterly frame
    # only starts 2006Q1 (~81 quarters of overlap); credit_cycle's CI_CC runs
    # from 1953Q3. Overlap of ~80 quarters is the expected full intersection.
    if len(both) < 60:
        check("FAIL", "CI_CC bridge cross-check",
              f"only {len(both)} overlapping quarters with CI_CC in both files")
    else:
        d = (dq.loc[both, "CI_CC"] - cp.loc[both, "CI_CC"]).abs()
        if d.max() <= 1e-6:
            check("PASS", "CI_CC bridge cross-check",
                  f"credit_cycle and consumer_pullforward CI_CC identical on "
                  f"all {len(both)} common quarters (max diff {d.max():.1e})")
        else:
            # diagnose which side is wrong using my independent re-derivation
            w = d.idxmax()
            mine_w = my_cicc.get(w, np.nan)
            cc_err = abs(dq.loc[w, "CI_CC"] - mine_w)
            cp_err = abs(cp.loc[w, "CI_CC"] - mine_w)
            culprit = ("consumer_pullforward_pull.py" if cp_err > cc_err
                       else "credit_cycle_pull.py")
            check("FAIL", "CI_CC bridge cross-check",
                  f"max diff {d.max():.4f} at {w.date()} "
                  f"(credit_cycle {dq.loc[w, 'CI_CC']:+.4f}, "
                  f"consumer_pullforward {cp.loc[w, 'CI_CC']:+.4f}, "
                  f"independent {mine_w:+.4f}) — {culprit} looks wrong")

# ---------------------------------------------------------------------------
# 7. Composites: re-derive z-scores (full-sample, ddof=1, clip +/-3) from the
#    stored raw columns, then STRESS/FROTH as equal-weight means over
#    non-missing legs; verify the clip actually binds
# ---------------------------------------------------------------------------
Z_MAP = {
    "z_spread": "SPREAD",
    "z_spread_chg": "SPREAD_CHG",
    "z_ebp": "ebp",
    "z_sloos": "DRTSCILM_M",
    "z_impulse": "IMPULSE_AGG_M",
    "z_ci_hh": "CI_HH_M",
}
my_z = {}
z_bad = []
for zcol, raw_col in Z_MAP.items():
    s = dm[raw_col]
    z = ((s - s.mean()) / s.std()).clip(-3, 3)
    my_z[zcol] = z
    d = (dm[zcol] - z).abs().max()
    if not (d <= 1e-9):
        z_bad.append(f"{zcol}: max diff {d:.2e}")
if z_bad:
    check("FAIL", "z-score re-derivation", "; ".join(z_bad))
else:
    check("PASS", "z-score re-derivation",
          "all 6 z columns = full-sample (x-mean)/std clipped to +/-3 "
          "(max abs diff < 1e-9)")

n_clip = sum(int((dm[z].abs() == 3.0).sum()) for z in Z_MAP)
raw_z_spread_32 = ((dm["SPREAD"] - dm["SPREAD"].mean()) / dm["SPREAD"].std()
                   ).loc["1932-05-01"]
if n_clip > 0 and raw_z_spread_32 > 3:
    check("PASS", "winsorization binds",
          f"{n_clip} stored z values sit exactly at +/-3; unclipped z_spread "
          f"at 1932-05 is {raw_z_spread_32:.2f} > 3 — the clip is real")
else:
    check("FAIL", "winsorization binds",
          f"clipped-count {n_clip}, unclipped 1932-05 z {raw_z_spread_32:.2f} "
          f"— +/-3 winsorization not in effect")

my_stress = pd.concat([my_z[c] for c in
                       ["z_spread", "z_spread_chg", "z_ebp", "z_sloos"]],
                      axis=1).mean(axis=1)
my_froth = pd.concat([-my_z["z_spread"], my_z["z_impulse"],
                      -my_z["z_sloos"], my_z["z_ci_hh"]], axis=1).mean(axis=1)
d_stress = (dm["STRESS"] - my_stress).abs().max()
d_froth = (dm["FROTH"] - my_froth).abs().max()
if d_stress <= 1e-9 and d_froth <= 1e-9:
    check("PASS", "composite construction",
          f"STRESS = mean(z_spread, z_spread_chg, z_ebp, z_sloos) over "
          f"non-missing legs and FROTH = mean(-z_spread, z_impulse, -z_sloos, "
          f"z_ci_hh) on every month (max diffs {d_stress:.1e}/{d_froth:.1e})")
else:
    check("FAIL", "composite construction",
          f"max abs diff STRESS {d_stress:.2e}, FROTH {d_froth:.2e}")

s_gfc = dm.loc["2008-12-01", "STRESS"]
zrow = dm.loc["2008-12-01", ["z_spread", "z_spread_chg", "z_ebp", "z_sloos"]]
if abs(s_gfc - 3.0) <= 1e-9 and (zrow == 3.0).all():
    check("PASS", "STRESS 2008-12 anchor",
          "STRESS = +3.00 with all four legs clipped at +3 (as expected at "
          "the GFC peak)")
else:
    check("FAIL", "STRESS 2008-12 anchor",
          f"STRESS {s_gfc:+.3f}, legs {zrow.tolist()} — expected +3.00 with "
          f"all legs at the +3 clip")

calm = pd.Timestamp("2019-06-01")
s_calm, m_calm = dm.loc[calm, "STRESS"], my_stress.loc[calm]
if abs(s_calm - m_calm) <= 1e-9 and abs(s_calm) < 1.0:
    check("PASS", "STRESS calm-month spot",
          f"2019-06 STRESS {s_calm:+.3f} matches re-derivation and sits in "
          f"the calm band (|z| < 1)")
else:
    check("FAIL", "STRESS calm-month spot",
          f"2019-06 xlsx {s_calm:+.3f} vs re-derived {m_calm:+.3f}")

# ---------------------------------------------------------------------------
# 8. Froth line advanced 24 months on the composite chart
# ---------------------------------------------------------------------------
# The shift exists only at chart time, so it cannot be verified from the xlsx
# alone. Verify (a) the xlsx FROTH column supports the shift (contiguous back
# to its start, so froth(t-24) exists for every plotted t), and (b) the source
# TEXT (never executed) applies +24 months to the froth index and labels the
# line as advanced.
src_text = SCRIPT.read_text(encoding="utf-8")
froth_v = dm["FROTH"].dropna()
froth_contig = froth_v.index.equals(
    pd.date_range(froth_v.index[0], froth_v.index[-1], freq="MS"))
shift_ok = re.search(
    r"froth_shifted\.index\s*=\s*froth_shifted\.index\s*\+\s*"
    r"pd\.DateOffset\(months=24\)", src_text)
legend_ok = re.search(r"FROTH composite, advanced 24m", src_text)
title_ok = re.search(r"advanced 24m", src_text)
if froth_contig and shift_ok and legend_ok and title_ok:
    check("PASS", "froth 24m forward shift",
          "FROTH column contiguous (value plotted at t reproducible as "
          "froth(t-24) from the xlsx); source shifts froth_shifted.index by "
          "+DateOffset(months=24) and both legend and title say 'advanced 24m'")
elif froth_contig and shift_ok:
    check("WARN", "froth 24m forward shift",
          "shift is +24 months in source but legend/title wording not found — "
          "chart may not disclose the shift")
else:
    check("FAIL", "froth 24m forward shift",
          f"contiguous={froth_contig}, source-shift-found={bool(shift_ok)} — "
          f"cannot confirm the plotted froth line equals froth(t-24)")

# ---------------------------------------------------------------------------
# 9. DRTSCILM quarterly -> monthly ffill with 8-month limit
# ---------------------------------------------------------------------------
midx = dm.index
my_sloos_lim = dq["DRTSCILM"].reindex(midx).ffill(limit=8)
my_sloos_unlim = dq["DRTSCILM"].reindex(midx).ffill()
d_lim = (dm["DRTSCILM_M"] - my_sloos_lim).abs().max()
lim_matches = d_lim <= 1e-9 or pd.isna(d_lim)
same_cols = my_sloos_lim.equals(my_sloos_unlim)
src_limit = re.search(r"ffill\(limit=8\)", src_text)
# independent quarterly pull: confirm no gap > 8 months hides in DRTSCILM
gaps = drtscilm.dropna().index.to_series().diff().dt.days.max()
if lim_matches and src_limit:
    detail = (f"Data_Monthly DRTSCILM_M == quarterly DRTSCILM reindexed+"
              f"ffill(limit=8) (max diff {0.0 if pd.isna(d_lim) else d_lim:.1e})"
              f"; source uses ffill(limit=8)")
    if same_cols:
        detail += (f" — note: limited and unlimited ffill coincide on current "
                   f"data (max quarterly gap {gaps:.0f} days), so the limit is "
                   f"verified from source text, not data")
    check("PASS", "DRTSCILM ffill 8-month limit", detail)
elif lim_matches:
    check("WARN", "DRTSCILM ffill 8-month limit",
          "monthly column matches limit=8 construction but ffill(limit=8) not "
          "found in source — cannot confirm the guard exists")
else:
    check("FAIL", "DRTSCILM ffill 8-month limit",
          f"max abs diff {d_lim:.3f} vs reindex+ffill(limit=8) of the "
          f"quarterly column")

# ---------------------------------------------------------------------------
# 10. Date axes: Data_Monthly 1290 rows 1919-01..2026-06 monotone unique;
#     Data_Quarterly regular quarterly
# ---------------------------------------------------------------------------
exp_m = pd.date_range("1919-01-01", "2026-06-01", freq="MS")
if len(dm) == 1290 and dm.index.equals(exp_m):
    check("PASS", "Data_Monthly date axis",
          "1290 rows, exactly monthly-start 1919-01..2026-06, monotone, "
          "unique, no gaps")
elif (dm.index.is_monotonic_increasing and dm.index.is_unique
      and dm.index[0] == exp_m[0]
      and dm.index.equals(pd.date_range(dm.index[0], dm.index[-1], freq="MS"))):
    check("WARN", "Data_Monthly date axis",
          f"{len(dm)} rows {dm.index[0]:%Y-%m}..{dm.index[-1]:%Y-%m} — regular "
          f"monthly grid but end differs from the expected 2026-06 (new data?)")
else:
    check("FAIL", "Data_Monthly date axis",
          f"{len(dm)} rows {dm.index[0]:%Y-%m}..{dm.index[-1]:%Y-%m}, "
          f"monotone={dm.index.is_monotonic_increasing}, "
          f"unique={dm.index.is_unique}")

qexp = pd.date_range(dq.index[0], dq.index[-1], freq="QS")
if dq.index.is_monotonic_increasing and dq.index.is_unique \
        and dq.index.equals(qexp):
    check("PASS", "Data_Quarterly date axis",
          f"{len(dq)} rows, regular quarter-start grid "
          f"{dq.index[0]:%Y-%m}..{dq.index[-1]:%Y-%m}")
else:
    check("FAIL", "Data_Quarterly date axis",
          f"{len(dq)} rows, monotone={dq.index.is_monotonic_increasing}, "
          f"unique={dq.index.is_unique}, regular={dq.index.equals(qexp)}")

# ---------------------------------------------------------------------------
# 11. Internal consistency: interior NaNs and unit jumps
# ---------------------------------------------------------------------------
holes = []
for col in ["SPREAD", "SPREAD_CHG", "STRESS", "FROTH", "ebp", "DRTSCILM_M"]:
    s = dm[col]
    valid = s.dropna()
    if valid.empty:
        holes.append(f"{col}: entirely NaN")
        continue
    inner = s.loc[valid.index[0]:valid.index[-1]]
    n = int(inner.isna().sum())
    if n:
        holes.append(f"{col}: {n} interior NaN(s)")
if holes:
    check("FAIL", "interior NaNs (monthly)", "; ".join(holes))
else:
    check("PASS", "interior NaNs (monthly)",
          "SPREAD/SPREAD_CHG/STRESS/FROTH/ebp/DRTSCILM_M have no NaNs between "
          "their first and last valid months (trailing publication-lag NaNs "
          "allowed)")

# A large step is only a script defect if it is NOT in the raw source series
# (known raw breaks exist: REVOLSL Jan-1977 coverage redefinition ~+86% m/m;
# the Baa-Aaa spread fell 5.53 -> 3.31pp in Aug-1932). So each xlsx level
# column is compared against my own raw pull; any jump present in both is
# attributed to the source data, and only an xlsx-not-in-raw jump FAILs.
jumps = []
attributed = []
for col, raw_s in [("BUSLOANS", busloans), ("REVOLSL", revolsl)]:
    s = dm[col].dropna()
    com = s.index.intersection(raw_s.index)
    fid = (s.loc[com] - raw_s.loc[com]).abs().max()
    if fid > max(1e-6, 1e-9 * raw_s.abs().max()):
        jumps.append(f"{col}: xlsx differs from raw FRED pull by {fid:.3g}")
        continue
    lg = np.abs(np.log(s / s.shift(1))).dropna()
    big = lg[lg >= 0.5]
    for d, v in big.items():
        attributed.append(f"{col} {v:.2f} log-chg @ {d:%Y-%m} (in raw data)")
sp_step = dm["SPREAD"].diff().abs().dropna()
# SPREAD == raw BAA-AAA everywhere (check 3), so any big move is source data.
if sp_step.max() >= 2.0:
    attributed.append(f"SPREAD {sp_step.max():.2f}pp @ "
                      f"{sp_step.idxmax():%Y-%m} (in raw data)")
if jumps:
    check("FAIL", "unit discontinuities", "; ".join(jumps))
else:
    detail = ("xlsx BUSLOANS/REVOLSL match my raw FRED pulls exactly, so no "
              "transformation-induced breaks")
    if attributed:
        detail += ("; source-data steps noted: " + "; ".join(attributed)
                   + " — REVOLSL Jan-1977 is the known series redefinition, "
                     "SPREAD Aug-1932 is the genuine post-peak collapse")
    check("PASS", "unit discontinuities", detail)

flow_neg_ok = (dq["FLOW_HH"].dropna().abs().max() < 6000)   # Bil$ SAAR scale
if flow_neg_ok:
    check("PASS", "flow units (Mil->Bil)",
          f"FLOW_HH magnitudes max {dq['FLOW_HH'].dropna().abs().max():,.0f} "
          f"Bil$ SAAR — consistent with /1000 conversion (Mil$ raw would be "
          f"1000x larger)")
else:
    check("FAIL", "flow units (Mil->Bil)",
          f"FLOW_HH magnitudes up to "
          f"{dq['FLOW_HH'].dropna().abs().max():,.0f} — looks like Mil$ SAAR "
          f"was written without the /1000 conversion")

# ---------------------------------------------------------------------------
# 12. Staleness
# ---------------------------------------------------------------------------
last_sp = dm["SPREAD"].dropna().index[-1]
age_m = (TODAY - last_sp).days
if age_m <= 75:
    check("PASS", "staleness (monthly)",
          f"latest spread month {last_sp:%Y-%m} is {age_m} days old (<= 75)")
else:
    check("FAIL", "staleness (monthly)",
          f"latest spread month {last_sp:%Y-%m} is {age_m} days old (> 75)")

last_bhh = dq["B_HH"].dropna().index[-1]
age_q = (TODAY - last_bhh).days
if age_q <= 220:
    check("PASS", "staleness (quarterly)",
          f"latest B_HH quarter {last_bhh:%Y}Q{last_bhh.quarter} is {age_q} "
          f"days old (<= 220; Z.1 lags ~10 weeks)")
else:
    check("FAIL", "staleness (quarterly)",
          f"latest B_HH quarter {last_bhh:%Y}Q{last_bhh.quarter} is {age_q} "
          f"days old (> 220) — output looks stale")

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
n_pass = RESULTS.count("PASS")
n_warn = RESULTS.count("WARN")
n_fail = RESULTS.count("FAIL")
print(f"\nSummary: {n_pass} PASS, {n_warn} WARN, {n_fail} FAIL "
      f"({len(RESULTS)} checks)")
sys.exit(1 if n_fail else 0)
