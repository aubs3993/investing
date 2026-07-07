# Independent audit of macro/shiller_pe_pull.py outputs.
#
# Re-parses the cached raw ie_data.xls (sheet "Data", skiprows=7, xlrd) with
# its own code, re-derives trailing P/E / peaks / summary stats, and compares
# against macro/output/shiller_pe/shiller_pe.xlsx. Does NOT import or exec the
# analysis script. External cross-check: multpl.com monthly S&P 500 P/E.
#
# Tolerances (stated per check):
# - REL_TOL = 1e-8 for raw-vs-xlsx re-derivations: both derive from the same
#   cached ie_data.xls and xlsx stores IEEE doubles losslessly, so differences
#   should be exactly zero; 1e-8 only absorbs benign float repr noise. Anything
#   larger means the builder parsed/derived differently.
# - ±0.2 abs for known published benchmarks (peak P/E 123.7x, peak CAPE 44.2x)
#   which are quoted to 1 decimal.
# - multpl.com comparison is WARN-level with wide bands (see check) because the
#   two sources use different earnings bases by design.
#
# Run from repo root: python macro/audits/audit_shiller_pe.py
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

AUDIT_DIR = Path(__file__).resolve().parent          # macro/audits
REPO_ROOT = AUDIT_DIR.parent.parent                  # repo root
sys.path.insert(0, str(REPO_ROOT))                   # shared helpers importable
# (No FRED access needed: raw data is cached locally; FRED is only used by the
# analysis script for recession shading, which is cosmetic.)

OUT_DIR = REPO_ROOT / "macro" / "output" / "shiller_pe"
RAW_XLS = OUT_DIR / "ie_data.xls"
XLSX = OUT_DIR / "shiller_pe.xlsx"

REL_TOL = 1e-8
TODAY = datetime.today()

results = []  # (status, name, detail)


def check(status, name, detail=""):
    results.append((status, name, detail))
    print(f"{status} {name}: {detail}")


def close(a, b, rel=REL_TOL):
    return abs(a - b) <= rel * max(1.0, abs(a), abs(b))


# ---------------------------------------------------------------- file checks
expected_files = [
    (XLSX, 20_000),
    (RAW_XLS, 500_000),
    (OUT_DIR / "shiller_pe_full.png", 30_000),
    (OUT_DIR / "shiller_cape_full.png", 30_000),
    (OUT_DIR / "shiller_pe_2006.png", 30_000),
    (OUT_DIR / ".gitkeep", 0),
]
for path, min_size in expected_files:
    if not path.exists():
        check("FAIL", f"file exists {path.name}", "missing")
    elif path.stat().st_size < min_size:
        check("FAIL", f"file size {path.name}",
              f"{path.stat().st_size} bytes < required {min_size}")
    else:
        check("PASS", f"file exists {path.name}",
              f"{path.stat().st_size:,} bytes (min {min_size:,})")

if not RAW_XLS.exists() or not XLSX.exists():
    print("\nCannot continue without raw + xlsx files.")
    print(f"SUMMARY: pass=0 warn=0 fail={len([r for r in results if r[0]=='FAIL'])}")
    sys.exit(1)

# ------------------------------------------------- independent raw re-parse
raw = pd.read_excel(RAW_XLS, sheet_name="Data", skiprows=7, engine="xlrd")
needed = ["Date", "P", "E", "CPI", "CAPE", "TR CAPE"]
missing = [c for c in needed if c not in raw.columns]
if missing:
    check("FAIL", "raw column layout", f"missing {missing}; got {list(raw.columns)}")
    print("SUMMARY: layout changed, aborting")
    sys.exit(1)
check("PASS", "raw column layout", "Date, P, E, CPI, CAPE, TR CAPE all present")

my = raw[needed].copy()
for c in needed:
    my[c] = pd.to_numeric(my[c], errors="coerce")
my = my.dropna(subset=["Date"]).reset_index(drop=True)

# Independent date decode. The Date column is a float whose FRACTION is the
# literal month string: 1871.1 means ".10" = October (not January). Decode by
# formatting to exactly 2 decimals and splitting — same trap, own code.
ds = my["Date"].map(lambda d: f"{d:.2f}")
my["year"] = ds.str.split(".").str[0].astype(int)
my["month"] = ds.str.split(".").str[1].astype(int)
my["dt"] = pd.to_datetime(dict(year=my["year"], month=my["month"], day=1))
my["PE"] = my["P"] / my["E"]

# Calendar integrity of MY OWN parse (guards the float-date trap at the
# source): one row per (year, month); every full year has exactly months 1-12,
# hence exactly one October — a naive float parse would turn every October
# into a duplicate January and drop month 10 entirely.
dup = my.duplicated(subset=["year", "month"]).sum()
check("PASS" if dup == 0 else "FAIL", "raw one row per (year, month)",
      f"{dup} duplicate year/month rows in own parse")

full_years = [y for y in range(my["year"].iloc[0] + 0, my["year"].iloc[-1])
              if y > my["year"].iloc[0] - 1]
# a "full year" = any year strictly before the final (possibly partial) year
bad_years = []
for y, grp in my.groupby("year"):
    if y == my["year"].iloc[-1]:
        continue  # current year may be partial
    if sorted(grp["month"].tolist()) != list(range(1, 13)):
        bad_years.append(y)
check("PASS" if not bad_years else "FAIL",
      "raw 12 distinct months per full year (October present)",
      f"{len(bad_years)} bad years{': ' + str(bad_years[:5]) if bad_years else ''}"
      f" out of {my['year'].iloc[-1] - my['year'].iloc[0]} full years")

# ------------------------------------------------------------- load the xlsx
xl = pd.read_excel(XLSX, sheet_name="Data")
xl["Date"] = pd.to_datetime(xl["Date"])

# Row/date alignment: xlsx dates must equal my independently decoded dates
# exactly. Catches any October/float-date corruption in the builder's parse.
if len(xl) != len(my):
    check("FAIL", "xlsx row count matches raw", f"xlsx {len(xl)} vs raw {len(my)}")
else:
    n_mismatch = int((xl["Date"].values != my["dt"].values).sum())
    check("PASS" if n_mismatch == 0 else "FAIL", "xlsx dates match own decode",
          f"{len(xl)} rows, {n_mismatch} date mismatches")

# xlsx calendar integrity: one row per month, 12 months every full year.
xdup = xl["Date"].duplicated().sum()
check("PASS" if xdup == 0 else "FAIL", "xlsx no duplicate dates",
      f"{xdup} duplicates")
mono = xl["Date"].is_monotonic_increasing
gaps = (xl["Date"].dt.to_period("M").astype(int).diff().dropna() != 1).sum()
check("PASS" if mono and gaps == 0 else "FAIL",
      "xlsx dates monotonic, exactly 1-month steps",
      f"monotonic={mono}, non-1-month steps={int(gaps)}")
oct_missing = []
for y, grp in xl.groupby(xl["Date"].dt.year):
    if y == xl["Date"].dt.year.iloc[-1]:
        continue
    months = sorted(grp["Date"].dt.month.tolist())
    if months != list(range(1, 13)):
        oct_missing.append(int(y))
check("PASS" if not oct_missing else "FAIL",
      "xlsx 12 months per full year (float-date trap)",
      f"{len(oct_missing)} incomplete full years"
      f"{': ' + str(oct_missing[:5]) if oct_missing else ''}")

# -------------------------------------------- spot re-derivations (P/E etc.)
mine = my.set_index("dt")
theirs = xl.set_index("Date")

# Three spot months incl. one October. 1871-10 comes from raw float 1871.1 —
# the exact float-date trap value. Tolerance REL_TOL (same source file).
spot_months = [pd.Timestamp(1871, 10, 1), pd.Timestamp(2008, 10, 1)]
last_e_dt = mine["PE"].dropna().index[-1]
spot_months.append(last_e_dt)
for dt in spot_months:
    if dt not in mine.index or dt not in theirs.index:
        check("FAIL", f"spot P/E {dt.date()}", "month missing from a dataset")
        continue
    mv = mine.at[dt, "PE"]
    tv = theirs.at[dt, "PE"]
    ok = pd.notna(mv) and pd.notna(tv) and close(mv, tv)
    check("PASS" if ok else "FAIL", f"spot trailing P/E {dt.date()}",
          f"own {mv:.6f} vs xlsx {tv:.6f} (rel tol {REL_TOL:g})")

# Latest CAPE / TR_CAPE: xlsx vs raw, REL_TOL.
last_dt = mine.index[-1]
for raw_col, xl_col in [("CAPE", "CAPE"), ("TR CAPE", "TR_CAPE")]:
    mv = mine[raw_col].dropna().iloc[-1]
    mdt = mine[raw_col].dropna().index[-1]
    tv = theirs[xl_col].dropna().iloc[-1]
    tdt = theirs[xl_col].dropna().index[-1]
    ok = mdt == tdt and close(mv, tv)
    check("PASS" if ok else "FAIL", f"latest {xl_col} xlsx vs raw",
          f"own {mv:.4f} @ {mdt.date()} vs xlsx {tv:.4f} @ {tdt.date()}")

# ------------------------------------------ known-benchmark peak re-derivation
# Published/known values quoted to 1dp -> abs tol 0.2.
pk_pe_dt = mine["PE"].idxmax()
pk_pe = mine["PE"].max()
ok = pk_pe_dt == pd.Timestamp(2009, 5, 1) and abs(pk_pe - 123.7) <= 0.2
check("PASS" if ok else "FAIL", "peak trailing P/E (own parse)",
      f"{pk_pe:.2f}x @ {pk_pe_dt.date()} (expect ~123.7x @ 2009-05, tol 0.2)")
tpk = theirs["PE"].max()
check("PASS" if close(pk_pe, tpk) else "FAIL", "peak P/E raw vs xlsx",
      f"own {pk_pe:.6f} vs xlsx {tpk:.6f}")

pk_c_dt = mine["CAPE"].idxmax()
pk_c = mine["CAPE"].max()
ok = pk_c_dt == pd.Timestamp(1999, 12, 1) and abs(pk_c - 44.2) <= 0.2
check("PASS" if ok else "FAIL", "peak CAPE (own parse)",
      f"{pk_c:.2f}x @ {pk_c_dt.date()} (expect ~44.2x @ 1999-12, tol 0.2)")
tpc = theirs["CAPE"].max()
check("PASS" if close(pk_c, tpc) else "FAIL", "peak CAPE raw vs xlsx",
      f"own {pk_c:.6f} vs xlsx {tpc:.6f}")

# Series-wide comparison, not just spots: max relative diff across all
# overlapping non-NaN PE values.
if len(xl) == len(my):
    both = pd.DataFrame({"m": my["PE"].values, "t": xl["PE"].values}).dropna()
    if len(both):
        max_rel = (both["m"] - both["t"]).abs().div(
            both[["m", "t"]].abs().max(axis=1).clip(lower=1.0)).max()
        check("PASS" if max_rel <= REL_TOL else "FAIL",
              "full PE series raw vs xlsx",
              f"{len(both)} rows, max rel diff {max_rel:.2e} (tol {REL_TOL:g})")

# --------------------------------------------------- Summary sheet re-derive
summ = pd.read_excel(XLSX, sheet_name="Summary", index_col=0)
derived = {
    ("PE", "median"): mine["PE"].median(),
    ("CAPE", "median"): mine["CAPE"].median(),
    ("TR_CAPE", "median"): mine["TR CAPE"].median(),
    ("P", "current"): mine["P"].dropna().iloc[-1],
    ("PE", "max"): mine["PE"].max(),
}
bad = []
for (row, col), val in derived.items():
    sval = summ.at[row, col]
    if not close(val, sval, rel=1e-6):
        bad.append(f"{row}.{col}: own {val:.6f} vs summary {sval:.6f}")
check("PASS" if not bad else "FAIL", "Summary sheet stats re-derived",
      "; ".join(bad) if bad else
      f"{len(derived)} stats match within rel 1e-6")

pr_bad = []
for name, raw_col in [("PE", "PE"), ("CAPE", "CAPE"), ("TR_CAPE", "TR CAPE")]:
    s = mine[raw_col].dropna()
    my_pr = s.rank(pct=True).iloc[-1] * 100
    their_pr = summ.at[name, "pct_rank_current"]
    if not close(my_pr, their_pr, rel=1e-6):
        pr_bad.append(f"{name}: own {my_pr:.3f} vs summary {their_pr:.3f}")
check("PASS" if not pr_bad else "FAIL", "percentile ranks re-derived",
      "; ".join(pr_bad) if pr_bad else "PE/CAPE/TR_CAPE pct ranks match")

# ------------------------------------------------------ data-range / staleness
start_ok = xl["Date"].iloc[0] == pd.Timestamp(1871, 1, 1)
check("PASS" if start_ok else "FAIL", "start date",
      f"{xl['Date'].iloc[0].date()} (expect 1871-01-01; intentional exception "
      f"to the 2006 convention, documented in script header)")

age_days = (TODAY - xl["Date"].iloc[-1].to_pydatetime()).days
# Shiller updates monthly; latest price month should be within ~2 months.
check("PASS" if age_days <= 62 else ("WARN" if age_days <= 122 else "FAIL"),
      "latest price month staleness",
      f"{xl['Date'].iloc[-1].date()}, {age_days} days old (tol 62d, warn 122d)")

e_age_days = (TODAY - last_e_dt.to_pydatetime()).days
# Earnings lag price by roughly two quarters by construction -> warn at 8mo.
check("PASS" if e_age_days <= 245 else "WARN", "latest earnings month staleness",
      f"{last_e_dt.date()}, {e_age_days} days old (earnings lag expected; "
      f"warn if > 245d)")

# ------------------------------------------------------- internal consistency
def interior_nans(s):
    v = s.reset_index(drop=True)
    nn = v.notna()
    if not nn.any():
        return -1
    first, last = nn.idxmax(), nn[::-1].idxmax()
    return int(v.loc[first:last].isna().sum())


for col in ["P", "CPI"]:
    n = interior_nans(theirs[col])
    check("PASS" if n == 0 else "FAIL", f"no NaN holes inside {col}",
          f"{n} interior NaNs")
for col in ["E", "PE", "CAPE", "TR_CAPE"]:
    n = interior_nans(theirs[col])
    check("PASS" if n == 0 else "FAIL",
          f"{col} contiguous (NaN only at head/tail)", f"{n} interior NaNs")

# Unit-jump discontinuities: worst genuine monthly moves in this dataset are
# ~±30% for P (1929-32) and <5% for CPI; a parse/paste error (e.g. index vs
# price level mix-up) would show as a far larger jump.
import numpy as np

logret = np.log(theirs["P"]).diff().abs().max()
check("PASS" if logret < 0.45 else "FAIL", "no unit jumps in P",
      f"max |monthly log change| {logret:.3f} (tol 0.45)")
cpiret = np.log(theirs["CPI"]).diff().abs().max()
check("PASS" if cpiret < 0.10 else "FAIL", "no unit jumps in CPI",
      f"max |monthly log change| {cpiret:.3f} (tol 0.10)")

# No forward-filling at the earnings tail: the final price row must have P but
# NaN E and NaN PE, and the raw file's last E month must equal the xlsx's.
last_row = theirs.iloc[-1]
tail_ok = (pd.notna(last_row["P"]) and pd.isna(last_row["E"])
           and pd.isna(last_row["PE"]))
raw_last_e = mine["E"].dropna().index[-1]
xl_last_e = theirs["E"].dropna().index[-1]
check("PASS" if tail_ok and raw_last_e == xl_last_e else "FAIL",
      "earnings tail not forward-filled",
      f"last row {theirs.index[-1].date()}: P={last_row['P']:.2f}, "
      f"E NaN={pd.isna(last_row['E'])}, PE NaN={pd.isna(last_row['PE'])}; "
      f"last E month raw {raw_last_e.date()} == xlsx {xl_last_e.date()}")

# ------------------------------------------------------- external cross-check
# multpl.com publishes a monthly S&P 500 trailing P/E built on LAST REPORTED
# TTM EPS. Shiller's E for recent months is interpolated toward estimated
# future quarters (higher E denominator), so the script's trailing P/E should
# sit BELOW multpl's. Expected as of Jul-2026: multpl ~32 vs script ~25.3.
# This is a methodology gap, not an error -> WARN with explanation, FAIL only
# if the gap direction contradicts the explanation or multpl is unparseable
# in a way that suggests our own regression.
MULTPL_URL = "https://www.multpl.com/s-p-500-pe-ratio/table/by-month"
UA = {"User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                     "AppleWebKit/537.36 (KHTML, like Gecko) "
                     "Chrome/126.0 Safari/537.36")}
script_pe = theirs["PE"].dropna().iloc[-1]
multpl_latest = None
multpl_rows = {}
try:
    import requests

    resp = requests.get(MULTPL_URL, headers=UA, timeout=60)
    resp.raise_for_status()
    pairs = re.findall(
        r"<td>([A-Z][a-z]{2} \d{1,2}, \d{4})</td>\s*<td>\s*"
        r"(?:<abbr[^>]*>[^<]*</abbr>|&#x2002;)?\s*([\d.]+)",
        resp.text,
    )
    for dstr, vstr in pairs:
        multpl_rows[pd.Timestamp(datetime.strptime(dstr, "%b %d, %Y")
                                 .replace(day=1))] = float(vstr)
    if pairs:
        multpl_latest = float(pairs[0][1])
except Exception as exc:  # network/parse issues -> WARN, not FAIL
    check("WARN", "multpl.com fetch", f"unavailable ({exc}); external "
          f"cross-check skipped")

if multpl_latest is not None:
    sane = 15.0 <= multpl_latest <= 60.0
    check("PASS" if sane else "FAIL", "multpl.com latest monthly P/E parsed",
          f"{multpl_latest:.2f} ({len(multpl_rows)} rows parsed; sanity band "
          f"15-60)")
    gap = multpl_latest - script_pe
    dir_ok = gap > 0
    # Reconciliation: reprice the script's latest P against Dec-2025 E (last
    # fully-reported year-end TTM, no forward interpolation). If the gap is
    # really the E-basis difference, this must land between the script value
    # and multpl's latest, in the ~28-31 zone.
    p_latest = theirs["P"].dropna().iloc[-1]
    e_dec25 = mine.at[pd.Timestamp(2025, 12, 1), "E"] \
        if pd.Timestamp(2025, 12, 1) in mine.index else float("nan")
    recon_pe = p_latest / e_dec25
    recon_ok = script_pe < recon_pe < multpl_latest and 27.0 <= recon_pe <= 32.0
    detail = (f"multpl {multpl_latest:.2f} vs script trailing {script_pe:.2f} "
              f"(gap {gap:+.1f}); repricing latest P {p_latest:.0f} on "
              f"Dec-2025 E {e_dec25:.2f} gives {recon_pe:.2f} -- consistent "
              f"with Shiller's E tail being interpolated toward estimated "
              f"future quarters while multpl uses last reported TTM EPS")
    if dir_ok and recon_ok:
        check("WARN", "multpl gap explained (methodology, not error)", detail)
    else:
        check("FAIL", "multpl gap NOT consistent with explanation", detail)
    # Bonus: multpl's own Dec-2025 monthly row vs our Dec-2025 P over
    # Dec-2025 E. Tol 1.5 pts for PASS; a larger gap is still expected to be
    # a vintage effect (Shiller uses monthly-average P and E interpolated to
    # calendar year-end; multpl uses a point-in-time close over the TTM EPS
    # last REPORTED at that date, i.e. an older/lower denominator -> higher
    # P/E) so it WARNs rather than FAILs as long as multpl >= own.
    dec25 = pd.Timestamp(2025, 12, 1)
    if dec25 in multpl_rows and dec25 in mine.index:
        own_dec = mine.at[dec25, "P"] / mine.at[dec25, "E"]
        diff = abs(own_dec - multpl_rows[dec25])
        if diff <= 1.5:
            status = "PASS"
        elif multpl_rows[dec25] >= own_dec and diff <= 4.0:
            status = "WARN"
        else:
            status = "FAIL"
        check(status, "external Dec-2025 P/E vs multpl",
              f"own (Shiller basis) {own_dec:.2f} vs multpl "
              f"{multpl_rows[dec25]:.2f} (diff {diff:.2f}; PASS tol 1.5, "
              f"vintage-consistent gap up to 4.0 is WARN: multpl's EPS "
              f"denominator is the older last-reported TTM)")

# ------------------------------------------------------------------- summary
n_pass = sum(1 for s, _, _ in results if s == "PASS")
n_warn = sum(1 for s, _, _ in results if s == "WARN")
n_fail = sum(1 for s, _, _ in results if s == "FAIL")
print(f"\nSUMMARY: pass={n_pass} warn={n_warn} fail={n_fail}")
sys.exit(1 if n_fail else 0)
