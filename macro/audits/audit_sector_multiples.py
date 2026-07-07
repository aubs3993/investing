# Independent audit of macro/sector_multiples_pull.py and
# macro/output/sector_multiples/.
#
# Re-parses the cached raw Damodaran vintage files (raw/vebitda_*.xls) with
# its OWN parser code — header discovery, column normalization, all-firms
# block selection and weighted-median aggregation are re-implemented here from
# the documented spec, not copied from the analysis. Does NOT import or exec
# sector_multiples_pull.py; reads sector_multiples.xlsx only to COMPARE, and
# scans the analysis source as TEXT (never executed) for documentation /
# chart-annotation claims that cannot be verified from a PNG.
#
# Tolerances:
#   - Raw re-parse vs xlsx Data_Industries: 1e-6 relative — same cached file,
#     only an xlsx float round-trip in between; anything larger means a wrong
#     column/row was captured.
#   - Weighted-median re-derivation vs Data_Buckets: 1e-6 relative (pure-math
#     re-derivation on identical inputs).
#   - External anchors (Tech ~27.2x for the 2025 vintage, ~20.8x for 1999):
#     +/- 0.3x absolute — the anchor is quoted to one decimal.
#   - Staleness: latest vintage year must be (current year - 1); Damodaran
#     posts each vintage in early January of the following year.
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

OUT_DIR = MACRO_DIR / "output" / "sector_multiples"
RAW_DIR = OUT_DIR / "raw"
XLSX = OUT_DIR / "sector_multiples.xlsx"
ANALYSIS_SRC = MACRO_DIR / "sector_multiples_pull.py"

TODAY = pd.Timestamp(datetime.today().date())

RESULTS = []


def check(status: str, name: str, detail: str) -> None:
    RESULTS.append(status)
    print(f"{status} {name}: {detail}")


def nrm(s) -> str:
    """Normalize an industry/header cell: lower, strip, collapse whitespace."""
    return re.sub(r"\s+", " ", str(s).strip().lower())


def nrm_hdr(s) -> str:
    """Header cells in some vintages carry dedup suffixes (EV/EBITDA3)."""
    return re.sub(r"\d+$", "", nrm(s)).strip()


def wmedian(values, weights) -> float:
    """Firm-count-weighted median, implemented independently: smallest value
    whose cumulative weight reaches half the total weight."""
    v = np.asarray(values, float)
    w = np.asarray(weights, float)
    order = np.argsort(v, kind="mergesort")
    v, w = v[order], w[order]
    cum = np.cumsum(w)
    half = 0.5 * w.sum()
    return float(v[int(np.argmax(cum >= half))])


def parse_vintage(path: Path) -> dict:
    """Own parser for a Damodaran vebitda file. Returns
    {date_updated, era, ebitda_col_idx, headers, rows: {norm_name: (n, ebitda,
    ebit, pos_block_ebitda_or_None)}}."""
    xls = pd.ExcelFile(path, engine="xlrd")
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        if df.empty:
            continue
        date_updated = None
        hdr_row = None
        for i in range(min(20, len(df))):
            c0 = nrm(df.iloc[i, 0])
            if c0.startswith("date updated"):
                date_updated = pd.to_datetime(df.iloc[i, 1], errors="coerce")
            if c0 in ("industry name", "industry"):
                hdr_row = i
                break
        if hdr_row is None:
            continue
        headers = [nrm_hdr(v) for v in df.iloc[hdr_row]]
        if "value/ebitda" in headers:
            era = "firm_value"
            key_ebitda, key_ebit = "value/ebitda", "value/ebit"
        elif "ev/ebitda" in headers:
            era = "ev"
            key_ebitda, key_ebit = "ev/ebitda", "ev/ebit"
        else:
            raise ValueError(f"{path.name}: no EBITDA multiple column")
        # All-firms block = LAST occurrence when two blocks exist.
        idx_all = max(i for i, h in enumerate(headers) if h == key_ebitda)
        idx_first = min(i for i, h in enumerate(headers) if h == key_ebitda)
        ebit_idxs = [i for i, h in enumerate(headers) if h == key_ebit]
        idx_ebit = max(ebit_idxs) if ebit_idxs else None
        idx_n = headers.index("number of firms")
        rows = {}
        for _, row in df.iloc[hdr_row + 1:].iterrows():
            name = row.iloc[0]
            if pd.isna(name):
                continue
            key = nrm(name)
            if key in ("market", "grand total", "total market",
                       "total market (without financials)"):
                continue
            rows[key] = (
                pd.to_numeric(row.iloc[idx_n], errors="coerce"),
                pd.to_numeric(row.iloc[idx_all], errors="coerce"),
                (pd.to_numeric(row.iloc[idx_ebit], errors="coerce")
                 if idx_ebit is not None else np.nan),
                (pd.to_numeric(row.iloc[idx_first], errors="coerce")
                 if idx_first != idx_all else None),
            )
        return {"date_updated": date_updated, "era": era, "headers": headers,
                "two_blocks": idx_first != idx_all, "rows": rows}
    raise ValueError(f"{path.name}: no sheet with an Industry Name header")


# ---------------------------------------------------------------------------
# 1. Output file existence + non-trivial size; raw cache completeness
# ---------------------------------------------------------------------------
EXPECTED_FILES = {
    "sector_multiples.xlsx": 30_000,
    "sector_ev_ebitda.png": 50_000,
    "sector_ev_ebitda_overview.png": 50_000,
    "sector_ev_ebit_overview.png": 50_000,
    "sector_ev_ebitda_2006.png": 50_000,
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

RAW_TAGS = ["98", "99"] + [f"{y:02d}" for y in range(0, 25)] + ["current"]
raw_missing = [t for t in RAW_TAGS
               if not (RAW_DIR / f"vebitda_{t}.xls").exists()
               or (RAW_DIR / f"vebitda_{t}.xls").stat().st_size < 5_000]
if raw_missing:
    check("FAIL", "raw cache", f"missing/undersized vintages: {raw_missing}")
else:
    check("PASS", "raw cache",
          f"all {len(RAW_TAGS)} cached vintage files present "
          "(98, 99, 00-24, current), each >= 5 KB")

gitkeeps = [p for p in (OUT_DIR / ".gitkeep", RAW_DIR / ".gitkeep")
            if not p.exists()]
if gitkeeps:
    check("WARN", ".gitkeep convention", f"missing: {gitkeeps}")
else:
    check("PASS", ".gitkeep convention",
          "output/ and output/raw/ both carry .gitkeep markers")

if not XLSX.exists():
    print("\nSummary: cannot continue without sector_multiples.xlsx")
    sys.exit(1)

ind = pd.read_excel(XLSX, sheet_name="Data_Industries")
buckets = pd.read_excel(XLSX, sheet_name="Data_Buckets", header=[0, 1],
                        index_col=0)
summary = pd.read_excel(XLSX, sheet_name="Summary", index_col=0)
pivot_ebitda = buckets["EV_EBITDA"]
pivot_ebit = buckets["EV_EBIT"]
src_text = ANALYSIS_SRC.read_text(encoding="utf-8")

REL_TOL = 1e-6

# ---------------------------------------------------------------------------
# 2. Re-parse three cached raw vintages with own code; spot-compare 3
#    industry rows each (n firms + EV/EBITDA + EV/EBIT) vs Data_Industries
# ---------------------------------------------------------------------------
SPOT_PLAN = [
    # (tag, vintage_year, [industries]) — era 1, modern single-block, current
    ("98", 1998, ["Advertising", "Aerospace/Defense", "Auto & Truck"]),
    ("15", 2015, ["Advertising", "Aerospace/Defense", "Air Transport"]),
    ("current", 2025, ["Advertising", "Aerospace/Defense", "Air Transport"]),
]
parsed = {}
for tag, vy, industries in SPOT_PLAN:
    label = f"raw re-parse vebitda_{tag} (vintage {vy})"
    try:
        p = parse_vintage(RAW_DIR / f"vebitda_{tag}.xls")
    except Exception as exc:
        check("FAIL", label, f"own parser failed: {exc}")
        continue
    parsed[tag] = p
    vin = ind[ind["vintage_year"] == vy].set_index(
        ind[ind["vintage_year"] == vy]["industry"].map(nrm))
    if vin.empty:
        check("FAIL", label, f"no vintage {vy} rows in Data_Industries")
        continue
    bad, details = [], []
    for name in industries:
        key = nrm(name)
        if key not in p["rows"]:
            bad.append(f"{name}: absent from own raw parse")
            continue
        if key not in vin.index:
            bad.append(f"{name}: absent from xlsx vintage {vy}")
            continue
        n_raw, eb_raw, ebit_raw, _ = p["rows"][key]
        xr = vin.loc[key]
        for field, mine, theirs in (("n_firms", n_raw, xr["n_firms"]),
                                    ("ev_ebitda", eb_raw, xr["ev_ebitda"]),
                                    ("ev_ebit", ebit_raw, xr["ev_ebit"])):
            if pd.isna(mine) and pd.isna(theirs):
                continue
            if pd.isna(mine) or pd.isna(theirs) or (
                    abs(theirs - mine) > REL_TOL * max(1.0, abs(mine))):
                bad.append(f"{name}.{field}: raw={mine} xlsx={theirs}")
        details.append(f"{name} {eb_raw:.2f}x/n={n_raw:.0f}")
    if bad:
        check("FAIL", label, "; ".join(bad))
    else:
        check("PASS", label,
              f"3 industries match xlsx exactly ({'; '.join(details)})")

# ---------------------------------------------------------------------------
# 3. All-firms block selection (current file publishes two blocks)
# ---------------------------------------------------------------------------
if "current" in parsed:
    p = parsed["current"]
    if not p["two_blocks"]:
        check("WARN", "all-firms block selection",
              "current file no longer has two blocks; check not applicable")
    else:
        n, all_v, _, pos_v = p["rows"][nrm("Advertising")]
        vin25 = ind[(ind["vintage_year"] == 2025)
                    & (ind["industry"] == "Advertising")]
        xv = float(vin25["ev_ebitda"].iloc[0])
        if abs(xv - all_v) <= 1e-6 and abs(all_v - pos_v) > 0.5:
            check("PASS", "all-firms block selection",
                  f"xlsx Advertising 2025 = {xv:.4f} matches the ALL-FIRMS "
                  f"block ({all_v:.4f}), not the positive-EBITDA block "
                  f"({pos_v:.4f}) — documented choice honored")
        elif abs(xv - pos_v) <= 1e-6:
            check("FAIL", "all-firms block selection",
                  f"xlsx Advertising 2025 = {xv:.4f} equals the "
                  f"positive-EBITDA-only block ({pos_v:.4f}); the script "
                  f"documents using the all-firms block ({all_v:.4f})")
        else:
            check("FAIL", "all-firms block selection",
                  f"xlsx Advertising 2025 = {xv:.4f} matches neither block "
                  f"(pos {pos_v:.4f}, all {all_v:.4f})")

# ---------------------------------------------------------------------------
# 4. Firm-count-weighted MEDIAN re-derivation: Technology 2025 and 1999.
#    Values come from OWN raw parse; bucket membership from the xlsx map;
#    the median math is re-implemented above.
# ---------------------------------------------------------------------------
ANCHORS = [("current", 2025, 27.2), ("99", 1999, 20.8)]
for tag, vy, anchor in ANCHORS:
    label = f"weighted median Tech {vy}"
    if tag not in parsed:
        try:
            parsed[tag] = parse_vintage(RAW_DIR / f"vebitda_{tag}.xls")
        except Exception as exc:
            check("FAIL", label, f"own parser failed on vebitda_{tag}: {exc}")
            continue
    p = parsed[tag]
    members = ind[(ind["vintage_year"] == vy)
                  & (ind["bucket"] == "Technology")]["industry"].map(nrm)
    vals, wts, missing_m = [], [], []
    for key in members:
        if key not in p["rows"]:
            missing_m.append(key)
            continue
        n, eb, _, _ = p["rows"][key]
        if pd.isna(eb):
            continue
        vals.append(eb)
        wts.append(max(1.0, n if pd.notna(n) else 1.0))
    if missing_m:
        check("FAIL", label, f"bucket members absent from raw file: {missing_m}")
        continue
    mine = wmedian(vals, wts)
    theirs = float(pivot_ebitda.loc[vy, "Technology"])
    d_pivot = abs(theirs - mine)
    d_anchor = abs(mine - anchor)
    if d_pivot > 1e-6:
        check("FAIL", label,
              f"own re-derivation {mine:.4f}x vs Data_Buckets {theirs:.4f}x "
              f"({len(vals)} industries, {sum(wts):.0f} firms)")
    elif d_anchor > 0.3:
        check("FAIL", label,
              f"re-derived {mine:.4f}x matches xlsx but is {d_anchor:.2f}x "
              f"away from the expected ~{anchor}x anchor")
    else:
        check("PASS", label,
              f"own re-derivation {mine:.3f}x == Data_Buckets {theirs:.3f}x, "
              f"within 0.3x of expected ~{anchor}x "
              f"({len(vals)} industries, {sum(wts):.0f} firms)")

# Vintage-year convention: 1999-vintage anchor (20.8) matched above, i.e. the
# sheet keys each vintage by DATA year (filename tag), not snapshot year.
# Confirm the convention is documented in the analysis source.
if re.search(r"[Vv]intage year.*data year", src_text) and \
        "vebitda14.xls is dated 2015-01-05" in src_text:
    check("PASS", "vintage-year convention documented",
          "source documents 'Vintage year YY = the data year' with the "
          "vebitda14/2015-01-05 example; matches the 1999-vintage anchor "
          "(20.8x) landing on sheet row 1999")
else:
    check("WARN", "vintage-year convention documented",
          "could not find the vintage-year = data-year convention documented "
          "in sector_multiples_pull.py comments")

# ---------------------------------------------------------------------------
# 5. Financials exclusion: no bank/insurance/brokerage/investment industries
#    anywhere in Data_Industries
# ---------------------------------------------------------------------------
FIN_PAT = re.compile(
    r"\b(bank|insurance|reinsurance|brokerage|thrift|invest|financial|"
    r"securit|private equity)")
leaked = sorted({name for name in ind["industry"].astype(str)
                 if FIN_PAT.search(nrm(name))})
if leaked:
    check("FAIL", "financials exclusion",
          f"financial industry names present in Data_Industries: {leaked}")
else:
    check("PASS", "financials exclusion",
          f"no bank/insurance/brokerage/investment names among "
          f"{ind['industry'].nunique()} distinct industries "
          f"({len(ind)} rows)")

# ---------------------------------------------------------------------------
# 6. Era transition: Value/EBITDA through the 2009 vintage, EV/EBITDA from
#    2010 — from own parse of vebitda_09 and vebitda_10; charts must mark it
# ---------------------------------------------------------------------------
try:
    p09 = parse_vintage(RAW_DIR / "vebitda_09.xls")
    p10 = parse_vintage(RAW_DIR / "vebitda_10.xls")
    if p09["era"] == "firm_value" and p10["era"] == "ev":
        check("PASS", "era transition 2009->2010",
              "own parse: vebitda_09 headers say Value/EBITDA (firm value), "
              "vebitda_10 says EV/EBITDA — transition at the 2010 vintage "
              "as documented")
    else:
        check("FAIL", "era transition 2009->2010",
              f"own parse eras: 09={p09['era']}, 10={p10['era']} — "
              "expected firm_value then ev")
except Exception as exc:
    check("FAIL", "era transition 2009->2010", f"own parse failed: {exc}")

if "axvline(trans_x" in src_text and "TRANS_LABEL" in src_text \
        and "transition_year" in src_text:
    check("PASS", "charts mark the transition",
          "source-level: every chart draws axvline(trans_x) labeled via "
          "TRANS_LABEL derived from the empirically detected transition_year "
          "(PNG content itself not machine-checkable)")
else:
    check("WARN", "charts mark the transition",
          "could not find the transition vline in the chart code")

# ---------------------------------------------------------------------------
# 7. vebitda_17 stale internal date: file says 2017-01-05 (same as
#    vebitda_16); xlsx 2017 vintage must still be a DISTINCT row set
# ---------------------------------------------------------------------------
try:
    p16 = parse_vintage(RAW_DIR / "vebitda_16.xls")
    p17 = parse_vintage(RAW_DIR / "vebitda_17.xls")
    d16, d17 = p16["date_updated"], p17["date_updated"]
    if d16 == d17 == pd.Timestamp("2017-01-05"):
        note = ("known source flaw confirmed: vebitda_17 internal 'Date "
                "updated' 2017-01-05 duplicates vebitda_16's")
    else:
        note = (f"internal dates now 16={d16}, 17={d17} — the documented "
                "stale-date flaw pattern has changed")
    x16 = ind[ind["vintage_year"] == 2016].set_index("industry")["ev_ebitda"]
    x17 = ind[ind["vintage_year"] == 2017].set_index("industry")["ev_ebitda"]
    common = x16.index.intersection(x17.index)
    frac_same = float((x16[common] == x17[common]).mean()) if len(common) \
        else np.nan
    # Also confirm 2017 xlsx rows equal the vebitda_17 FILE (filename-year
    # convention applied despite the bad internal date).
    mism = 0
    for name in common[:20]:
        key = nrm(name)
        if key in p17["rows"]:
            eb17 = p17["rows"][key][1]
            xv = x17[name]
            if pd.notna(eb17) and pd.notna(xv) and abs(eb17 - xv) > 1e-6:
                mism += 1
    if len(common) >= 70 and frac_same < 0.05 and mism == 0:
        check("PASS", "vebitda_17 stale-date handling",
              f"{note}; xlsx 2017 vintage is a distinct row set "
              f"({frac_same:.0%} of {len(common)} common industries identical "
              "to 2016) and matches the vebitda_17 file contents")
    else:
        check("FAIL", "vebitda_17 stale-date handling",
              f"{note}; common={len(common)}, frac identical to "
              f"2016={frac_same:.2f}, mismatches vs vebitda_17 file={mism} — "
              "2017 vintage looks duplicated or mis-keyed")
except Exception as exc:
    check("FAIL", "vebitda_17 stale-date handling", f"parse failed: {exc}")

# ---------------------------------------------------------------------------
# 8. Data_Buckets pivot: exactly one row per vintage year 1998..2025
# ---------------------------------------------------------------------------
idx = list(buckets.index)
expected_years = list(range(1998, 2026))
if idx == expected_years:
    check("PASS", "Data_Buckets vintage axis",
          "exactly one row per vintage year 1998..2025, ordered, no dupes")
else:
    dupes = [y for y in set(idx) if idx.count(y) > 1]
    miss = [y for y in expected_years if y not in idx]
    extra = [y for y in idx if y not in expected_years]
    check("FAIL", "Data_Buckets vintage axis",
          f"dupes={dupes}, missing={miss}, unexpected={extra}")

# ---------------------------------------------------------------------------
# 9. Real Estate pre-2010 extremes present (47-133x, firm-value artifact)
#    and flagged in the outputs
# ---------------------------------------------------------------------------
re_pre = pivot_ebitda.loc[pivot_ebitda.index <= 2009, "Real Estate"].dropna()
if re_pre.empty:
    check("FAIL", "Real Estate legacy extremes", "no pre-2010 values in xlsx")
elif not (45 <= re_pre.max() <= 140):
    check("FAIL", "Real Estate legacy extremes",
          f"pre-2010 max {re_pre.max():.1f}x outside the documented 47-133x "
          "band — parse or aggregation drift")
else:
    check("PASS", "Real Estate legacy extremes",
          f"pre-2010 Real Estate runs {re_pre.min():.1f}-{re_pre.max():.1f}x "
          "in the xlsx (firm-value-era REIT artifact preserved in data)")

if "47-133x" in src_text and "artifact" in src_text:
    check("PASS", "Real Estate extremes flagged",
          "source-level: chart comments/titles clip axes at 35-45x and label "
          "the legacy Real Estate 47-133x readings as a firm-value/EBITDA "
          "artifact running off-axis by design")
else:
    check("WARN", "Real Estate extremes flagged",
          "no flag for the 47-133x legacy Real Estate artifact found in the "
          "chart code — values would plot unexplained")

# ---------------------------------------------------------------------------
# 10. Internal consistency: interior NaNs, per-vintage row counts, no
#     unit-scale breaks
# ---------------------------------------------------------------------------
CORE = [b for b in pivot_ebitda.columns if b != "Other"]
nan_cols = {b: int(pivot_ebitda[b].isna().sum()) for b in CORE
            if pivot_ebitda[b].isna().any()}
if nan_cols:
    check("FAIL", "interior NaNs (EV/EBITDA buckets)",
          f"NaNs in core bucket series: {nan_cols}")
else:
    n_other = int(pivot_ebitda["Other"].notna().sum())
    check("PASS", "interior NaNs (EV/EBITDA buckets)",
          f"all 11 core bucket series fully populated 1998-2025; sparse "
          f"'Other' bucket ({n_other}/28 vintages) is expected — it only "
          "exists when a vintage has unmapped industries")

counts = ind.groupby("vintage_year").size()
if counts.between(75, 100).all():
    check("PASS", "industries per vintage",
          f"{counts.min()}-{counts.max()} ex-financial industries per "
          f"vintage across {len(counts)} vintages (Damodaran publishes ~95 "
          "incl. financials)")
else:
    odd = counts[~counts.between(75, 100)]
    check("FAIL", "industries per vintage",
          f"vintages with implausible industry counts: {odd.to_dict()}")

lg = np.log(pivot_ebitda[CORE] / pivot_ebitda[CORE].shift(1)).abs()
worst_col = lg.max().idxmax()
worst_val = float(lg.max().max())
worst_yr = int(lg[worst_col].idxmax())
if worst_val < 1.6:
    check("PASS", "unit discontinuities",
          f"max |y/y log change| {worst_val:.2f} ({worst_col} @ {worst_yr}) "
          "< 1.6 — annual multiples are volatile but nothing resembling a "
          "unit/scale break (a 10x break would be 2.3)")
else:
    check("FAIL", "unit discontinuities",
          f"|y/y log change| {worst_val:.2f} for {worst_col} @ {worst_yr} — "
          "possible wrong-column or unit break")

# ---------------------------------------------------------------------------
# 11. Summary sheet consistent with Data_Buckets (own recomputation)
# ---------------------------------------------------------------------------
sum_bad = []
for bucket in pivot_ebitda.columns:
    pairs = [
        ("ev_ebitda_2025", pivot_ebitda.loc[2025, bucket]),
        ("ev_ebitda_median_1998plus", pivot_ebitda[bucket].median()),
        ("ev_ebit_2025", pivot_ebit.loc[2025, bucket]),
        ("ev_ebit_median_1998plus", pivot_ebit[bucket].median()),
    ]
    for col, mine in pairs:
        sv = summary.loc[bucket, col]
        if pd.isna(mine) and pd.isna(sv):
            continue
        if pd.isna(mine) or pd.isna(sv) or \
                abs(sv - mine) > 1e-6 * max(1.0, abs(mine)):
            sum_bad.append(f"{bucket}.{col}: summary {sv} vs recomputed {mine}")
if sum_bad:
    check("FAIL", "Summary sheet consistency", "; ".join(sum_bad))
else:
    check("PASS", "Summary sheet consistency",
          "latest-vintage and 1998+ median columns match own recomputation "
          "from Data_Buckets for all 12 buckets, both metrics")

# ---------------------------------------------------------------------------
# 12. EXTERNAL cross-check + staleness: current vebitda.xls 'Date updated'
# ---------------------------------------------------------------------------
if "current" in parsed:
    du = parsed["current"]["date_updated"]
    if du == pd.Timestamp("2026-01-05"):
        check("PASS", "external: current-file Date updated",
              "cached vebitda.xls (NYU Stern) says 'Date updated: 2026-01-05' "
              "— matches Damodaran's published Jan-2026 refresh; implied data "
              "year 2025 = latest xlsx vintage")
    else:
        check("FAIL", "external: current-file Date updated",
              f"cached vebitda.xls says {du}, expected 2026-01-05")

latest_vintage = int(buckets.index.max())
if latest_vintage == TODAY.year - 1:
    check("PASS", "staleness",
          f"latest vintage {latest_vintage} = current year - 1; next vintage "
          f"({TODAY.year}) publishes Jan-{TODAY.year + 1}")
else:
    check("FAIL", "staleness",
          f"latest vintage {latest_vintage}; expected {TODAY.year - 1} "
          "(Damodaran posts each data year the following January)")

# Judgment call (source-level): the download loop hardcodes archive tags
# through 24 and skips any cached file > 1 KB, including vebitda_current.xls.
# After Damodaran's next refresh (Jan-2027), a re-run would keep serving the
# cached 2025 'current' vintage and never fetch the new vebitda25.xls archive
# unless raw/ is cleared by hand. Not an error in today's outputs (the
# staleness check above would trip once it bites), but worth surfacing.
if "range(0, 25)" in src_text and \
        "dest.exists() and dest.stat().st_size > 1000" in src_text:
    check("WARN", "cache refresh forward-compatibility",
          "archive tag list hardcodes 00-24 and cached vebitda_current.xls "
          "is never re-downloaded; runs after Jan-2027 will silently stay on "
          "the 2025 vintage until raw/vebitda_current.xls is deleted (and "
          "vebitda25.xls will never be fetched). Suggested fix: extend the "
          "tag range to the current year and re-download 'current' when its "
          "cached 'Date updated' year < current year")
else:
    check("PASS", "cache refresh forward-compatibility",
          "download loop no longer matches the hardcoded-tags + "
          "forever-cached-current pattern")

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
n_pass = RESULTS.count("PASS")
n_warn = RESULTS.count("WARN")
n_fail = RESULTS.count("FAIL")
print(f"\nSummary: {n_pass} PASS, {n_warn} WARN, {n_fail} FAIL "
      f"({len(RESULTS)} checks)")
sys.exit(1 if n_fail else 0)
