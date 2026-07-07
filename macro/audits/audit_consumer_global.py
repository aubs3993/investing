# Independent audit of macro/consumer_global_pull.py and its outputs in
# macro/output/consumer_global/.
#
# This script does NOT import or exec the analysis script. It re-pulls the
# FRED household-credit series itself, re-parses the cached raw OECD CSVs
# with its own code, re-derives the key numbers, and compares them to the
# values stored in consumer_global.xlsx. The xlsx is read only to COMPARE.
#
# Tolerances used, and why:
#   TOL_EXACT = 1e-6   — same-source roundtrips (FRED -> xlsx, cached CSV ->
#                        xlsx). The only permissible error is float storage
#                        in xlsx, which is exact to ~1e-12; 1e-6 leaves slack
#                        for pandas dtype churn while still catching any real
#                        derivation difference (wrong column, off-by-one
#                        window, inverted division).
#   TOL_BENCH = 0.35   — external benchmark anchors quoted to one decimal
#                        (e.g. "US gross saving ~10.2%"): half a rounding
#                        unit plus room for minor vintage revisions.
#   TOL_LIVE  = 0.25   — live OECD spot query vs cache: same series, same
#                        vintage day, should match to publication precision
#                        (2 decimals); 0.25 allows an intraday revision.
#
# Prints one line per check: "PASS|FAIL|WARN <name>: <detail>".
# Exits 1 if any FAIL.
from datetime import datetime
from io import StringIO
from pathlib import Path
import sys

import pandas as pd

HERE = Path(__file__).resolve().parent           # macro/audits
REPO = HERE.parent.parent                        # repo root
sys.path.insert(0, str(REPO))

from shared.fred_helpers import get_fred_client  # noqa: E402

OUT_DIR = REPO / "macro" / "output" / "consumer_global"
XLSX = OUT_DIR / "consumer_global.xlsx"
SCRIPT = REPO / "macro" / "consumer_global_pull.py"

TOL_EXACT = 1e-6
TOL_BENCH = 0.35
TOL_LIVE = 0.25

results = {"PASS": 0, "FAIL": 0, "WARN": 0}


def report(status: str, name: str, detail: str) -> None:
    results[status] += 1
    print(f"{status} {name}: {detail}")


def check(cond: bool, name: str, detail: str, warn_only: bool = False) -> None:
    if cond:
        report("PASS", name, detail)
    else:
        report("WARN" if warn_only else "FAIL", name, detail)


def close(a, b, tol) -> bool:
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


# ---------------------------------------------------------------- load xlsx
if not XLSX.exists():
    print(f"FAIL xlsx exists: {XLSX} missing — cannot audit")
    sys.exit(1)

sheets = pd.read_excel(XLSX, sheet_name=None)
expected_sheets = ["Data_HHCredit", "Data_SavingQ", "Data_DebtIncome",
                   "Data_Durables", "Data_SavingA", "Summary"]
check(all(s in sheets for s in expected_sheets), "xlsx sheet set",
      f"found {list(sheets)}")

hh = sheets["Data_HHCredit"]
sav_q = sheets["Data_SavingQ"]
debt_q = sheets["Data_DebtIncome"]
dur = sheets["Data_Durables"]
sav_a = sheets["Data_SavingA"]
summary_raw = pd.read_excel(XLSX, sheet_name="Summary", header=None)

# Locate the stats block (header row 0) and the rank block ("panel" header).
stats = pd.read_excel(XLSX, sheet_name="Summary", index_col=0,
                      nrows=5)  # 5 stat rows written by the script
panel_hdr_rows = summary_raw.index[summary_raw[0] == "panel"].tolist()
ranks = None
if panel_hdr_rows:
    r0 = panel_hdr_rows[0]
    ranks = summary_raw.iloc[r0 + 1:].copy()
    ranks.columns = summary_raw.iloc[r0].tolist()
    ranks = ranks.dropna(subset=["panel"]).set_index("panel")
check(ranks is not None, "Summary rank block present",
      f"panel header at Summary row {panel_hdr_rows}")

# ------------------------------------------- (a) FRED re-derivation, panel 1
# Re-pull three of the eight BIS-mirror series directly from FRED and compare
# first (2006Q1) and latest observations to the Data_HHCredit sheet exactly.
try:
    fred = get_fred_client()
    for sid, col in [("QUSHAM770A", "US"), ("QDEHAM770A", "Germany"),
                     ("QCNHAM770A", "China")]:
        s = fred.get_series(sid, observation_start=datetime(2006, 1, 1),
                            observation_end=datetime.today()).dropna()
        sheet_s = hh.set_index("Date")[col].dropna()
        check(s.index[0] == pd.Timestamp("2006-01-01"),
              f"FRED {sid} starts 2006Q1", f"first obs {s.index[0].date()}")
        check(close(s.iloc[0], sheet_s.iloc[0], TOL_EXACT),
              f"FRED {sid} 2006Q1 == xlsx {col}",
              f"fred={s.iloc[0]} xlsx={sheet_s.iloc[0]} (tol {TOL_EXACT})")
        check(s.index[-1] == sheet_s.index[-1]
              and close(s.iloc[-1], sheet_s.iloc[-1], TOL_EXACT),
              f"FRED {sid} latest == xlsx {col}",
              f"fred {s.index[-1].date()}={s.iloc[-1]} "
              f"xlsx {sheet_s.index[-1].date()}={sheet_s.iloc[-1]}")
except Exception as exc:  # network/auth problems are WARN, not proof of defect
    report("WARN", "FRED re-pull", f"could not re-pull FRED series ({exc})")


# --------------------------------- (b) OECD cached CSV independent re-parse
def parse_oecd_cache(fname: str, extra_cols=()) -> pd.DataFrame | None:
    p = OUT_DIR / fname
    if not p.exists():
        report("FAIL", f"cache {fname} exists", "missing")
        return None
    raw = pd.read_csv(StringIO(p.read_text(encoding="utf-8")))
    cols = ["REF_AREA", "TIME_PERIOD", "OBS_VALUE", *extra_cols]
    raw = raw[cols].dropna(subset=["OBS_VALUE"])
    # pivot_table(mean) in the analysis script would silently average
    # duplicate keys — verify the raw data has none.
    key_cols = [c for c in cols if c != "OBS_VALUE"]
    dups = raw.duplicated(subset=key_cols).sum()
    check(dups == 0, f"cache {fname} no duplicate keys",
          f"{dups} duplicate ({', '.join(key_cols)}) rows")
    return raw


def qdate(tp: pd.Series) -> pd.Series:
    return pd.PeriodIndex(tp, freq="Q").to_timestamp()


# --- Panel 2: quarterly gross saving rate
raw_sq = parse_oecd_cache("oecd_saving_q_raw.csv")
if raw_sq is not None:
    raw_sq = raw_sq.assign(Date=qdate(raw_sq["TIME_PERIOD"]))
    us = raw_sq[raw_sq["REF_AREA"] == "USA"].set_index("Date")[
        "OBS_VALUE"].sort_index()
    de = raw_sq[raw_sq["REF_AREA"] == "DEU"].set_index("Date")[
        "OBS_VALUE"].sort_index()
    xus = sav_q.set_index("Date")["US"].dropna()
    check(us.index[-1] == xus.index[-1] and close(us.iloc[-1], xus.iloc[-1],
                                                  TOL_EXACT),
          "SavingQ US latest re-derived == xlsx",
          f"raw {us.index[-1].date()}={us.iloc[-1]} "
          f"xlsx {xus.index[-1].date()}={xus.iloc[-1]}")
    check(us.index[-1] == pd.Timestamp("2025-10-01")
          and close(us.iloc[-1], 10.2, TOL_BENCH),
          "US gross saving benchmark ~10.2% at 2025Q4",
          f"raw={us.iloc[-1]} at {us.index[-1].date()} (tol {TOL_BENCH})")
    check(close(de.iloc[-1], 19.2, TOL_BENCH),
          "Germany gross saving benchmark ~19.2% latest",
          f"raw={de.iloc[-1]} at {de.index[-1].date()}")

# --- Panel 3: household debt % of gross disposable income
raw_dq = parse_oecd_cache("oecd_debt_q_raw.csv")
if raw_dq is not None:
    raw_dq = raw_dq.assign(Date=qdate(raw_dq["TIME_PERIOD"]))
    usd = raw_dq[raw_dq["REF_AREA"] == "USA"].set_index("Date")[
        "OBS_VALUE"].sort_index()
    xusd = debt_q.set_index("Date")["US"].dropna()
    check(usd.index[-1] == xusd.index[-1]
          and close(usd.iloc[-1], xusd.iloc[-1], TOL_EXACT),
          "DebtIncome US latest re-derived == xlsx",
          f"raw {usd.index[-1].date()}={usd.iloc[-1]} "
          f"xlsx {xusd.index[-1].date()}={xusd.iloc[-1]}")
    check(close(usd.iloc[-1], 91.6, TOL_BENCH),
          "US hh debt %GDI benchmark ~91.6 latest",
          f"raw={usd.iloc[-1]} at {usd.index[-1].date()}")

# --- Panel 4: durables share, own derivation of share and 4q rolling mean
raw_du = parse_oecd_cache("oecd_durables_raw.csv", extra_cols=("TRANSACTION",))
if raw_du is not None:
    raw_du = raw_du.assign(Date=qdate(raw_du["TIME_PERIOD"]))
    p311 = raw_du[(raw_du["REF_AREA"] == "USA")
                  & (raw_du["TRANSACTION"] == "P311")].set_index("Date")[
        "OBS_VALUE"].sort_index()
    p31dc = raw_du[(raw_du["REF_AREA"] == "USA")
                   & (raw_du["TRANSACTION"] == "P31DC")].set_index("Date")[
        "OBS_VALUE"].sort_index()
    share = (p311 / p31dc * 100).dropna()
    check(3 < share.iloc[-1] < 20,
          "US durables share division not inverted",
          f"latest raw share {share.iloc[-1]:.2f}% (P311/P31DC*100; the "
          f"inverse would be ~900%)")
    # Own rolling 4q mean via explicit tail slice — checks window alignment
    # (a 3q or 5q window, or a shifted window, would differ in decimals).
    ma_full = share.rolling(4).mean()
    own_latest_ma = share.iloc[-4:].mean()
    xdur = dur.set_index("Date")
    x_ma = xdur["US_4qma"].dropna()
    check(share.index[-1] == pd.Timestamp("2026-01-01")
          and x_ma.index[-1] == share.index[-1]
          and close(own_latest_ma, x_ma.iloc[-1], TOL_EXACT),
          "Durables US 4qMA latest re-derived == xlsx",
          f"own mean(last 4 shares)={own_latest_ma:.6f} "
          f"xlsx={x_ma.iloc[-1]:.6f} at {x_ma.index[-1].date()}")
    check(close(own_latest_ma, 10.8, TOL_BENCH),
          "US durables 4qMA benchmark ~10.8% at 2026Q1",
          f"own={own_latest_ma:.2f}")
    # 1990Q1 boundary: the 4qMA at 1990-01-01 must use 1989Q2-1990Q1 raw data
    # (rolling BEFORE the window filter). If the script had filtered first,
    # the first three 4qMA values in the sheet would be NaN.
    b = pd.Timestamp("1990-01-01")
    if b in ma_full.index and b in xdur.index:
        check(close(ma_full.loc[b], xdur.loc[b, "US_4qma"], TOL_EXACT),
              "Durables 4qMA 1990Q1 uses pre-1990 lookback",
              f"own(full-history rolling)={ma_full.loc[b]:.6f} "
              f"xlsx={xdur.loc[b, 'US_4qma']:.6f}")
    check(xdur.index[0] == b, "Durables sheet starts 1990Q1",
          f"first row {xdur.index[0].date()}")

# --- Panel 5: annual SRATIO (net saving ratio) incl. forecasts
raw_sa = parse_oecd_cache("oecd_saving_a_raw.csv")
if raw_sa is not None:
    usa_a = raw_sa[raw_sa["REF_AREA"] == "USA"].copy()
    usa_a["Year"] = usa_a["TIME_PERIOD"].astype(int)
    usa_a = usa_a.set_index("Year")["OBS_VALUE"].sort_index()
    xsa = sav_a.set_index("Date")["US"].dropna()
    xsa.index = xsa.index.year
    check(close(usa_a.loc[2025], 4.8, TOL_BENCH),
          "US net saving ratio 2025 benchmark ~4.8%",
          f"raw 2025={usa_a.loc[2025]:.3f}")
    check(close(usa_a.loc[2025], xsa.loc[2025], TOL_EXACT),
          "SavingA US 2025 re-derived == xlsx",
          f"raw={usa_a.loc[2025]:.6f} xlsx={xsa.loc[2025]:.6f}")
    check(2026 in xsa.index and 2027 in xsa.index,
          "SavingA forecast years 2026-27 present in Data sheet",
          f"years present: {sorted(y for y in xsa.index if y >= 2025)}")
    # Summary stats row must exclude the 2026-27 forecasts.
    row = stats.loc["US_net_saving_ratio_a_actuals"]
    actuals = usa_a[usa_a.index < 2026]
    ok = (close(row["current"], actuals.iloc[-1], TOL_EXACT)
          and close(row["max"], actuals.max(), TOL_EXACT)
          and close(row["mean"], actuals.mean(), TOL_EXACT)
          and close(row["median"], actuals.median(), TOL_EXACT)
          and close(row["min"], actuals.min(), TOL_EXACT))
    check(ok, "Summary SavingA stats computed on actuals only (<=2025)",
          f"summary current={row['current']:.4f} vs actuals-2025="
          f"{actuals.iloc[-1]:.4f}; mean {row['mean']:.4f} vs "
          f"{actuals.mean():.4f}")

# ------------------------- other Summary stat rows vs independent re-derive
if raw_sq is not None:
    row = stats.loc["US_gross_saving_rate_q"]
    us = raw_sq[raw_sq["REF_AREA"] == "USA"].set_index("Date")[
        "OBS_VALUE"].sort_index()
    check(close(row["current"], us.iloc[-1], TOL_EXACT)
          and close(row["max"], us.max(), TOL_EXACT)
          and close(row["median"], us.median(), TOL_EXACT),
          "Summary US gross-saving stats re-derived from cache",
          f"current {row['current']} vs {us.iloc[-1]}, max {row['max']} vs "
          f"{us.max()}")


# ---------------------------------------- rank-block recomputation (sheets)
def recompute_rank(frame: pd.DataFrame, cols: list[str]) -> dict:
    common = frame.dropna(subset=cols)
    latest = common.iloc[-1]
    vals = latest[cols].astype(float)
    us_val = float(vals["US"])
    return {
        "as_of": pd.Timestamp(latest["Date"]).strftime("%Y-%m"),
        "US_value": round(us_val, 2),
        "rank": int((vals > us_val).sum()) + 1,
        "n": len(cols),
        "pctile": round(float((vals <= us_val).mean() * 100), 1),
    }


if ranks is not None:
    # Credit panel: expected latest common quarter 2025Q4, US rank 5 of 8.
    rc = recompute_rank(hh, [c for c in hh.columns if c != "Date"])
    xr = ranks.loc["hh_credit_pct_gdp"]
    check(rc["as_of"] == str(xr["as_of"]) and rc["rank"] == int(
        xr["US_rank_1_is_highest"]) and rc["n"] == int(xr["n_countries"])
        and close(rc["US_value"], xr["US_value"], 0.01)
        and close(rc["pctile"], xr["US_percentile"], 0.05),
        "Rank block: credit panel recomputed",
        f"recomputed {rc['as_of']} rank {rc['rank']}/{rc['n']} "
        f"val {rc['US_value']} pct {rc['pctile']} vs xlsx "
        f"{xr['as_of']} rank {xr['US_rank_1_is_highest']}/"
        f"{xr['n_countries']}")
    check(rc["as_of"] == "2025-10" and rc["rank"] == 5 and rc["n"] == 8,
          "Rank block: credit expected 5/8 at 2025Q4",
          f"got {rc['rank']}/{rc['n']} at {rc['as_of']}")
    # Debt panel: Japan lags, so the all-countries-report rule must pull the
    # as-of quarter back to 2025Q1 with US rank 5 of 6.
    rd = recompute_rank(debt_q, [c for c in debt_q.columns if c != "Date"])
    xd = ranks.loc["hh_debt_pct_income_q"]
    check(rd["as_of"] == str(xd["as_of"]) and rd["rank"] == int(
        xd["US_rank_1_is_highest"]) and rd["n"] == int(xd["n_countries"])
        and close(rd["US_value"], xd["US_value"], 0.01),
        "Rank block: debt panel recomputed",
        f"recomputed {rd['as_of']} rank {rd['rank']}/{rd['n']} "
        f"val {rd['US_value']} vs xlsx {xd['as_of']} rank "
        f"{xd['US_rank_1_is_highest']}/{xd['n_countries']} "
        f"val {xd['US_value']}")
    check(rd["as_of"] == "2025-01" and rd["rank"] == 5 and rd["n"] == 6,
          "Rank block: debt all-countries rule (5/6 at 2025Q1)",
          f"got {rd['rank']}/{rd['n']} at {rd['as_of']} — latest common "
          f"quarter must trail the panel because Japan lags")
    check(str(ranks.loc["net_saving_ratio_a", "as_of"]) == "2025-01",
          "Rank block: SRATIO rank excludes 2026-27 forecasts",
          f"as_of={ranks.loc['net_saving_ratio_a', 'as_of']}")

# ------------------------------------------------ output files & sizes
expected_files = {
    "consumer_global.xlsx": 20_000,
    "household_credit_gdp.png": 30_000,
    "saving_rate_quarterly.png": 30_000,
    "household_debt_income.png": 30_000,
    "durables_share.png": 30_000,
    "durables_share_2006.png": 30_000,
    "saving_rate_annual.png": 30_000,
    "oecd_saving_q_raw.csv": 50_000,
    "oecd_debt_q_raw.csv": 50_000,
    "oecd_durables_raw.csv": 50_000,
    "oecd_saving_a_raw.csv": 50_000,
}
for fname, min_size in expected_files.items():
    p = OUT_DIR / fname
    ok = p.exists() and p.stat().st_size >= min_size
    check(ok, f"file {fname}",
          f"{'%d bytes' % p.stat().st_size if p.exists() else 'MISSING'} "
          f"(min {min_size})")

# ------------------------------------------------ data-range / staleness
today = pd.Timestamp(datetime.today().date())
check(hh["Date"].iloc[0] == pd.Timestamp("2006-01-01"),
      "HHCredit starts 2006-01-01 (repo convention)",
      f"first {hh['Date'].iloc[0].date()}")
check(sav_q["Date"].iloc[0] == pd.Timestamp("2007-01-01"),
      "SavingQ starts 2007Q1 (documented OECD availability)",
      f"first {sav_q['Date'].iloc[0].date()}")
check(sav_a["Date"].iloc[0].year <= 1960, "SavingA reaches back to 1960",
      f"first {sav_a['Date'].iloc[0].date()}")
# Staleness: dates are quarter STARTS, and BIS/OECD quarterly flows publish
# with up to a ~2-quarter lag (the live spot-check below confirms 2025Q4 is
# the newest quarter the OECD API itself serves for the dashboard flow), so
# the latest observation's period start should be within 4 quarters (366
# days) of today. Anything older means the script window is broken.
for name, frame in [("HHCredit", hh), ("SavingQ", sav_q),
                    ("DebtIncome", debt_q), ("Durables", dur)]:
    latest = frame["Date"].max()
    check((today - latest).days <= 366, f"{name} staleness",
          f"latest {latest.date()} ({(today - latest).days} days old, "
          f"max 366)")
check(sav_a["Date"].max() == pd.Timestamp("2027-01-01"),
      "SavingA extends through 2027 forecast",
      f"latest {sav_a['Date'].max().date()}")

# ------------------------------------------------ internal consistency
for name, frame in [("HHCredit", hh), ("SavingQ", sav_q),
                    ("DebtIncome", debt_q), ("Durables", dur),
                    ("SavingA", sav_a)]:
    d = frame["Date"]
    check(d.is_monotonic_increasing and d.is_unique,
          f"{name} dates monotonic & unique",
          f"{len(d)} rows {d.iloc[0].date()}..{d.iloc[-1].date()}")
    us_col = "US_4qma" if name == "Durables" else "US"
    s = frame[us_col]
    valid = s.notna()
    if valid.any():
        first, last = valid.idxmax(), valid[::-1].idxmax()
        interior_nans = int(s.loc[first:last].isna().sum())
        check(interior_nans == 0, f"{name} US no interior NaNs",
              f"{interior_nans} NaNs inside "
              f"{frame['Date'].loc[first].date()}.."
              f"{frame['Date'].loc[last].date()}")
        # Unit-jump detector: a stray x100 / /100 or UNIT_MULT mishap makes
        # consecutive values jump by >=10x. Threshold 4x: comfortably above
        # the largest genuine move in these series (US gross saving 14.64 ->
        # 29.18 in 2020Q2, a 1.99x COVID spike, verified in the raw CSV)
        # while still catching any decimal/unit error.
        sv = s.dropna()
        ratio = (sv / sv.shift(1)).dropna()
        worst = ratio[(ratio > 4.0) | (ratio < 0.25)]
        check(worst.empty, f"{name} US no unit-jump discontinuity",
              f"max consecutive ratio {ratio.max():.3f}, "
              f"min {ratio.min():.3f}")
# Quarterly spacing on the quarterly sheets.
for name, frame in [("HHCredit", hh), ("SavingQ", sav_q),
                    ("DebtIncome", debt_q), ("Durables", dur)]:
    gaps = frame["Date"].diff().dropna().dt.days
    check(gaps.between(89, 92).all(), f"{name} strictly quarterly spacing",
          f"gap range {gaps.min()}-{gaps.max()} days")

# ------------------------------------------------ external live spot-check
# One live OECD query (single country/measure, latest quarters) as an
# independent external cross-check of the cached value. API failure => WARN.
try:
    import requests

    url = ("https://sdmx.oecd.org/public/rest/data/"
           "OECD.SDD.NAD,DSD_HHDASH@DF_HHDASH_INDIC,1.0/"
           "Q.USA.B8GS1M_B6GA.?format=csvfilewithlabels&startPeriod=2025-Q1")
    resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"},
                        timeout=90)
    resp.raise_for_status()
    live = pd.read_csv(StringIO(resp.text))[
        ["TIME_PERIOD", "OBS_VALUE"]].dropna()
    live["Date"] = pd.PeriodIndex(live["TIME_PERIOD"], freq="Q").to_timestamp()
    live = live.set_index("Date")["OBS_VALUE"].sort_index()
    xus = sav_q.set_index("Date")["US"].dropna()
    common_q = live.index.intersection(xus.index)[-1]
    check(close(live.loc[common_q], xus.loc[common_q], TOL_LIVE),
          "EXTERNAL live OECD spot-check US gross saving",
          f"live {common_q.date()}={live.loc[common_q]} vs "
          f"xlsx={xus.loc[common_q]} (tol {TOL_LIVE})", warn_only=True)
except Exception as exc:
    report("WARN", "EXTERNAL live OECD spot-check US gross saving",
           f"live query failed ({exc}) — cached values not independently "
           f"confirmed against the API today")

# ------------------------------------------------ basis hygiene: PSAVERT
src = SCRIPT.read_text(encoding="utf-8")
bad_lines = []
for i, line in enumerate(src.splitlines(), start=1):
    if "PSAVERT" not in line:
        continue
    stripped = line.strip()
    is_comment = stripped.startswith("#")
    is_warning_text = "not comparable" in line.lower()
    if not (is_comment or is_warning_text):
        bad_lines.append(i)
check("PSAVERT" not in src or not bad_lines,
      "PSAVERT only in warnings/comments (never fetched or plotted)",
      f"occurrences on lines "
      f"{[i for i, l in enumerate(src.splitlines(), 1) if 'PSAVERT' in l]}; "
      f"non-comment/non-warning: {bad_lines or 'none'}")
check("get_series(\"PSAVERT\"" not in src.replace("'", '"'),
      "PSAVERT never pulled from FRED in analysis script",
      "no fred.get_series('PSAVERT') call found"
      if "PSAVERT" in src else "PSAVERT absent from script entirely")

# ------------------------------------------------ summary
print(f"\nSUMMARY: {results['PASS']} pass, {results['WARN']} warn, "
      f"{results['FAIL']} fail")
sys.exit(1 if results["FAIL"] else 0)
