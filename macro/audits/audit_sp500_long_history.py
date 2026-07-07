# Independent audit of macro/sp500_long_history_pull.py and
# macro/output/sp500_long_history/.
#
# Re-parses the cached raw ie_data.xls itself (positional columns, header=None,
# own month decoding via round(d*100) % 100 — deliberately NOT the analysis
# script's f"{d:.2f}" string-split approach) and re-derives the nominal
# total-return index with an explicit recursive loop (not cumprod), drawdowns,
# and the annualized return. Does NOT import or exec sp500_long_history_pull.py;
# reads sp500_monthly.csv / sp500_long_history.xlsx only to compare.
#
# sp500_monthly.csv is a DATA CONTRACT consumed by the credit-cycle script:
# columns Date,P,D,E,CPI,TR,TR_real,drawdown — checked exactly here.
#
# Tolerances:
#   - TR spot re-derivations: 1e-6 relative. Same raw file, so only float
#     round-trip noise is expected; a construction difference (off-by-one
#     dividend month, D vs D/12, wrong base) is orders of magnitude larger.
#   - Era drawdown benchmarks (1932-06 -81.8%, 2009-03 -49.0%, 2020-03 -18.9%
#     on monthly-average prices): 0.5pp absolute — Shiller occasionally
#     back-revises dividends/prices slightly.
#   - Annualized nominal TR: 9.39%/yr +/- 0.05pp (each new month moves the
#     155-year annualization only ~0.001pp, so this is stable).
#   - Staleness: last CSV month within 70 days of today (monthly series; the
#     current partial month is intentionally dropped for missing D).
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

OUT_DIR = MACRO_DIR / "output" / "sp500_long_history"
CSV = OUT_DIR / "sp500_monthly.csv"
XLSX = OUT_DIR / "sp500_long_history.xlsx"
RAW_XLS = OUT_DIR / "ie_data.xls"

TODAY = pd.Timestamp(datetime.today().date())

RESULTS = []


def check(status: str, name: str, detail: str) -> None:
    RESULTS.append(status)
    print(f"{status} {name}: {detail}")


# ---------------------------------------------------------------------------
# 1. Output file existence + non-trivial size
# ---------------------------------------------------------------------------
EXPECTED_FILES = {
    "sp500_monthly.csv": 50_000,
    "sp500_long_history.xlsx": 50_000,
    "ie_data.xls": 500_000,
    "sp500_tr_log.png": 20_000,
    "sp500_drawdown.png": 20_000,
    "sp500_drawdown_2006.png": 20_000,
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

if not CSV.exists() or not RAW_XLS.exists():
    print("\nSummary: cannot continue without sp500_monthly.csv + ie_data.xls")
    sys.exit(1)

# ---------------------------------------------------------------------------
# 2. Independent re-parse of the raw ie_data.xls (positional, header=None)
# ---------------------------------------------------------------------------
raw = pd.read_excel(RAW_XLS, sheet_name="Data", header=None, engine="xlrd")
# Row 7 carries the short header labels; data begins at row 8. Verify the
# anchor instead of assuming, so a layout change fails loudly.
if str(raw.iloc[7, 0]).strip() != "Date" or abs(float(raw.iloc[8, 0]) - 1871.01) > 1e-9:
    check("FAIL", "raw layout anchor",
          f"expected header 'Date' at row 7 / 1871.01 at row 8, got "
          f"{raw.iloc[7, 0]!r} / {raw.iloc[8, 0]!r}")
    print("\nSummary: raw layout changed; aborting")
    sys.exit(1)

body = raw.iloc[8:, [0, 1, 2, 3, 4, 9]].copy()
body.columns = ["DateF", "P", "D", "E", "CPI", "RealTRPrice"]
for c in body.columns:
    body[c] = pd.to_numeric(body[c], errors="coerce")
body = body.dropna(subset=["DateF"]).reset_index(drop=True)

# Own month decoding: integer arithmetic, not string formatting.
cents = (body["DateF"] * 100).round().astype(int)
body["Date"] = pd.to_datetime(
    {"year": cents // 100, "month": cents % 100, "day": 1})
if not body["Date"].dt.month.between(1, 12).all():
    check("FAIL", "raw date decoding", "month outside 1..12 after decode")
    sys.exit(1)

# Partial trailing month(s): price present, D NaN. Keep a copy for check 6,
# then trim to the last D-valid row (mirrors the documented contract: TR must
# be defined on every CSV row).
raw_tail = body.dropna(subset=["P"]).iloc[-1]
last_d_idx = body["D"].last_valid_index()
mine = body.loc[:last_d_idx].reset_index(drop=True)
if mine["D"].isna().any() or mine["P"].isna().any():
    check("FAIL", "raw interior NaNs",
          "interior NaN in P or D within the D-valid range of ie_data.xls")
else:
    check("PASS", "raw interior NaNs",
          f"P and D fully populated {mine['Date'].iloc[0].date()}.."
          f"{mine['Date'].iloc[-1].date()} ({len(mine)} months)")

# Recursive TR construction: buy at last month's average price, collect one
# month (1/12) of the annual-rate dividend, reinvest. Explicit loop.
P = mine["P"].to_numpy()
D = mine["D"].to_numpy()
tr = np.empty(len(mine))
tr[0] = 100.0
for i in range(1, len(mine)):
    tr[i] = tr[i - 1] * (P[i] + D[i] / 12.0) / P[i - 1]
mine["TR"] = tr
mine["TR_real"] = 100.0 * mine["RealTRPrice"] / mine["RealTRPrice"].iloc[0]
mine["drawdown"] = mine["TR"] / np.maximum.accumulate(tr) - 1.0
mine = mine.set_index("Date")

# ---------------------------------------------------------------------------
# 3. Load the CSV under audit
# ---------------------------------------------------------------------------
csv = pd.read_csv(CSV)
CONTRACT_COLS = ["Date", "P", "D", "E", "CPI", "TR", "TR_real", "drawdown"]
if list(csv.columns) == CONTRACT_COLS:
    check("PASS", "CSV contract columns",
          f"exact column names/order {CONTRACT_COLS}")
else:
    check("FAIL", "CSV contract columns",
          f"got {list(csv.columns)}, contract requires {CONTRACT_COLS}")
csv["Date"] = pd.to_datetime(csv["Date"])
xl = csv.set_index("Date")

n_nan = int(csv.isna().sum().sum())
if n_nan == 0:
    check("PASS", "CSV contract NaNs", f"zero NaNs in {len(csv)} rows x "
          f"{len(csv.columns)} cols")
else:
    check("FAIL", "CSV contract NaNs",
          f"{n_nan} NaN cell(s): {csv.isna().sum()[csv.isna().sum() > 0].to_dict()}")

# Strictly monthly, monotone, unique, first-of-month, 1871-01..2026-06.
d = csv["Date"]
mono = d.is_monotonic_increasing
dupes = int(d.duplicated().sum())
day1 = (d.dt.day == 1).all()
per = pd.PeriodIndex(d, freq="M")
missing_m = len(pd.period_range(per[0], per[-1], freq="M")) - len(per)
exp_start, exp_end = pd.Timestamp("1871-01-01"), pd.Timestamp("2026-06-01")
if (mono and dupes == 0 and day1 and missing_m == 0
        and d.iloc[0] == exp_start and d.iloc[-1] == exp_end):
    check("PASS", "CSV contract date axis",
          f"{len(d)} rows, strictly monthly {d.iloc[0].date()}.."
          f"{d.iloc[-1].date()}, no dupes/gaps")
else:
    check("FAIL", "CSV contract date axis",
          f"monotonic={mono}, dupes={dupes}, all day-1={day1}, "
          f"missing months={missing_m}, range {d.iloc[0].date()}.."
          f"{d.iloc[-1].date()} (expected {exp_start.date()}..{exp_end.date()})")

# ---------------------------------------------------------------------------
# 4. TR spot re-derivations at 5 months spanning the sample, 1e-6 rel
# ---------------------------------------------------------------------------
REL_TOL = 1e-6
SPOTS = [pd.Timestamp("1871-06-01"), pd.Timestamp("1930-06-01"),
         pd.Timestamp("1987-12-01"), pd.Timestamp("2009-03-01"),
         xl.index[-1]]
for spot in SPOTS:
    label = f"TR re-derivation {spot.date()}"
    if spot not in xl.index or spot not in mine.index:
        check("FAIL", label, "date missing from CSV or independent parse")
        continue
    mv, xv = mine.loc[spot, "TR"], xl.loc[spot, "TR"]
    rel = abs(xv - mv) / abs(mv)
    if rel <= REL_TOL:
        check("PASS", label, f"csv {xv:,.4f} vs recursive re-derivation "
              f"{mv:,.4f} (rel {rel:.2e})")
    else:
        check("FAIL", label, f"csv {xv:,.4f} vs recursive re-derivation "
              f"{mv:,.4f} (rel {rel:.2e} > {REL_TOL:.0e})")

# TR_real and raw passthrough columns (P, D, E, CPI) at the same spots.
bad_cols = []
for spot in SPOTS:
    if spot not in xl.index or spot not in mine.index:
        continue
    for col in ["P", "D", "E", "CPI", "TR_real"]:
        mv, xv = mine.loc[spot, col], xl.loc[spot, col]
        if abs(xv - mv) / max(abs(mv), 1e-12) > REL_TOL:
            bad_cols.append(f"{col}@{spot.date()}: csv {xv} vs raw {mv}")
if bad_cols:
    check("FAIL", "raw passthrough columns", "; ".join(bad_cols))
else:
    check("PASS", "raw passthrough columns",
          "P/D/E/CPI/TR_real match the independent raw parse at all 5 spot "
          f"months (rel <= {REL_TOL:.0e})")

# ---------------------------------------------------------------------------
# 5. Drawdown re-derivation + era benchmarks (abs tol 0.5pp)
# ---------------------------------------------------------------------------
ABS_TOL_PP = 0.5
BENCHMARKS = [
    ("Depression trough", "1929-01-01", "1940-12-31", "1932-06-01", -81.8),
    ("GFC trough", "2007-10-01", "2009-12-31", "2009-03-01", -49.0),
    ("COVID trough", "2020-01-01", "2020-12-31", "2020-03-01", -18.9),
]
my_dd = mine["drawdown"]
for name, w0, w1, exp_date, exp_pct in BENCHMARKS:
    win = my_dd.loc[w0:w1]
    t_date, t_val = win.idxmin(), win.min() * 100.0
    csv_win = xl["drawdown"].loc[w0:w1]
    c_date, c_val = csv_win.idxmin(), csv_win.min() * 100.0
    ok_mine = (t_date == pd.Timestamp(exp_date)
               and abs(t_val - exp_pct) <= ABS_TOL_PP)
    ok_csv = (c_date == t_date and abs(c_val - t_val) <= 0.01)
    if ok_mine and ok_csv:
        check("PASS", name,
              f"{c_val:.1f}% at {c_date.date()} in CSV; independent "
              f"re-derivation {t_val:.1f}% (expected ~{exp_pct}% at {exp_date})")
    elif ok_csv:
        check("WARN", name,
              f"CSV matches my re-derivation ({c_val:.1f}% at {c_date.date()}) "
              f"but both differ from the ~{exp_pct}% @ {exp_date} benchmark — "
              "possible Shiller data revision")
    else:
        check("FAIL", name,
              f"CSV {c_val:.1f}% at {c_date.date()} vs independent "
              f"{t_val:.1f}% at {t_date.date()} (expected ~{exp_pct}% at "
              f"{exp_date})")

# Global max drawdown should BE the Depression trough.
g_date, g_val = xl["drawdown"].idxmin(), xl["drawdown"].min() * 100.0
if g_date == pd.Timestamp("1932-06-01") and abs(g_val - (-81.8)) <= ABS_TOL_PP:
    check("PASS", "global max drawdown",
          f"{g_val:.1f}% at {g_date.date()} (matches -81.8% @ 1932-06)")
else:
    check("FAIL", "global max drawdown",
          f"{g_val:.1f}% at {g_date.date()}, expected -81.8% at 1932-06-01")

# Internal consistency: CSV drawdown column == TR/cummax(TR)-1 recomputed from
# the CSV's OWN TR column (catches a drawdown computed on a different basis).
recomputed = xl["TR"] / xl["TR"].cummax() - 1.0
max_abs = float((xl["drawdown"] - recomputed).abs().max())
if max_abs <= 1e-9:
    check("PASS", "drawdown internal consistency",
          f"CSV drawdown == TR/cummax(TR)-1 from CSV's own TR "
          f"(max abs diff {max_abs:.1e})")
else:
    check("FAIL", "drawdown internal consistency",
          f"CSV drawdown deviates from TR/cummax(TR)-1 by up to {max_abs:.2e}")

# ---------------------------------------------------------------------------
# 6. Annualized nominal TR ~ 9.39%/yr
# ---------------------------------------------------------------------------
n_months = len(mine) - 1
my_ann = (mine["TR"].iloc[-1] / mine["TR"].iloc[0]) ** (12.0 / n_months) - 1.0
EXP_ANN = 0.0939
if abs(my_ann - EXP_ANN) <= 0.0005:
    check("PASS", "annualized nominal TR",
          f"{my_ann:.4%} over {n_months / 12:.1f} yrs (expected ~9.39%)")
else:
    check("FAIL", "annualized nominal TR",
          f"{my_ann:.4%}, expected ~9.39% +/- 0.05pp — TR construction or "
          "sample-window error")

# xlsx Summary sheet should carry the same figure.
if XLSX.exists():
    summary = pd.read_excel(XLSX, sheet_name="Summary", index_col=0)
    if "tr_annualized_nominal" in summary.index:
        xv = float(summary.loc["tr_annualized_nominal", "current"])
        if abs(xv - my_ann) <= 1e-6:
            check("PASS", "xlsx Summary annualized TR",
                  f"Summary current {xv:.4%} matches re-derivation")
        else:
            check("FAIL", "xlsx Summary annualized TR",
                  f"Summary {xv:.4%} vs re-derivation {my_ann:.4%}")
    else:
        check("FAIL", "xlsx Summary annualized TR",
              "row 'tr_annualized_nominal' missing from Summary sheet")
    xdata = pd.read_excel(XLSX, sheet_name="Data")
    if len(xdata) == len(csv):
        check("PASS", "xlsx/CSV row parity",
              f"Data sheet and CSV both have {len(csv)} rows")
    else:
        check("FAIL", "xlsx/CSV row parity",
              f"Data sheet {len(xdata)} rows vs CSV {len(csv)}")

# ---------------------------------------------------------------------------
# 7. Last-row values and partial-month exclusion
# ---------------------------------------------------------------------------
last = csv.iloc[-1]
if abs(last["P"] - 7450.03) <= 0.01 and last["Date"] == pd.Timestamp("2026-06-01"):
    check("PASS", "last CSV row",
          f"{last['Date'].date()} P={last['P']:.2f} (expected 2026-06, "
          "7450.03)")
else:
    check("FAIL", "last CSV row",
          f"{last['Date'].date()} P={last['P']:.4f}, expected 2026-06-01 / "
          "7450.03")

raw_tail_date = raw_tail["Date"]
if (raw_tail_date == pd.Timestamp("2026-07-01")
        and abs(raw_tail["P"] - 7483.23) <= 0.01):
    if raw_tail_date not in set(csv["Date"]):
        check("PASS", "partial month excluded",
              f"raw file has partial {raw_tail_date.date()} "
              f"(P={raw_tail['P']:.2f}, D=NaN) and it is correctly absent "
              "from the CSV")
    else:
        check("FAIL", "partial month excluded",
              f"partial month {raw_tail_date.date()} (D=NaN) leaked into the "
              "CSV — TR undefined there")
else:
    check("WARN", "partial month excluded",
          f"raw tail is {raw_tail_date.date()} P={raw_tail['P']:.2f}; expected "
          "partial 2026-07 P=7483.23 (raw file may have been refreshed since "
          "the analysis run)")

# ---------------------------------------------------------------------------
# 8. Staleness + no unit jumps
# ---------------------------------------------------------------------------
age_days = (TODAY - xl.index[-1]).days
if age_days <= 70:
    check("PASS", "staleness", f"last CSV month {xl.index[-1].date()} is "
          f"{age_days} days old (<= 70; current partial month is "
          "intentionally dropped)")
else:
    check("FAIL", "staleness", f"last CSV month {xl.index[-1].date()} is "
          f"{age_days} days old (> 70) — output looks stale")

jump_fails = []
worst = ("", 0.0)
for col in ["P", "CPI", "TR", "TR_real"]:
    s = xl[col]
    if (s <= 0).any():
        jump_fails.append(f"{col}: non-positive values")
        continue
    lg = np.abs(np.log(s / s.shift(1))).dropna()
    # Largest true monthly moves in this history are ~30% (1931-32); a unit
    # break (x10, x100) or splice error would be >= log(2).
    if lg.max() >= 0.7:
        jump_fails.append(
            f"{col}: |m/m log change| {lg.max():.2f} at {lg.idxmax().date()}")
    if lg.max() > worst[1]:
        worst = (f"{col} @ {lg.idxmax().date()}", lg.max())
if jump_fails:
    check("FAIL", "unit discontinuities", "; ".join(jump_fails))
else:
    check("PASS", "unit discontinuities",
          f"max |m/m log change| {worst[1]:.2f} ({worst[0]}) < 0.7 — no unit "
          "breaks")

dd = xl["drawdown"]
if (dd <= 1e-12).all() and dd.max() > -1e-9:
    check("PASS", "drawdown bounds",
          f"drawdown always <= 0 and touches 0 at new highs "
          f"(min {dd.min():.4f})")
else:
    check("FAIL", "drawdown bounds",
          f"drawdown outside (-1, 0]: min {dd.min()}, max {dd.max()}")

# ---------------------------------------------------------------------------
# 9. EXTERNAL cross-check: 1929-32 price-only crash vs TR drawdown direction
# ---------------------------------------------------------------------------
# Known benchmark: on Shiller monthly-average prices, the Sep-1929 -> Jun-1932
# price-only decline is ~ -84.8% (daily-close basis is the famous -89%, but
# this dataset is monthly averages). Dividends cushion total return, so the
# TR-basis max drawdown (-81.8%) must be SHALLOWER (less negative) than
# price-only by a few pp. If the TR drawdown came out DEEPER than price-only,
# the TR construction would be subtracting/misplacing dividends.
p = mine["P"]
dd_price = p / p.cummax() - 1.0
p_win = dd_price.loc["1929-01-01":"1935-12-31"]
p_date, p_val = p_win.idxmin(), p_win.min() * 100.0
tr_val = my_dd.loc["1929-01-01":"1935-12-31"].min() * 100.0
if abs(p_val - (-84.8)) <= 1.0:
    check("PASS", "external 1929-32 price-only crash",
          f"independent price-only max drawdown {p_val:.1f}% at "
          f"{p_date.date()} vs known Shiller monthly benchmark ~-84.8%")
else:
    check("WARN", "external 1929-32 price-only crash",
          f"price-only max drawdown {p_val:.1f}% at {p_date.date()} differs "
          "from the ~-84.8% benchmark by >1pp (possible data revision)")
if p_val < tr_val - 1.0:  # price-only must be deeper (more negative)
    check("PASS", "dividend-cushion direction",
          f"price-only {p_val:.1f}% is deeper than TR {tr_val:.1f}% — "
          "dividends cushion total return as expected")
else:
    check("WARN", "dividend-cushion direction",
          f"TR drawdown {tr_val:.1f}% is NOT shallower than price-only "
          f"{p_val:.1f}% — suggests a TR construction error (dividends "
          "misapplied)")

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
n_pass = RESULTS.count("PASS")
n_warn = RESULTS.count("WARN")
n_fail = RESULTS.count("FAIL")
print(f"\nSummary: {n_pass} PASS, {n_warn} WARN, {n_fail} FAIL "
      f"({len(RESULTS)} checks)")
sys.exit(1 if n_fail else 0)
