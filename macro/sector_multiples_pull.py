# US sector EV/EBITDA and EV/EBIT, annual 1998-present, from Damodaran's
# industry-average archives (vebitda.xls + vebitdaYY.xls, NYU Stern).
#
# STAGE 1 of the sector roadmap — deliberately APPROXIMATE groundwork:
#   - Each vintage is Damodaran's early-JANUARY snapshot: prices as of the
#     update date, trailing financials through roughly the prior Q3 filings.
#     Vintage year YY = the data year (vebitda14.xls is dated 2015-01-05);
#     points are plotted at the January snapshot date (data year + 1). Do NOT
#     treat these as calendar year-end observations.
#   - Damodaran's ~95-industry classification is his own (Value Line-derived,
#     then his own scheme) — NOT GICS. The ~11 buckets here are a hand-built
#     GICS-like mapping and industry names drift across vintages.
#   - Modern files carry RATIOS only (no aggregate EV / EBITDA dollars), so a
#     true value-weighted sector multiple is impossible from this source.
#     Bucket aggregate = firm-count-weighted MEDIAN of member-industry
#     multiples — labeled "approx. (industry medians, firm-count weighted)"
#     on every chart.
#   - Where the modern format publishes both an "only positive EBITDA firms"
#     block and an "All firms" block, the ALL-FIRMS block is used (it is the
#     broader universe and exists for every two-block vintage); single-block
#     vintages have no choice. This is another small definitional seam.
#   - Stage 2 (future work): CapIQ-based quarterly per-sector multiples,
#     point-in-time constituents, true value-weighting — per-ticker plumbing
#     already exists in companies/scripts/fetch_multiple_history.py.
#
# DEFINITIONAL LEVEL BREAK: early vintages report "Value/EBITDA" where Value =
# MV of equity + total debt WITHOUT netting cash (firm value, not enterprise
# value); later vintages report true EV/EBITDA. The transition vintage is
# detected empirically from the column headers actually seen ("Value/EBITDA"
# vs "EV/EBITDA" — it lands at the 2010 vintage) and is marked with a dotted
# vline on every chart. Levels are NOT comparable across that line.
#
# FINANCIALS EXCLUDED: banks, insurers, brokerages, investment companies and
# financial-services industries are dropped entirely — EV/EBITDA is
# ill-defined for financial firms (debt is raw material, not capital
# structure; Damodaran himself blanks many of these cells). REITs are kept in
# the Real Estate bucket.
#
# INTENTIONAL EXCEPTION to the repo's 2006-01-01 start convention: the point
# of this series is the cross-cycle sector history (dot-com, GFC, COVID, AI),
# so primary charts run 1998-present, annual. A 2006+ companion overview
# (sector_ev_ebitda_2006.png) is emitted so it slots into the comparable-axis
# macro chart set. FRED is used ONLY for recession shading.
#
# Archive files are legacy .xls (engine xlrd) and are cached under
# macro/output/sector_multiples/raw/ (gitignored). Missing archive years are
# tolerated: skipped with a console note.

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
    resolve_output_dir,
)

CURRENT_URL = "https://pages.stern.nyu.edu/~adamodar/pc/datasets/vebitda.xls"
ARCHIVE_URL = "https://pages.stern.nyu.edu/~adamodar/pc/archives/vebitda{yy}.xls"
# The Stern pages 403 default python-requests UAs; present a browser UA.
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"
    )
}

OUT_DIR = resolve_output_dir(__file__, "sector_multiples")
(OUT_DIR / ".gitkeep").touch()
RAW_DIR = OUT_DIR / "raw"
RAW_DIR.mkdir(exist_ok=True)
(RAW_DIR / ".gitkeep").touch()

end = datetime.today()

BUCKETS = [
    "Technology", "Healthcare", "Energy", "Materials", "Industrials",
    "Consumer Discretionary", "Consumer Staples", "Telecom & Media",
    "Utilities", "Real Estate", "Other",
]
MARKET = "All industries (ex-financials)"

# Damodaran industry name (normalized: lower/strip/collapsed spaces) -> bucket.
# Built from the union of names actually observed across the 1998-current
# vintages, including drifted spellings ("publshing & newspapers",
# "insurance (prop/casualty" with the missing paren, etc.). Unmapped names
# fall to Other and are printed once per vintage so the map can be improved.
BUCKET_MAP = {
    # --- Technology ---
    "computer & peripherals": "Technology",
    "computers/peripherals": "Technology",
    "computer services": "Technology",
    "computer software": "Technology",
    "computer software & svcs": "Technology",
    "computer software/svcs": "Technology",
    "electronics": "Technology",
    "electronics (general)": "Technology",
    "electronics (consumer & office)": "Technology",
    "entertainment tech": "Technology",
    "foreign electron/entertn": "Technology",
    "foreign electronics": "Technology",
    "information services": "Technology",
    "it services": "Technology",
    "internet": "Technology",
    "internet software and services": "Technology",
    "office equip & supplies": "Technology",
    "office equip/supplies": "Technology",
    "office equipment & services": "Technology",
    "precision instrument": "Technology",
    "semiconductor": "Technology",
    "semiconductor cap eq": "Technology",
    "semiconductor cap equip": "Technology",
    "semiconductor equip": "Technology",
    "software (entertainment)": "Technology",
    "software (internet)": "Technology",
    "software (system & application)": "Technology",
    "telecom. equipment": "Technology",  # GICS: comms equipment sits in IT
    # --- Healthcare ---
    "biotechnology": "Healthcare",
    "drug": "Healthcare",
    "drugs (biotechnology)": "Healthcare",
    "drugs (pharmaceutical)": "Healthcare",
    "pharma & drugs": "Healthcare",
    "pharmacy services": "Healthcare",
    "healthcare equipment": "Healthcare",
    "healthcare facilities": "Healthcare",
    "healthcare info systems": "Healthcare",
    "healthcare information": "Healthcare",
    "heathcare information and technology": "Healthcare",  # sic, in source
    "healthcare products": "Healthcare",
    "healthcare services": "Healthcare",
    "healthcare support services": "Healthcare",
    "hospitals/healthcare facilities": "Healthcare",
    "med supp invasive": "Healthcare",
    "med supp non-invasive": "Healthcare",
    "medical services": "Healthcare",
    "medical supplies": "Healthcare",
    # --- Energy ---
    "alternate energy": "Energy",
    "canadian energy": "Energy",
    "coal": "Energy",
    "coal & related energy": "Energy",
    "coal/alternate energy": "Energy",
    "green & renewable energy": "Energy",
    "natural gas (div.)": "Energy",
    "natural gas (diversified": "Energy",  # sic, truncated in source
    "natural gas (diversified)": "Energy",
    "oil/gas (integrated)": "Energy",
    "oil/gas (production and exploration)": "Energy",
    "oil/gas distribution": "Energy",
    "oilfield services/equip.": "Energy",
    "oilfield svcs/equip.": "Energy",
    "petroleum (integrated)": "Energy",
    "petroleum (producing)": "Energy",
    "pipeline mlps": "Energy",
    # --- Materials ---
    "aluminum": "Materials",
    "copper": "Materials",
    "gold/silver mining": "Materials",
    "precious metals": "Materials",
    "metals & mining": "Materials",
    "metals & mining (div.)": "Materials",
    "steel": "Materials",
    "steel (general)": "Materials",
    "steel (integrated)": "Materials",
    "chemical (basic)": "Materials",
    "chemical (diversified)": "Materials",
    "chemical (specialty)": "Materials",
    "cement & aggregates": "Materials",
    "paper & forest products": "Materials",
    "paper/forest products": "Materials",
    "packaging & container": "Materials",
    # --- Industrials ---
    "aerospace/defense": "Industrials",
    "air transport": "Industrials",
    "building materials": "Industrials",
    "construction supplies": "Industrials",
    "construction": "Industrials",
    "engineering": "Industrials",
    "engineering & const": "Industrials",
    "engineering/construction": "Industrials",
    "heavy construction": "Industrials",
    "electrical equipment": "Industrials",
    "machinery": "Industrials",
    "heavy truck & equip": "Industrials",
    "heavy truck/equip makers": "Industrials",
    "metal fabricating": "Industrials",
    "environmental": "Industrials",
    "environmental & waste services": "Industrials",
    "industrial services": "Industrials",
    "business & consumer services": "Industrials",
    "human resources": "Industrials",
    "maritime": "Industrials",
    "shipbuilding & marine": "Industrials",
    "railroad": "Industrials",
    "transportation": "Industrials",
    "transportation (railroads)": "Industrials",
    "trucking": "Industrials",
    "trucking/transp. leasing": "Industrials",
    "diversified": "Industrials",      # conglomerates, per GICS treatment
    "diversified co.": "Industrials",
    # --- Consumer Discretionary ---
    "apparel": "Consumer Discretionary",
    "shoe": "Consumer Discretionary",
    "textile": "Consumer Discretionary",
    "auto & truck": "Consumer Discretionary",
    "auto parts": "Consumer Discretionary",
    "auto parts (oem)": "Consumer Discretionary",
    "auto parts (replacement)": "Consumer Discretionary",
    "automotive": "Consumer Discretionary",
    "tire & rubber": "Consumer Discretionary",
    "rubber& tires": "Consumer Discretionary",  # sic, in source
    "e-commerce": "Consumer Discretionary",
    "education": "Consumer Discretionary",
    "educational services": "Consumer Discretionary",
    "funeral services": "Consumer Discretionary",
    "furn./home furnishings": "Consumer Discretionary",
    "furn/home furnishings": "Consumer Discretionary",
    "home appliance": "Consumer Discretionary",
    "homebuilding": "Consumer Discretionary",
    "manuf. housing/rec veh": "Consumer Discretionary",
    "manuf. housing/rv": "Consumer Discretionary",
    "recreation": "Consumer Discretionary",
    "restaurant": "Consumer Discretionary",
    "restaurant/dining": "Consumer Discretionary",
    "hotel/gaming": "Consumer Discretionary",
    "retail (automotive)": "Consumer Discretionary",
    "retail (building supply)": "Consumer Discretionary",
    "retail (distributors)": "Consumer Discretionary",
    "retail (general)": "Consumer Discretionary",
    "retail (hardlines)": "Consumer Discretionary",
    "retail (internet)": "Consumer Discretionary",
    "retail (online)": "Consumer Discretionary",
    "retail (softlines)": "Consumer Discretionary",
    "retail (special lines)": "Consumer Discretionary",
    "retail automotive": "Consumer Discretionary",
    "retail building supply": "Consumer Discretionary",
    "retail store": "Consumer Discretionary",
    # --- Consumer Staples ---
    "beverage": "Consumer Staples",
    "beverage (alcoholic)": "Consumer Staples",
    "beverage (soft drink)": "Consumer Staples",
    "beverage (soft)": "Consumer Staples",
    "food processing": "Consumer Staples",
    "food wholesalers": "Consumer Staples",
    "farming/agriculture": "Consumer Staples",
    "grocery": "Consumer Staples",
    "retail (grocery and food)": "Consumer Staples",
    "retail/wholesale food": "Consumer Staples",
    "household products": "Consumer Staples",
    "tobacco": "Consumer Staples",
    "toiletries/cosmetics": "Consumer Staples",
    "drugstore": "Consumer Staples",
    # --- Telecom & Media ---
    "advertising": "Telecom & Media",
    "broadcasting": "Telecom & Media",
    "cable tv": "Telecom & Media",
    "entertainment": "Telecom & Media",
    "newspaper": "Telecom & Media",
    "publishing": "Telecom & Media",
    "publishing & newspapers": "Telecom & Media",
    "publshing & newspapers": "Telecom & Media",  # sic, in source
    "telecom (wireless)": "Telecom & Media",
    "telecom. services": "Telecom & Media",
    "telecom. utility": "Telecom & Media",
    "foreign telecom.": "Telecom & Media",
    "wireless networking": "Telecom & Media",
    # --- Utilities ---
    "electric util. (central)": "Utilities",
    "electric utility (east)": "Utilities",
    "electric utility (west)": "Utilities",
    "power": "Utilities",
    "utility (foreign)": "Utilities",
    "utility (general)": "Utilities",
    "utility (water)": "Utilities",
    "water utility": "Utilities",
    "natural gas utility": "Utilities",
    "natural gas (distrib.)": "Utilities",
    # --- Real Estate ---
    "r.e.i.t.": "Real Estate",
    "real estate (development)": "Real Estate",
    "real estate (general/diversified)": "Real Estate",
    "real estate (operations & services)": "Real Estate",
    "property management": "Real Estate",
    "retail (reits)": "Real Estate",
    # --- Other (explicit) ---
    "unclassified": "Other",
    "other": "Other",
}

# EV is ill-defined for financial firms (debt is raw material, not financing);
# these industries are dropped entirely rather than bucketed.
FINANCIAL_EXCLUDE = {
    "bank", "bank (canadian)", "bank (foreign)", "bank (midwest)",
    "bank (money center)", "banks (regional)", "thrift",
    "insurance (diversified)", "insurance (general)", "insurance (life)",
    "insurance (prop/cas.)", "insurance (prop/casualty",
    "insurance (prop/casualty)", "reinsurance",
    "brokerage & investment banking", "securities brokerage",
    "financial services", "financial svcs.", "financial svcs. (div.)",
    "financial svcs. (non-bank & insurance)",
    "investment (domestic)", "investment co.", "investment co. (domestic)",
    "investment co. (foreign)", "investment co. (income)",
    "investment co.(foreign)", "investment companies",
    "investments & asset management", "public/private equity",
}

# Aggregate/footer rows in the source files — not industries.
FOOTER_ROWS = {"market", "grand total", "total market",
               "total market (without financials)"}


def norm_name(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())


def norm_header(s: str) -> str:
    # The 2020-era files carry literal dedup artifacts in header cells
    # ("EV/EBITDA3", "EV/EBIT (1-t)5") — strip trailing digits.
    return re.sub(r"\d+$", "", norm_name(s)).strip()


def weighted_median(values: np.ndarray, weights: np.ndarray) -> float:
    """Firm-count-weighted median: value where cumulative weight crosses 50%."""
    order = np.argsort(values)
    v, w = values[order], weights[order]
    cum = np.cumsum(w)
    return float(v[np.searchsorted(cum, 0.5 * cum[-1])])


# --- Download (cache under raw/; tolerate missing archive years) -------------
# Archive tags run 98..(current year - 2): the Jan-YYYY snapshot of data year
# YYYY-1 gets archived as vebitda{YY-1}.xls once the next snapshot supersedes
# it, so the newest possible archive is always two calendar years back.
last_archive_year = datetime.today().year - 2
tags = (["98", "99"]
        + [f"{y % 100:02d}" for y in range(2000, last_archive_year + 1)]
        + ["current"])
missing: list[str] = []
for tag in tags:
    dest = RAW_DIR / f"vebitda_{tag}.xls"
    # Archives are immutable -> cache forever. The "current" file is refreshed
    # in place by Damodaran each January -> always re-fetch it, keeping the
    # cached copy only as a fallback if the fetch fails.
    if tag != "current" and dest.exists() and dest.stat().st_size > 1000:
        continue
    url = CURRENT_URL if tag == "current" else ARCHIVE_URL.format(yy=tag)
    try:
        resp = requests.get(url, headers=HEADERS, timeout=60)
        if resp.status_code != 200 or len(resp.content) < 1000:
            if dest.exists() and dest.stat().st_size > 1000:
                print(f"WARNING: re-fetch of vintage {tag} failed "
                      f"(HTTP {resp.status_code}); using cached copy")
            else:
                missing.append(tag)
            continue
        dest.write_bytes(resp.content)
    except requests.RequestException as exc:
        if dest.exists() and dest.stat().st_size > 1000:
            print(f"WARNING: re-fetch of vintage {tag} failed ({exc}); "
                  f"using cached copy")
        else:
            print(f"WARNING: download failed for vintage {tag}: {exc}")
            missing.append(tag)
if missing:
    print(f"NOTE: skipped missing/unfetchable vintages: {', '.join(missing)}")

# --- Parse (two format generations; fail soft per vintage) -------------------
# Legacy era (vintages 98-09): header row 0 on Sheet1, columns
#   [Industry Name, Number of Firms, Value/EBITDA, Value/EBIT, ...] where
#   Value = MV equity + TOTAL debt, cash NOT netted (firm value, not EV).
# Modern era (vintage 10+): "EV/..." columns; from ~2014 a preamble pushes the
#   header to row ~7-8; from ~2019 the sheet is "Industry Averages" with an
#   "only positive EBITDA firms" block and an "All firms" block (we take the
#   all-firms block = LAST occurrence of each column name).
records: list[dict] = []
parsed_tags: list[str] = []
seen_years: set[int] = set()
for tag in tags:
    path = RAW_DIR / f"vebitda_{tag}.xls"
    if not path.exists():
        continue
    try:
        xls = pd.ExcelFile(path, engine="xlrd")
        sheet = xls.sheet_names[0]
        for s in xls.sheet_names:
            if "average" in s.lower():
                sheet = s
                break
        raw = pd.read_excel(xls, sheet_name=sheet, header=None)

        # Vintage year = data year. For archives the filename tag IS the data
        # year (verified: vebitda14 is dated 2015-01-05, ..., vebitda24 is
        # dated 2025-01-05). The "Date updated" preamble is NOT trusted for
        # archives because vebitda17.xls carries a stale 2017-01-05 date
        # (copy of 16's); it is only used for the undated "current" file.
        date_year = None
        for i in range(min(10, len(raw))):
            if "date updated" in norm_name(raw.iloc[i, 0]):
                upd = pd.to_datetime(raw.iloc[i, 1], errors="coerce")
                if pd.notna(upd):
                    date_year = upd.year - 1
                break
        if tag == "current":
            if date_year is None:
                raise ValueError("current file has no 'Date updated' row")
            vintage_year = date_year
        else:
            vintage_year = (1900 if int(tag) >= 90 else 2000) + int(tag)
            if date_year is not None and date_year != vintage_year:
                print(f"NOTE: vintage {tag} 'Date updated' implies data year "
                      f"{date_year}; using filename year {vintage_year}")

        # Header row: first row whose col-0 is "Industry Name" / "Industry"
        # (vintage 14 says just "Industry"; 19 has a double space).
        hdr_row = None
        for i in range(min(15, len(raw))):
            if norm_name(raw.iloc[i, 0]) in ("industry name", "industry"):
                hdr_row = i
                break
        if hdr_row is None:
            raise ValueError("header row not found")
        headers = [norm_header(v) for v in raw.iloc[hdr_row]]

        if "value/ebitda" in headers:
            era = "firm_value"
            col_ebitda, col_ebit = "value/ebitda", "value/ebit"
        elif "ev/ebitda" in headers:
            era = "ev"
            col_ebitda, col_ebit = "ev/ebitda", "ev/ebit"
        else:
            raise ValueError(f"no EBITDA multiple column in headers {headers}")
        # LAST occurrence = the all-firms block where two blocks exist;
        # identical to the only block otherwise. Exact match, so
        # "ev/ebitdar&d" and "ev/ebit (1-t)" never collide.
        i_ebitda = len(headers) - 1 - headers[::-1].index(col_ebitda)
        i_ebit = (len(headers) - 1 - headers[::-1].index(col_ebit)
                  if col_ebit in headers else None)
        i_n = headers.index("number of firms")

        if vintage_year in seen_years:
            print(f"NOTE: vintage {tag} duplicates year {vintage_year}; skipped")
            continue

        body = raw.iloc[hdr_row + 1:]
        n_rows = 0
        unmapped: list[str] = []
        for _, row in body.iterrows():
            name = row.iloc[0]
            if pd.isna(name):
                continue
            key = norm_name(name)
            if key in FOOTER_ROWS:
                continue
            if key in FINANCIAL_EXCLUDE:
                continue
            bucket = BUCKET_MAP.get(key)
            if bucket is None:
                unmapped.append(str(name).strip())
                bucket = "Other"
            records.append({
                "vintage_year": vintage_year,
                "industry": str(name).strip(),
                "bucket": bucket,
                "n_firms": pd.to_numeric(row.iloc[i_n], errors="coerce"),
                "ev_ebitda": pd.to_numeric(row.iloc[i_ebitda], errors="coerce"),
                "ev_ebit": (pd.to_numeric(row.iloc[i_ebit], errors="coerce")
                            if i_ebit is not None else np.nan),
                "era": era,
            })
            n_rows += 1
        seen_years.add(vintage_year)
        parsed_tags.append(tag)
        print(f"vintage {vintage_year} ({tag}): {n_rows} industries, era={era}")
        if unmapped:
            print(f"  unmapped -> Other: {unmapped}")
    except Exception as exc:  # fail soft: one bad vintage shouldn't kill the run
        print(f"WARNING: could not parse vintage {tag}: {exc}")

ind = pd.DataFrame(records).sort_values(["vintage_year", "industry"])
ind = ind.reset_index(drop=True)
if ind.empty:
    raise RuntimeError("no vintages parsed — check downloads under raw/")

# Definitional level break: first vintage reporting true EV (found empirically
# from the headers, not hardcoded).
transition_year = int(ind.loc[ind["era"] == "ev", "vintage_year"].min())
print(f"Definitional transition (firm value -> EV): {transition_year} vintage")

# --- Aggregate: firm-count-weighted median of industry multiples -------------
years = sorted(ind["vintage_year"].unique())
agg_rows: list[dict] = []
for year in years:
    vin = ind[ind["vintage_year"] == year]
    for bucket in BUCKETS + [MARKET]:
        sub = vin if bucket == MARKET else vin[vin["bucket"] == bucket]
        row = {"vintage_year": year, "bucket": bucket}
        for metric in ("ev_ebitda", "ev_ebit"):
            m = sub.dropna(subset=[metric])
            if m.empty:
                row[metric] = np.nan
                continue
            w = m["n_firms"].fillna(1.0).clip(lower=1.0).to_numpy(float)
            row[metric] = weighted_median(m[metric].to_numpy(float), w)
        row["n_industries"] = len(sub)
        row["n_firms"] = sub["n_firms"].sum()
        agg_rows.append(row)
buckets_long = pd.DataFrame(agg_rows)

pivot_ebitda = buckets_long.pivot(index="vintage_year", columns="bucket",
                                  values="ev_ebitda")[BUCKETS + [MARKET]]
pivot_ebit = buckets_long.pivot(index="vintage_year", columns="bucket",
                                values="ev_ebit")[BUCKETS + [MARKET]]

latest_year = years[-1]
summary = pd.DataFrame({
    f"ev_ebitda_{latest_year}": pivot_ebitda.loc[latest_year],
    "ev_ebitda_median_1998plus": pivot_ebitda.median(),
    f"ev_ebit_{latest_year}": pivot_ebit.loc[latest_year],
    "ev_ebit_median_1998plus": pivot_ebit.median(),
})

xlsx_path = OUT_DIR / "sector_multiples.xlsx"
with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    ind[["vintage_year", "industry", "bucket", "n_firms",
         "ev_ebitda", "ev_ebit"]].to_excel(
        writer, sheet_name="Data_Industries", index=False)
    pd.concat({"EV_EBITDA": pivot_ebitda, "EV_EBIT": pivot_ebit},
              axis=1).to_excel(writer, sheet_name="Data_Buckets")
    summary.to_excel(writer, sheet_name="Summary")

# --- Charts -------------------------------------------------------------------
# X positions: the January snapshot date for each vintage (data year + 1),
# so the 1998 vintage plots at Jan-1999 and the 2025 vintage at Jan-2026.
def snap(year: int) -> pd.Timestamp:
    return pd.Timestamp(year + 1, 1, 5)


x_dates = [snap(y) for y in years]
trans_x = snap(transition_year)
fred = get_fred_client()  # FRED used only for USREC recession shading
recessions = get_recession_periods(fred, snap(years[0] - 1), end)
XLIM = (snap(years[0]) - pd.Timedelta(days=200),
        snap(years[-1]) + pd.Timedelta(days=200))
MAIN = "#1f3b73"
LIGHT = "#9ec5e8"
RED = "#c0392b"
APPROX = "approx. (industry medians, firm-count weighted)"
TRANS_LABEL = f"firm value → EV ({transition_year} vintage)"

# Chart 1 — small multiples: one panel per bucket + a 12th panel for the
# all-industries line. Shared y-limits clipped at 45x so cross-panel levels
# compare directly: recent peaks stay visible (Tech 34x in 2024, Healthcare
# 41x in 2021) while the legacy-era Real Estate REIT readings (47-133x, a
# firm-value/EBITDA artifact) and the sparse Other bucket (42-115x, 3
# vintages) run off-axis by design.
PANEL_YLIM = (0, 45)
fig, axes = plt.subplots(3, 4, figsize=(13, 10), sharex=True, sharey=True)
for ax, bucket in zip(axes.flat, BUCKETS + [MARKET]):
    sub = pivot_ebitda[bucket]
    for r_start, r_end in recessions:
        ax.axvspan(r_start, r_end, color="0.85", alpha=0.5, zorder=0)
    ax.axvline(trans_x, color=RED, linestyle=":", linewidth=1.0, zorder=1)
    ax.plot(x_dates, sub.values, color=MAIN, linewidth=1.6,
            marker="o", markersize=2.5)
    ax.set_title(bucket if bucket != MARKET else "All industries (ex-fin.)",
                 fontsize=10)
    ax.set_xlim(*XLIM)
    ax.set_ylim(*PANEL_YLIM)
    ax.grid(True, alpha=0.3)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
for ax in axes[:, 0]:
    ax.set_ylabel("EV/EBITDA (x)")
fig.suptitle(
    "US sector EV/EBITDA by vintage, 1998–2025 — Damodaran industry data, "
    f"{APPROX}\nDotted red line: {TRANS_LABEL} — levels not comparable across it. "
    "Gray bands: US recessions.",
    fontsize=11,
)
fig.tight_layout(rect=(0, 0, 1, 0.94))
fig.savefig(OUT_DIR / "sector_ev_ebitda.png", dpi=150)
plt.close(fig)

# Bucket colors for the overlay charts: Tech emphasized in the house main
# color; the rest thin in a muted qualitative cycle.
OVERLAY_COLORS = {
    "Healthcare": "#7f9cc4", "Energy": "#b08968", "Materials": "#8a9a5b",
    "Industrials": "#708090", "Consumer Discretionary": "#c497b2",
    "Consumer Staples": "#c9b26b", "Telecom & Media": "#7fb3a8",
    "Utilities": "#9ec5e8", "Real Estate": "#b3a2c7", "Other": "#c0c0c0",
}


def overlay_chart(pivot: pd.DataFrame, metric_label: str, title: str,
                  fname: str, year_filter=None, ylim=None) -> None:
    p = pivot if year_filter is None else pivot[pivot.index >= year_filter]
    xs = [snap(y) for y in p.index]
    fig, ax = plt.subplots(figsize=(11, 5.5))
    for k, (r_start, r_end) in enumerate(recessions):
        if r_end < xs[0]:
            continue
        ax.axvspan(r_start, r_end, color="0.85", alpha=0.5, zorder=0,
                   label="US recessions" if k == len(recessions) - 1 else None)
    if xs[0] <= trans_x <= xs[-1]:
        ax.axvline(trans_x, color=RED, linestyle=":", linewidth=1.2,
                   label=TRANS_LABEL)
    for bucket in BUCKETS:
        if bucket == "Technology":
            continue
        ax.plot(xs, p[bucket].values, color=OVERLAY_COLORS[bucket],
                linewidth=1.0, alpha=0.9, label=bucket)
    ax.plot(xs, p[MARKET].values, color="black", linewidth=1.8,
            linestyle="--", label="All industries (ex-fin.)")
    ax.plot(xs, p["Technology"].values, color=MAIN, linewidth=2.6,
            label="Technology")
    ax.set_xlim(xs[0] - pd.Timedelta(days=200), xs[-1] + pd.Timedelta(days=200))
    if ylim is not None:
        ax.set_ylim(*ylim)
    ax.set_title(f"{title}\n{APPROX}; Damodaran classification, financials excluded",
                 fontsize=11)
    ax.set_ylabel(f"{metric_label} (x)")
    ax.grid(True, alpha=0.3)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.legend(loc="upper left", frameon=False, fontsize=7.5, ncol=2)
    fig.tight_layout()
    fig.savefig(OUT_DIR / fname, dpi=150)
    plt.close(fig)


# Chart 2 — EV/EBITDA overview, all buckets on one axis. Clipped at 35x so
# the emphasized Tech line stays fully visible (34x peak, 2024 vintage);
# Healthcare's 2021 spike (41x), legacy Real Estate (47-133x) and sparse
# Other run off-axis by design.
overlay_chart(
    pivot_ebitda, "EV/EBITDA",
    "US sector EV/EBITDA, 1998–2025 vintages — Damodaran industry archives",
    "sector_ev_ebitda_overview.png", ylim=(0, 35),
)

# Chart 3 — EV/EBIT overview. Higher levels than EBITDA by construction;
# clipped at 45x — the COVID Energy EBIT collapse (102x, 2020 vintage),
# Healthcare 2021 (52x) and legacy Real Estate run off-axis by design.
overlay_chart(
    pivot_ebit, "EV/EBIT",
    "US sector EV/EBIT, 1998–2025 vintages — Damodaran industry archives",
    "sector_ev_ebit_overview.png", ylim=(0, 45),
)

# Chart 4 — 2006+ EV/EBITDA companion so this slots into the comparable-axis
# macro chart set (vintages 2005+, plotted at their Jan-2006+ snapshots).
overlay_chart(
    pivot_ebitda, "EV/EBITDA",
    "US sector EV/EBITDA, 2006–present snapshots — Damodaran industry archives",
    "sector_ev_ebitda_2006.png", year_filter=2005, ylim=(0, 35),
)

# --- Print summary -------------------------------------------------------------
latest = pivot_ebitda.loc[latest_year]
median_all = pivot_ebitda.median()
print()
print(f"Vintages parsed:        {len(years)} ({years[0]}-{years[-1]})")
print(f"Latest vintage:         {latest_year} (Jan-{latest_year + 1} snapshot)")
print(f"Transition vintage:     {transition_year} (firm value -> EV)")
print(f"Industry rows (long):   {len(ind)}")
med_n = ind.groupby("vintage_year")["industry"].count().median()
print(f"Median industries/vintage (ex-financials): {med_n:.0f}")
print(f"{'Bucket':<32}{'latest':>8}{'median 98+':>12}")
for bucket in BUCKETS + [MARKET]:
    lv = f"{latest[bucket]:.1f}x" if pd.notna(latest[bucket]) else "n/a"
    mv = f"{median_all[bucket]:.1f}x" if pd.notna(median_all[bucket]) else "n/a"
    print(f"{bucket:<32}{lv:>8}{mv:>12}")
print(f"Wrote {xlsx_path.name}, sector_ev_ebitda.png, "
      f"sector_ev_ebitda_overview.png, sector_ev_ebit_overview.png, "
      f"sector_ev_ebitda_2006.png to {OUT_DIR}")
