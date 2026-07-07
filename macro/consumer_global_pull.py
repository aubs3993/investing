# US consumer vs the world — structural cross-country panel: is the US consumer
# more credit-fueled, lower-saving, and durables-heavy than peers?
#
# Five panels:
#   1. Household credit, % of GDP (FRED mirrors of BIS series, quarterly, 2006+)
#   2. Household GROSS saving rate, % of gross disposable income (OECD household
#      dashboard, quarterly, 2007Q1+)
#   3. Household debt, % of gross disposable income (same OECD flow, 2007Q1+)
#   4. Durables share of household consumption (OECD QNA durability split,
#      current prices, quarterly, charted 1990+)
#   5. Annual NET household saving ratio, % of net disposable income (OECD
#      Economic Outlook, 1960+ incl. 2026-27 forecasts)
#
# INTENTIONAL EXCEPTION to the repo's 2006-01-01 start convention: panels 4
# and 5 chart much longer windows (1990+ and 1960+) because the point is the
# structural, multi-decade comparison. A companion 2006+ durables chart
# (durables_share_2006.png) is also emitted so it slots into the
# comparable-axis macro chart set. Panels 1-3 use the standard 2006/2007+
# window.
#
# BASIS WARNING — saving rates: panel 2 is a GROSS rate (% of gross disposable
# income); panel 5 is a NET rate (% of net disposable income). They are
# different accounting bases and are kept on separate charts/sheets. Neither
# is comparable to US PSAVERT (net, BEA basis) — do not mix.
#
# Data availability caveats (verified against live API responses 2026-07-07):
#   - Panel 2: Japan and Korea publish no quarterly gross saving rate in the
#     OECD dashboard flow (Korea was requested but returns no rows) — noted
#     on the chart.
#   - Panel 3: Korea returns no rows in this flow either; Japan lags one
#     quarter behind the others.
#   - Panel 4: Japan publishes no current-price durables split (chained
#     volumes only) — omitted and noted on the chart.
#   - Panel 5: GBR and FRA return no SRATIO rows in the EO flow; charted set
#     is USA/JPN/DEU/CAN/KOR/ITA/ESP. 2026-27 are OECD forecasts, shaded on
#     the chart; summary stats and ranks use actuals (through 2025) only.
#
# Recession shading: these are multi-country panels, but US NBER recessions
# are still useful timing markers — shaded and explicitly labeled
# "US recessions" so they aren't read as global.
#
# OECD fetch caching: each raw CSV response is cached in the output folder
# (gitignored). If a live fetch fails but a cache exists, the cache is used
# with a console warning; if neither is available the panel is skipped with a
# clear console message (graceful-skip pattern used elsewhere in the repo).
from datetime import datetime
from io import StringIO
from pathlib import Path
import sys

import matplotlib.pyplot as plt
import pandas as pd
import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.fred_helpers import (
    get_fred_client,
    get_recession_periods,
    pull_series,
    resolve_output_dir,
    series_stats,
    style_macro_chart,
)

fred = get_fred_client()

end = datetime.today()
chart_start = datetime(2006, 1, 1)
durables_chart_start = datetime(1990, 1, 1)  # documented exception, see header
annual_chart_start = datetime(1960, 1, 1)    # documented exception, see header
EO_FORECAST_START_YEAR = 2026  # EO 119: actuals through 2025, forecasts 2026-27

OUT_DIR = resolve_output_dir(__file__, "consumer_global")
(OUT_DIR / ".gitkeep").touch()

US_COLOR = "#1f3b73"
COUNTRY_COLORS = {
    "Japan": "#c0392b",
    "Germany": "#2ca02c",
    "France": "#e377c2",
    "UK": "#9467bd",
    "Canada": "#e67e22",
    "Korea": "#17becf",
    "Australia": "#8c564b",
    "Italy": "#bcbd22",
    "Spain": "#9ec5e8",
    "China": "#7f7f7f",
}
CODE_TO_NAME = {
    "USA": "US", "JPN": "Japan", "DEU": "Germany", "FRA": "France",
    "GBR": "UK", "CAN": "Canada", "KOR": "Korea", "AUS": "Australia",
    "ITA": "Italy", "ESP": "Spain", "CHN": "China",
}

# --- Panel 1: household credit, % of GDP (FRED mirrors of BIS series) -------
# All series exist by 2006 (China starts exactly 2006Q1), so the standard
# window applies with no lookback buffer needed (levels, no derived calcs).
credit_series = {
    "US": "QUSHAM770A",
    "Japan": "QJPHAM770A",
    "Germany": "QDEHAM770A",
    "UK": "QGBHAM770A",
    "Canada": "QCAHAM770A",
    "Korea": "QKRHAM770A",
    "Australia": "QAUHAM770A",
    "China": "QCNHAM770A",
}
df_credit = pull_series(fred, credit_series, chart_start, end)

# --- OECD SDMX fetches (with cache fallback / graceful skip) -----------------
OECD_BASE = "https://sdmx.oecd.org/public/rest/data/"
# sdmx.oecd.org works from scripts, but present a browser UA to be safe.
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"
    )
}
# Key forms verified 2026-07-07. HHDASH keys are FREQ.REF_AREA.MEASURE plus
# ONE trailing dot (the "..." multi-dot form 422s). The QNA durability key
# requires all 13 dimensions.
oecd_queries = {
    "saving_q": (
        "OECD.SDD.NAD,DSD_HHDASH@DF_HHDASH_INDIC,1.0/"
        "Q.USA+DEU+FRA+GBR+CAN+AUS+KOR.B8GS1M_B6GA."
        "?format=csvfilewithlabels&startPeriod=2007-Q1"
    ),
    "debt_q": (
        "OECD.SDD.NAD,DSD_HHDASH@DF_HHDASH_INDIC,1.0/"
        "Q.USA+JPN+DEU+GBR+CAN+KOR+AUS.LES1M_FD4."
        "?format=csvfilewithlabels&startPeriod=2007-Q1"
    ),
    "durables": (
        "OECD.SDD.NAD,DSD_NAMAIN1@DF_QNA_EXPENDITURE_DURABILITY,1.1/"
        "Q.Y.USA+CAN+FRA+DEU+GBR+ITA+ESP+KOR.S14.S1.P311+P31DC._Z._Z._T.XDC.V.N.T0117"
        "?format=csvfilewithlabels"
    ),
    "saving_a": (
        "OECD.ECO.MAD,DSD_EO@DF_EO,1.5/"
        "USA+JPN+DEU+CAN+KOR+ITA+ESP+GBR+FRA.SRATIO.A"
        "?format=csvfilewithlabels"
    ),
}
oecd_raw: dict[str, str | None] = {}
for name, query in oecd_queries.items():
    cache_path = OUT_DIR / f"oecd_{name}_raw.csv"
    try:
        resp = requests.get(OECD_BASE + query, headers=HEADERS, timeout=180)
        resp.raise_for_status()
        resp.encoding = "utf-8-sig"
        oecd_raw[name] = resp.text
        cache_path.write_text(resp.text, encoding="utf-8")
    except requests.RequestException as exc:
        if cache_path.exists():
            print(f"WARNING: OECD fetch '{name}' failed ({exc}); "
                  f"using cached {cache_path.name}")
            oecd_raw[name] = cache_path.read_text(encoding="utf-8")
        else:
            print(f"WARNING: OECD fetch '{name}' failed ({exc}) and no cache "
                  f"exists — skipping that panel and its chart")
            oecd_raw[name] = None

# --- Panel 2: quarterly gross saving rate ------------------------------------
df_saving_q = None
if oecd_raw["saving_q"] is not None:
    raw = pd.read_csv(StringIO(oecd_raw["saving_q"]))
    raw = raw[["REF_AREA", "TIME_PERIOD", "OBS_VALUE"]].dropna()
    raw["Date"] = pd.PeriodIndex(raw["TIME_PERIOD"], freq="Q").to_timestamp()
    piv = raw.pivot_table(index="Date", columns="REF_AREA", values="OBS_VALUE")
    df_saving_q = piv.rename(columns=CODE_TO_NAME).sort_index().reset_index()

# --- Panel 3: quarterly household debt, % of gross disposable income --------
df_debt_q = None
if oecd_raw["debt_q"] is not None:
    raw = pd.read_csv(StringIO(oecd_raw["debt_q"]))
    raw = raw[["REF_AREA", "TIME_PERIOD", "OBS_VALUE"]].dropna()
    raw["Date"] = pd.PeriodIndex(raw["TIME_PERIOD"], freq="Q").to_timestamp()
    piv = raw.pivot_table(index="Date", columns="REF_AREA", values="OBS_VALUE")
    df_debt_q = piv.rename(columns=CODE_TO_NAME).sort_index().reset_index()

# --- Panel 4: durables share of household consumption ------------------------
# share = P311 (durables, current prices) / P31DC (total household consumption,
# current prices), then a 4-quarter rolling mean to kill residual seasonality.
df_dur = None
if oecd_raw["durables"] is not None:
    raw = pd.read_csv(StringIO(oecd_raw["durables"]))
    raw = raw[["REF_AREA", "TRANSACTION", "TIME_PERIOD", "OBS_VALUE"]].dropna()
    raw["Date"] = pd.PeriodIndex(raw["TIME_PERIOD"], freq="Q").to_timestamp()
    piv = raw.pivot_table(index="Date", columns=["REF_AREA", "TRANSACTION"],
                          values="OBS_VALUE")
    shares = {}
    for code in sorted({c for c, _ in piv.columns}):
        if (code, "P311") in piv.columns and (code, "P31DC") in piv.columns:
            shares[CODE_TO_NAME.get(code, code)] = (
                piv[(code, "P311")] / piv[(code, "P31DC")] * 100
            )
    df_dur = pd.DataFrame(shares).sort_index()
    dur_countries = list(df_dur.columns)
    for c in dur_countries:
        df_dur[f"{c}_4qma"] = df_dur[c].rolling(4).mean()
    # Filter to the charted window so xlsx data and summary stats match the
    # chart (repo convention: stats over the charted window).
    df_dur = df_dur[df_dur.index >= pd.Timestamp(durables_chart_start)]
    df_dur = df_dur.reset_index()

# --- Panel 5: annual net household saving ratio (incl. EO forecasts) --------
df_saving_a = None
if oecd_raw["saving_a"] is not None:
    raw = pd.read_csv(StringIO(oecd_raw["saving_a"]))
    raw = raw[["REF_AREA", "TIME_PERIOD", "OBS_VALUE"]].dropna()
    raw["Date"] = pd.to_datetime(raw["TIME_PERIOD"].astype(int), format="%Y")
    piv = raw.pivot_table(index="Date", columns="REF_AREA", values="OBS_VALUE")
    df_saving_a = piv.rename(columns=CODE_TO_NAME).sort_index().reset_index()

# --- Summary: US stats per panel + US rank vs peers at latest common period --
summary_rows = {
    "US_hh_credit_pct_gdp": series_stats(df_credit["US"]),
}
if df_saving_q is not None:
    summary_rows["US_gross_saving_rate_q"] = series_stats(df_saving_q["US"])
if df_debt_q is not None:
    summary_rows["US_hh_debt_pct_income_q"] = series_stats(df_debt_q["US"])
if df_dur is not None:
    summary_rows["US_durables_share_4qma"] = series_stats(df_dur["US_4qma"])
if df_saving_a is not None:
    actual_mask = df_saving_a["Date"].dt.year < EO_FORECAST_START_YEAR
    summary_rows["US_net_saving_ratio_a_actuals"] = series_stats(
        df_saving_a.loc[actual_mask, "US"]
    )
summary = pd.DataFrame(summary_rows).T[["min", "max", "mean", "median", "current"]]

# US rank vs peers as of the latest quarter/year where ALL countries in the
# panel report (rank 1 = highest value; percentile = share of countries at or
# below the US value).
rank_rows = []
rank_specs = [
    ("hh_credit_pct_gdp", df_credit, None),
    ("gross_saving_rate_q", df_saving_q, None),
    ("hh_debt_pct_income_q", df_debt_q, None),
    ("durables_share_4qma", df_dur, "_4qma"),
    ("net_saving_ratio_a", df_saving_a, "actuals_only"),
]
for panel_name, frame, mode in rank_specs:
    if frame is None:
        continue
    if mode == "_4qma":
        cols = [c for c in frame.columns if c.endswith("_4qma")]
        display = {c: c.replace("_4qma", "") for c in cols}
    else:
        cols = [c for c in frame.columns if c != "Date"]
        display = {c: c for c in cols}
    sub = frame.copy()
    if mode == "actuals_only":
        sub = sub[sub["Date"].dt.year < EO_FORECAST_START_YEAR]
    common = sub.dropna(subset=cols)
    if common.empty:
        continue
    latest = common.iloc[-1]
    vals = latest[cols].astype(float)
    us_col = [c for c in cols if display[c] == "US"][0]
    us_val = float(vals[us_col])
    us_rank = int((vals > us_val).sum()) + 1
    rank_rows.append({
        "panel": panel_name,
        "as_of": pd.Timestamp(latest["Date"]).strftime("%Y-%m"),
        "US_value": round(us_val, 2),
        "US_rank_1_is_highest": us_rank,
        "n_countries": len(cols),
        "US_percentile": round(float((vals <= us_val).mean() * 100), 1),
    })
ranks = pd.DataFrame(rank_rows)

# --- Write xlsx (one Data sheet per available panel + Summary) ---------------
xlsx_path = OUT_DIR / "consumer_global.xlsx"
with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
    df_credit.to_excel(writer, sheet_name="Data_HHCredit", index=False)
    if df_saving_q is not None:
        df_saving_q.to_excel(writer, sheet_name="Data_SavingQ", index=False)
    if df_debt_q is not None:
        df_debt_q.to_excel(writer, sheet_name="Data_DebtIncome", index=False)
    if df_dur is not None:
        df_dur.to_excel(writer, sheet_name="Data_Durables", index=False)
    if df_saving_a is not None:
        df_saving_a.to_excel(writer, sheet_name="Data_SavingA", index=False)
    summary.to_excel(writer, sheet_name="Summary")
    # US-vs-peers rank block below the stats block on the same Summary sheet.
    ranks.to_excel(writer, sheet_name="Summary", index=False,
                   startrow=len(summary) + 3)

# --- Charts -------------------------------------------------------------------
recessions_2006 = get_recession_periods(fred, chart_start, end)
recessions_1990 = get_recession_periods(fred, durables_chart_start, end)
recessions_1960 = get_recession_periods(fred, annual_chart_start, end)
XLIM_2006 = (pd.Timestamp(chart_start), pd.Timestamp(end))
XLIM_1990 = (pd.Timestamp(durables_chart_start), pd.Timestamp(end))

# Chart 1 — household credit, % of GDP. US thick house-blue, peers thinner
# distinct colors, China dashed (its catch-up path is the story).
fig, ax = plt.subplots(figsize=(11, 5))
for i, (r0, r1) in enumerate(recessions_2006):
    ax.axvspan(r0, r1, color="0.85", alpha=0.5, zorder=0,
               label="US recessions" if i == 0 else None)
for country in credit_series:
    if country == "US":
        ax.plot(df_credit["Date"], df_credit["US"], color=US_COLOR,
                linewidth=2.4, label="US", zorder=5)
    else:
        ax.plot(df_credit["Date"], df_credit[country],
                color=COUNTRY_COLORS[country], linewidth=1.2,
                linestyle="--" if country == "China" else "-",
                label=country)
ax.set_xlim(*XLIM_2006)
style_macro_chart(
    ax,
    title="Household credit, % of GDP — US vs peers (BIS), 2006–present",
    ylabel="% of GDP",
    ylim=(0, 135),
)
ax.legend(loc="upper right", frameon=False, ncol=3, fontsize=8)
fig.tight_layout()
fig.savefig(OUT_DIR / "household_credit_gdp.png", dpi=150)
plt.close(fig)

# Chart 2 — quarterly GROSS household saving rate.
if df_saving_q is not None:
    fig, ax = plt.subplots(figsize=(11, 5))
    for i, (r0, r1) in enumerate(recessions_2006):
        ax.axvspan(r0, r1, color="0.85", alpha=0.5, zorder=0,
                   label="US recessions" if i == 0 else None)
    for country in [c for c in df_saving_q.columns if c != "Date"]:
        if country == "US":
            ax.plot(df_saving_q["Date"], df_saving_q["US"], color=US_COLOR,
                    linewidth=2.4, label="US", zorder=5)
        else:
            ax.plot(df_saving_q["Date"], df_saving_q[country],
                    color=COUNTRY_COLORS[country], linewidth=1.2, label=country)
    ax.set_xlim(*XLIM_2006)
    style_macro_chart(
        ax,
        title="Household gross saving rate, % of gross disposable income, "
              "2007–present",
        ylabel="% of gross disposable income",
        hlines=[{"y": 0.0}],
    )
    ax.legend(loc="upper right", frameon=False, ncol=3, fontsize=8)
    ax.text(0.01, 0.02,
            "Gross basis — not comparable to US PSAVERT (net). "
            "Japan & Korea: no quarterly series in OECD dashboard.",
            transform=ax.transAxes, fontsize=7, color="0.4")
    fig.tight_layout()
    fig.savefig(OUT_DIR / "saving_rate_quarterly.png", dpi=150)
    plt.close(fig)
else:
    print("SKIP: saving_rate_quarterly.png (no OECD data and no cache)")

# Chart 3 — household debt, % of gross disposable income.
if df_debt_q is not None:
    fig, ax = plt.subplots(figsize=(11, 5))
    for i, (r0, r1) in enumerate(recessions_2006):
        ax.axvspan(r0, r1, color="0.85", alpha=0.5, zorder=0,
                   label="US recessions" if i == 0 else None)
    for country in [c for c in df_debt_q.columns if c != "Date"]:
        if country == "US":
            ax.plot(df_debt_q["Date"], df_debt_q["US"], color=US_COLOR,
                    linewidth=2.4, label="US", zorder=5)
        else:
            ax.plot(df_debt_q["Date"], df_debt_q[country],
                    color=COUNTRY_COLORS[country], linewidth=1.2, label=country)
    ax.set_xlim(*XLIM_2006)
    style_macro_chart(
        ax,
        title="Household debt, % of gross disposable income, 2007–present",
        ylabel="% of gross disposable income",
    )
    ax.legend(loc="upper right", frameon=False, ncol=3, fontsize=8)
    ax.text(0.01, 0.02,
            "Korea: not published in this OECD flow. Japan lags one quarter.",
            transform=ax.transAxes, fontsize=7, color="0.4")
    fig.tight_layout()
    fig.savefig(OUT_DIR / "household_debt_income.png", dpi=150)
    plt.close(fig)
else:
    print("SKIP: household_debt_income.png (no OECD data and no cache)")

# Chart 4 — durables share of household consumption (4-quarter MA), 1990+
# primary window (documented exception) plus a 2006+ companion.
if df_dur is not None:
    dur_note = ("Current-price durables / total household consumption, "
                "4-qtr rolling mean. Japan: no current-price split published.")
    for fname, x0, suffix in [
        ("durables_share.png", XLIM_1990[0], "1990–present"),
        ("durables_share_2006.png", XLIM_2006[0], "2006–present"),
    ]:
        recs = recessions_1990 if x0 == XLIM_1990[0] else recessions_2006
        fig, ax = plt.subplots(figsize=(11, 5))
        for i, (r0, r1) in enumerate(recs):
            ax.axvspan(r0, r1, color="0.85", alpha=0.5, zorder=0,
                       label="US recessions" if i == 0 else None)
        for country in dur_countries:
            if country == "US":
                ax.plot(df_dur["Date"], df_dur["US_4qma"], color=US_COLOR,
                        linewidth=2.4, label="US", zorder=5)
            else:
                ax.plot(df_dur["Date"], df_dur[f"{country}_4qma"],
                        color=COUNTRY_COLORS[country], linewidth=1.2,
                        label=country)
        ax.set_xlim(x0, pd.Timestamp(end))
        style_macro_chart(
            ax,
            title=f"Durables share of household consumption, {suffix}",
            ylabel="% of household consumption",
        )
        ax.legend(loc="upper right", frameon=False, ncol=3, fontsize=8)
        ax.text(0.01, 0.02, dur_note, transform=ax.transAxes, fontsize=7,
                color="0.4")
        fig.tight_layout()
        fig.savefig(OUT_DIR / fname, dpi=150)
        plt.close(fig)
else:
    print("SKIP: durables_share.png / durables_share_2006.png "
          "(no OECD data and no cache)")

# Chart 5 — annual NET household saving ratio, long history incl. Japan.
# 2026-27 OECD forecasts are kept on the chart but shaded and labeled, so the
# forward view is visible without being mistaken for history.
if df_saving_a is not None:
    fig, ax = plt.subplots(figsize=(11, 5))
    for i, (r0, r1) in enumerate(recessions_1960):
        ax.axvspan(r0, r1, color="0.85", alpha=0.5, zorder=0,
                   label="US recessions" if i == 0 else None)
    fc_start = pd.Timestamp(datetime(EO_FORECAST_START_YEAR, 1, 1))
    fc_end = pd.Timestamp(df_saving_a["Date"].max())
    ax.axvspan(fc_start, fc_end, color="#f0e6c8", alpha=0.6, zorder=0,
               label="OECD forecast")
    for country in [c for c in df_saving_a.columns if c != "Date"]:
        if country == "US":
            ax.plot(df_saving_a["Date"], df_saving_a["US"], color=US_COLOR,
                    linewidth=2.4, label="US", zorder=5)
        else:
            ax.plot(df_saving_a["Date"], df_saving_a[country],
                    color=COUNTRY_COLORS[country], linewidth=1.2, label=country)
    ax.set_xlim(pd.Timestamp(annual_chart_start), fc_end)
    style_macro_chart(
        ax,
        title="Net household saving ratio, % of net disposable income "
              "(annual), 1960–2027",
        ylabel="% of net disposable income",
        hlines=[{"y": 0.0}],
    )
    ax.legend(loc="upper right", frameon=False, ncol=3, fontsize=8)
    ax.text(0.01, 0.02,
            "Net basis — different from the gross quarterly panel. "
            "UK & France: no SRATIO series in OECD EO flow.",
            transform=ax.transAxes, fontsize=7, color="0.4")
    fig.tight_layout()
    fig.savefig(OUT_DIR / "saving_rate_annual.png", dpi=150)
    plt.close(fig)
else:
    print("SKIP: saving_rate_annual.png (no OECD data and no cache)")

# --- Console summary ----------------------------------------------------------
print(f"Chart window (panels 1-3): {chart_start.date()} -> {end.date()}")
print(f"HH credit rows:            {len(df_credit)}")
us_credit = df_credit["US"].dropna()
print(f"Latest US credit/GDP:      {us_credit.iloc[-1]:.1f}% "
      f"(2006+ max {us_credit.max():.1f}%)")
for c in ("Korea", "Germany", "China"):
    s = df_credit[c].dropna()
    print(f"Latest {c + ' credit/GDP:':<20}{s.iloc[-1]:.1f}%")
if df_saving_q is not None:
    print(f"SavingQ rows:              {len(df_saving_q)}")
    print(f"Latest US gross saving:    "
          f"{df_saving_q['US'].dropna().iloc[-1]:.1f}%")
if df_debt_q is not None:
    print(f"DebtIncome rows:           {len(df_debt_q)}")
    print(f"Latest US debt/income:     "
          f"{df_debt_q['US'].dropna().iloc[-1]:.1f}%")
if df_dur is not None:
    print(f"Durables rows (1990+):     {len(df_dur)}")
    print(f"Latest US durables share:  "
          f"{df_dur['US_4qma'].dropna().iloc[-1]:.1f}% (4-qtr MA)")
if df_saving_a is not None:
    us_a = df_saving_a.loc[df_saving_a["Date"].dt.year
                           < EO_FORECAST_START_YEAR, "US"].dropna()
    print(f"SavingA rows:              {len(df_saving_a)}")
    print(f"US net saving ratio 2025:  {us_a.iloc[-1]:.1f}%")
print("US rank vs peers (1 = highest), latest common period per panel:")
for row in rank_rows:
    print(f"  {row['panel']:<24}{row['as_of']}: {row['US_value']} "
          f"(rank {row['US_rank_1_is_highest']}/{row['n_countries']}, "
          f"pctile {row['US_percentile']})")
written = ["consumer_global.xlsx", "household_credit_gdp.png"]
if df_saving_q is not None:
    written.append("saving_rate_quarterly.png")
if df_debt_q is not None:
    written.append("household_debt_income.png")
if df_dur is not None:
    written += ["durables_share.png", "durables_share_2006.png"]
if df_saving_a is not None:
    written.append("saving_rate_annual.png")
print(f"Wrote {', '.join(written)} to {OUT_DIR}")
