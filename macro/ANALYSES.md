# Macro analyses — research notes and interpretation guide

Reference notes for the long-horizon analysis modules added July 2026: aggregate US
valuation multiples, the credit-cycle gauge, and the consumer pull-forward gauge.
Every data claim in this document was verified against the primary source
(FRED API, live dataset downloads, or the cited paper's published record) on
2026-07-07 by a multi-agent research pass with adversarial claim verification.
Refresh cadence: all modules are plain `python macro/<script>.py` runs.

---

## 1. Valuation multiples family

### Modules

| Script | Series | Window | Frequency |
|---|---|---|---|
| `shiller_pe_pull.py` | Trailing P/E, CAPE, TR CAPE | 1871+ | monthly |
| `nipa_pe_pull.py` | Z.1/NIPA macro P/E (4 scope/tax variants) | 1947+ | quarterly |
| `ev_multiples_pull.py` | EV/EBIT, EV/EBITDA, FCF yields, MC/distributed-FCF | 1952Q4+ | quarterly |
| `valuation_multiple_drivers_pull.py` | D&A share, tax rate, yields vs real 10y, margin, intangibles | 1947+ | quarterly |
| `sp500_long_history_pull.py` | S&P price + total-return index, drawdowns | 1871+ | monthly |
| `sector_multiples_pull.py` | Damodaran sector EV/EBITDA & EV/EBIT (approx.) | 1998+ | annual |

### Construction (top-down, all FRED)

The 50-year quarterly multiples are built from macro accounts, not company data:

- **EV** = market value of nonfinancial-corporate equities (`NCBCEL`, Z.1 B.103,
  1945Q4+) + debt securities & loans (`BCNSDODNS`) − broad liquid assets
  (`BOGZ1FL104001005Q`).
- **EBIT** = NIPA Table 1.14 pretax profits with IVA & CCAdj (`A463RC1Q027SBEA`)
  + net interest (`B471RC1Q027SBEA`); **EBITDA** adds consumption of fixed
  capital (`B456RC1Q027SBEA`).
- **Unlevered FCF** = EBIT − taxes (`B465RC1Q027SBEA`) + CFC − Z.1 total capex
  (`BOGZ1FA105050005Q`); **levered FCF** = after-tax profits (`W328RC1Q027SBEA`)
  + CFC − capex. A distribution-based companion (dividends `BOGZ1FA106121075Q`
  minus net equity issuance `NCBCEBQ027S` = dividends + buybacks) is always
  positive and closer to what shareholders actually receive.
- NIPA/Z.1 flows are SAAR: the trailing-year denominator is a trailing
  4-quarter rolling **mean** (8-quarter for the noisy FCF legs). FCF legs are
  charted as **yields** because aggregate FCF is genuinely negative in the
  mid-1970s and around 2000.

**Known-bad series (verified discontinued/removed — do not use):**
`MVEONWMVBSNNCB` (ends 2017Q4), `HSTCMDODNS` (2015Q2; use `CMDEBT`),
Wilshire indices (removed from FRED entirely), `NCBTSDQ027S`/`NCBPFAQ027S`
(2017Q4), `FODSP` (discontinued 2023Q3). FRED `SP500` is licensed trailing-10y
only. ICE BofA OAS series serve only a trailing ~3-year window since April 2026
(the repo archives HY OAS in `macro/output/credit_spreads/`). NAR existing-home
sales (`EXHOSLUSM495S`) was truncated the same way in 2025.

### How to read the multiple charts

Level comparisons across decades are treacherous; the charts are percentile
machines, not fair-value rulers:

1. **The mean is not stationary.** Corporate tax cuts (2018 especially), the
   falling cost of debt 1982–2021, rising intangible intensity, and sector
   drift toward asset-light business models all shift the *justified* multiple
   upward across the sample. Naive reversion-to-the-1975-average is the single
   biggest interpretive error with 50-year valuation charts.
2. **Decompose before concluding.** EV/EBITDA = (EV/GVA) ÷ (EBITDA margin on
   GVA). A high multiple with a high margin is double-extended; a high multiple
   on a depressed margin is partly denominators. The drivers script plots both.
3. **EBIT vs EBITDA wedge** is mechanical: the D&A share of EBITDA. Compare
   EV/EBIT growth vs EV/EBITDA growth against that share before attributing
   meaning.
4. **After-tax vs pretax**: TCJA (2018) permanently raised after-tax profits
   with no pretax valuation change — the NIPA P/E pretax variant isolates this.
   Note (verified correction): the *effective* federal corporate rate shows
   **no** 1986–88 step-down — TRA86 base-broadening raised the effective rate
   ~20%→25% even as the statutory rate fell 46%→34%. Only 2018 is a clean step.
5. **Scope**: the NCB aggregate covers all US nonfinancial corporations
   (public + private) and domestic production only, while listed-market indices
   capitalize global earnings. Levels are NOT comparable to S&P 500 multiples
   (~mid-teens EV/EBITDA on company data vs ~14x on 2026Q1 macro data is a
   coincidence of scale, not agreement). Use each series only against its own
   history.
6. **Trailing P/E in recessions lies rich**: FAS 142 goodwill write-downs made
   Q4-2008 S&P as-reported EPS negative and pushed trailing P/E to ~124x at the
   2009 *bottom*. The NIPA P/E (economic profits, no write-downs) is the
   Siegel-consistent cross-check; operating EPS exists only from 1988.

### Predictive-power context (verified citations)

- Asness (AQR, 2012): deciles of starting Shiller P/E since 1926 give
  near-monotonic next-decade real returns; the richest decile (CAPE 25.1–46.1)
  averaged +0.5%/yr real (range −6.1% to +6.3%). (A commonly quoted 0.9%/yr
  figure is the 9th decile, not the 10th.)
- Hussman's MarketCap/GVA (introduced May 2015) reports ~0.9 correlation with
  subsequent 10–12y S&P returns — in-sample, overlapping horizons; motivating
  evidence, not a tested forecast model.
- Smolyansky (Fed, FEDS 2023-041, "End of an era"): 1989–2019 real corporate
  profit growth was substantially driven by falling interest expense and tax
  rates — machinery for interpreting the EBIT→net-income wedge.
- Gray-Vogel and Loughran-Wellman support EV/EBITDA cross-sectionally (stock
  selection); they say nothing directly about aggregate market timing.

### Segment groundwork (roadmap)

- **Stage 1 (built)**: `sector_multiples_pull.py` — Damodaran annual industry
  archives (1998+), mapped to ~11 GICS-like buckets, firm-count-weighted
  medians. Deliberately approximate: his classification is not GICS, modern
  files publish ratios only (no aggregation numerators), and the early files
  use firm value (cash not netted) vs EV later — a level break marked on the
  charts. Financials excluded (EV ill-defined).
- **Stage 2 (deferred)**: CapIQ quarterly sector aggregates via a
  `templates/sector_multiple_fetcher.xlsx` clone of the multiple-history
  pattern (IQ_TEV / IQ_EBITDA / IQ_GICS_SECTOR over constituents). ~2004+
  realistically. Unverified: whether `IQ_CONSTITUENTS` accepts historical
  as-of dates (determines point-in-time vs survivorship-biased membership).
  Needs a 5-minute live-Excel probe before building.
- **Stage 3 (deferred)**: sector FCF multiples and true point-in-time GICS
  membership.
- The Z.1/NIPA macro approach **cannot** go below the sector-aggregate level —
  Z.1 has no industry split of nonfinancial corporates.

---

## 2. Credit cycle (`credit_cycle_pull.py`)

### Design

Two blocks, deliberately kept separate (they have opposite signs and a ~2-year
phase offset — averaging them produces mush):

- **STRESS** (high = credit tight): z-scores of Baa−Aaa spread level (1919+),
  its 12m change, the Fed's Excess Bond Premium (1973+), and SLOOS C&I
  net-tightening (1990+). Components enter as they begin; equal weight over
  what's available.
- **FROTH** (high = exuberance; leads reversals by ~2 years): z of inverted
  spread, aggregate credit impulse, inverted SLOOS, household credit impulse.
  Charted advanced 24 months against the stress block.

Credit impulse (Biggs-Mayer-Pick): change in the trailing-year borrowing flow,
scaled by GDP — from Z.1 *transaction* series (`BOGZ1FA384104005Q` aggregate;
`BOGZ1FA154104005Q` households), not deltas of debt levels (charge-offs sit in
"other volume changes" and would understate gross borrowing exactly in stress
periods).

**Deliberate omission**: the Greenwood-Hanson high-yield-share-of-issuance leg
("issuer quality"). The SIFMA source sits behind a registration form with
personal-use licensing and unclear history depth. Adding it manually later:
HY/(IG+HY) of gross nonconvertible issuance, annual, z-scored, lagged 2 years.

### Lead/lag context (verified citations, with corrections)

- Gilchrist & Zakrajšek (AER 2012): EBP innovations predict economic activity;
  the Fed republishes EBP monthly (revising full history each time — always
  re-download, never append).
- López-Salido, Stein & Zakrajšek (QJE 2017, *published* magnitudes): elevated
  credit-market sentiment at t−2 predicts, over t..t+1, roughly −1.2pp
  cumulative real GDP per capita and +0.8pp unemployment (postwar sample).
  (Larger figures often quoted — −4.2pp GDP — are from the 2015 working paper.)
- Greenwood & Hanson (2013): HY share of issuance leads credit returns ~2y.
- SLOOS: a one-SD business lending-standards shock lowers output ~0.5pp with a
  ~4-quarter lag (Richmond Fed WP 24-07). SLOOS measures *changes* in
  standards, is bank-only, and misses the growing private-credit share.
- BIS credit-to-GDP gap: slow (2–5y horizon) early-warning gauge; belongs on a
  separate chart, not in a monthly composite. Not on FRED; recompute from
  `QUSPAM770A` with a one-sided HP filter (λ=400k) if wanted.

### Pairing conventions

Equities pairing uses the Shiller-derived monthly total-return index and
drawdowns from `sp500_long_history_pull.py` (run it first; the credit script
reads its CSV). Rates pairing: GS10 + fed funds. The classic reads: stress
peaks precede fed-funds peaks→cuts; EBP > ~1 has historically flagged most
major drawdown windows; SLOOS tightening leads EPS contractions by ~2–4Q.

### Consumer-credit bridge (shared with pull-forward gauge)

To avoid double-counting between analyses: the credit composite owns business
credit, spreads, SLOOS and the *household* impulse (CI_HH); the pull-forward
gauge owns revolving credit growth, card delinquencies, saving rate, and the
*consumer-credit* impulse (CI_CC). Both scripts compute and chart CI_HH/CI_CC
with identical formulas (flows over 4q-mean disposable income) as labeled
bridge series. Interpretation rule: pull-forward elevated + CI_CC strongly
positive ⇒ the spending is credit-financed (borrowed demand); pull-forward
elevated + CI_CC negative ⇒ payback underway.

---

## 3. Consumer pull-forward (`consumer_pullforward_pull.py`, `consumer_global_pull.py`)

### What pull-forward is and how it's measured

Intertemporal substitution: buying durables earlier than otherwise because of
incentives, expected price increases, or windfalls. Four measurement designs
generalize (the script implements 1, 3, 4 and an ex-ante survey leg):

1. **Trend deviation + cumulative-excess integral** (SF Fed excess-savings
   design applied to quantities): fit a frozen pre-event trend, cumulate the
   gap; payback horizon = when the cumulative gap re-crosses zero.
2. **Cross-sectional exposure diff-in-diff** (Mian-Sufi): the identification
   gold standard, needs micro data — used here only to interpret published
   magnitudes.
3. **Deadline bunching windows** (−6m/+18m around known policy dates).
4. **Durables stock-adjustment**: purchases vs depreciation-implied replacement
   demand; a stock above trend implies future payback even after flows
   normalize.

### Episode library (verified magnitudes)

| Episode | Pull-forward | Payback |
|---|---|---|
| Cash for Clunkers (Jul–Aug 2009) | ~360k purchases (Mian-Sufi, QJE 2012) | almost fully reversed by Mar 2010 (~7 months) |
| Germany Abwrackprämie 2009 | only ~30% of subsidized purchases pulled forward (Klößner-Pfeifer) | scheme design changes the split |
| Japan VAT Apr 1997 (3%→5%) | durables/storables surge; implied IES = 0.21 (Cashin-Unayama, REStat 2016) | sharp, within quarters |
| Japan VAT Apr 2014 (5%→8%) | ~+0.7pp GDP pulled into FY2013 (consensus think-tank estimate — not peer-reviewed) | symmetric payback FY2014 |
| Germany VAT Jan 2007 (16%→19%) | anticipatory durables buying (Carare-Danninger, IMF) | 2007 payback |
| COVID goods boom 2020–21 | real durables ~+25-30% above pre-COVID trend | **never fully paid back** — partial permanent level shift; cumulative gap does not re-cross zero |
| Tariff front-running Mar–Apr 2025 | imports surge (Q1-25 GDP drag); auto SAAR ~17.9M Mar-25 | payback into H2-25; Cox 2026 forecast 15.8M (−2.4%) cites policy payback context |
| EV credit expiry Sep 2025 | EV unit bunching (438k→234k quarterly) | ~1 quarter |

2018 washer tariffs (Flaaen-Hortaçsu-Tintelnot, AER) is a *price pass-through*
study — importer stockpiling, not consumer pull-forward; cite for prices only.

### The US-vs-world context (`consumer_global_pull.py`)

The cross-country panels quantify whether the US consumer is structurally more
able and willing to pull demand forward: household credit % GDP (BIS mirrors on
FRED, `Q{ISO2}HAM770A`), gross saving rate and household debt % income (OECD
SDMX API), durables share of consumption (OECD QNA durability split), annual
net saving ratios (OECD EO `SRATIO` — note `SAVH` is a currency *level*, a
verification catch). Supporting literature: Havránek-Horváth-Iršová-Rusnák
(JIE 2015) meta-analysis of 2,735 elasticity-of-intertemporal-substitution
estimates across 104 countries — EIS is systematically higher in richer,
higher-asset-participation countries, i.e. US-style consumers substitute more
across time. Current readings: the US is mid-pack on household leverage
(68% of GDP, down from 98% peak), near the bottom of peers on saving, and near
the top on durables share — consistent with high pull-forward capacity.

Basis warnings: OECD quarterly saving rate is GROSS, the annual `SRATIO` and
US `PSAVERT` are NET — never mix on one panel. Japan lacks a quarterly OECD
saving rate and current-price durables split.

### Ex-ante leg (UMich SCA buying conditions)

The one *leading* pull-forward indicator: the share citing "buy in advance of
rising prices" as why now is a good time to buy durables/vehicles/houses
(tables 36/38/42 at data.sca.isr.umich.edu, monthly 1978+, quarterly 1960–77).
Peaks ~45–53 in 1979–80; spiked ~20-25 around Dec-2024/2025 (tariff
anticipation — the spike slightly predates the tariff announcements). Survey
moved phone→web in 2024; treat level breaks accordingly. Multiple mentions
allowed — columns can exceed 100 summed; never renormalize.

### Known gaps / future work

- BNPL is essentially invisible in `REVOLSL` and has no recurring public time
  series (CFPB annual snapshots; NY Fed SCE 3×/yr) — credit-financed spending
  is slightly understated.
- Monthly PCE subcategories (motor vehicles etc.) exist in BEA underlying-
  detail tables via the **BEA API** (free key; `U20405`/`U20406`, 1959+ monthly)
  — not on FRED. Deferred: needs a BEA_API_KEY in `.env`; would sharpen episode
  windows from quarterly to monthly.
- Existing-home sales history: FRED series truncated to trailing window (2025);
  Zillow (2008+) / Redfin (2012+) are the free companions; an append-archive of
  the FRED window would need refreshing at least every ~12 months.
- Auto replacement baseline is a hardcoded constant (fleet 289M × 4.5%
  scrappage ≈ 13.0M SAAR, S&P Global Mobility May-2025) — refresh annually.

---

## 4. Fragile-source registry

| Source | Fragility | Mitigation |
|---|---|---|
| Shiller `ie_data.xls` | blob URL `?ver=` token rotates per update; GoDaddy hosting | scrape link from shillerdata.com at run time; fallback URL constant; raw file cached in output |
| Fed EBP CSV | full-history revisions monthly; URL can move with FEDS Note updates | re-download every run; archived copy in output folder |
| ICE BofA OAS on FRED | trailing ~3y window since Apr 2026 | repo's own archive (credit_spreads module) |
| FRED `SP500` | trailing 10y (license) | Shiller monthly via `sp500_long_history_pull.py` |
| OECD SDMX API | keyless, throttled; key-path syntax picky (one trailing dot) | raw CSV cache + graceful skip in consumer_global |
| UMich SCA archive | ~1-2 month lag; header quirks (`<br>`); 2024 methodology break | CSV cache + graceful skip |
| S&P DJI EPS sheet | 403 to all scripted requests (Akamai) | manual browser download if ever needed (operating vs as-reported P/E, 1988+) |
| SIFMA issuance stats | registration wall + personal-use license | HY-share froth leg deliberately omitted; manual annual add possible |
| Damodaran archives | sporadic 404s per vintage; format drift; his own industry taxonomy | per-era parsers, soft-fail per vintage, raw cache |
| NAR existing-home sales | FRED series truncated 2025 | flagged; Zillow/Redfin companions if needed |
