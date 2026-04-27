# Driver Research Playbook

When the user asks to research drivers for a ticker, follow this exact sequence. Do not skip steps. Don't ask clarifying questions before starting — make reasonable defaults and flag them at the end.

## Inputs

User provides: ticker (e.g., `AAPL`).

## Prerequisites

Verify before starting:
1. `companies/output/<TICKER>/<TICKER>_model.xlsx` exists. If not, instruct: `python -m shared.fetch_capiq <TICKER>` then `python -m shared.fetch_broker_estimates <TICKER>`.
2. Both `_CapIQ_Data` and `_Broker_Data` tabs in the model have recent timestamps. If `_Broker_Data` is empty, instruct user to run the broker fetch.
3. Notion MCP is available (try a basic search). If not, run the playbook anyway and note Notion steps were skipped.

## Step 1 — Extract historicals

```
python -m companies.scripts.extract_historicals <TICKER> --format markdown
```

Read the output. Note especially:
- 3-year revenue CAGR
- Latest YoY revenue (decelerating from CAGR? accelerating?)
- Margin trajectory (gross, EBITDA, EBIT — direction and magnitude)
- CapEx intensity (% of revenue)
- D&A vs. CapEx ratio (growth mode vs. steady state)
- Effective tax rate

## Step 2 — Extract broker estimates

```
python -m companies.scripts.extract_broker_estimates <TICKER> --format markdown
```

Read the output. Note especially:
- Implied revenue growth FY1 / FY2 / FY3 — this is the consensus view
- Implied margins FY1 — this is what the street has baked in
- Number of analysts covering (low coverage = thinner consensus, more variance possible)
- Dispersion (high/low spread) — proxy for analyst disagreement
- Average price target & recommendation — overall sentiment

This is your **framing**, not your answer. Consensus tells you what's already priced in. Your job is to figure out where you differ and why.

## Step 3 — Read Notion company page

Use Notion MCP `notion-search` for the ticker or company name.

If the page exists, read the full content. Pay special attention to:
- **Historical Analysis section** — durable knowledge built up over prior research sessions. This is gold.
- **Prior Driver Research entries** (dated) — what you (or prior-you) thought last time. Useful for tracking how views have evolved.
- **Thesis / View** section if present — the user's high-level take.

If no page exists, search for sector-level pages and read those for context. Plan to create a company page in Step 8.

## Step 4 — Read recent SEC filings

Use `web_fetch` to retrieve from SEC EDGAR. Target the most recent 10-K and most recent 10-Q for the ticker.

**EDGAR URL gotcha:** The inline XBRL viewer URL prefix `/ix?doc=` must be stripped to the direct `/Archives/` path for content retrieval. Convert before fetching.

For 10-K: focus on Item 7 (MD&A), Item 1 (Business). Skim Risk Factors only for items relevant to driver assumptions.

For 10-Q: most recent quarter results vs. prior year, forward guidance, management commentary on trajectory.

Take notes: what's management saying about growth drivers? Margin direction? Capex plans? Capital allocation?

## Step 5 — Update Notion Historical Analysis section

Before proposing forward drivers, refine the durable Historical Analysis section in Notion.

Open (or create) the company page. Locate the "Historical Analysis" section (create if not present). Update each subsection:

### Subsections to maintain

- **Revenue Drivers** — what drove growth in each historical year? Segments, geographies, products, M&A.
- **Margin Drivers** — what drove gross margin trajectory? OpEx leverage or pressure? Mix shifts?
- **Capital Allocation** — capex cycles, M&A, buybacks, dividends. What pattern is management following?
- **Notable Events** — restructurings, accounting changes, leadership changes, regulatory shifts.

This section is **refined, not appended**. Each session improves it. Don't duplicate content — if you find new information, integrate it; if existing content is still accurate, leave it.

Cite sources inline (10-K Item 7 page X, 10-Q Q3'25, prior session research date). This makes the Historical Analysis auditable and lets future-you trace any claim back to its origin.

If significant new information has emerged since the last update (new filing, major event, segment reorganization), add a brief "Update [date]:" note at the section level.

## Step 6 — Synthesize driver proposals

Now propose forward driver values. For each driver, do this in order:

1. **State the historical anchor.** "3-year CAGR was X%. FY-1 was Y%."
2. **State the consensus.** "Broker consensus FY1 is Z%."
3. **State your proposal.** "I propose A%."
4. **Articulate the variant view.** "[Above/below/in line with] consensus by [N] bps because [reason from filings, Notion historical analysis, or qualitative judgment]."
5. **Specify the trajectory.** Decline path, expansion path, or hold-flat — and why.

If you can't articulate step 4 with specificity, your proposal should match consensus. "I just think it'll be higher" is not an acceptable variant view.

### Driver-by-driver guidance

**Revenue Growth %:**
- Anchor: 3Y CAGR + most recent YoY trend
- Consensus: broker FY1/FY2/FY3 implied growth
- Default shape: decline over 10 years toward terminal (2-3% mature, GDP+ healthy growers, can stay >10% for high-quality growth for 3-5 years)

**Gross Margin %:**
- Anchor: historical average + recent trend
- Consensus: broker FY1 implied gross margin
- Most companies have stable gross margins. Big swings need explicit justification (mix shift, input cost change, pricing power change).

**Total OpEx % of Revenue:**
- Anchor: 3-year average
- Consensus: typically not directly available; back-into from EBITDA margin and gross margin
- Operating leverage thesis (declining OpEx %) needs a specific story (scale, restructuring, automation).

**CapEx % of Revenue:**
- Anchor: 3-year average
- Consensus: broker capex estimates if available (not all sectors have good capex coverage)
- Sector capital intensity matters more than company-specific patterns for terminal years.

**D&A % of CapEx:**
- Anchor: historical ratio
- Consensus: usually not directly modeled
- Mature businesses converge to ~1.0x. Growth businesses have D&A < CapEx.

**Exit EBITDA Multiple:**
- Anchor: current trading multiple (compute from `_CapIQ_Data!E22 / _CapIQ_Data!D32` for current EV / FY-1 EBITDA, or use consensus FY1 multiple)
- Consensus: not really a "consensus multiple" but you can reference average price targets implied multiples
- Judgment-heavy. Default: gradual compression toward sector terminal (10-15x for industrials, 18-25x for high-quality compounders, 6-10x for cyclicals).

## Step 7 — Write YAML config

File: `companies/configs/<TICKER>.yaml`. Preserve any existing fields not being modified.

Schema:

```yaml
ticker: AAPL
company_name: Apple Inc.
sector: Information Technology
research_date: 2026-04-27
analyst: Aubrey
notion_page_url: https://notion.so/...

drivers:
  revenue_growth:
    values: [0.06, 0.05, 0.05, 0.04, 0.04, 0.035, 0.03, 0.03, 0.025, 0.025]
    historical_anchor: "3Y CAGR 1.0%, FY-1 YoY -0.8%"
    consensus:
      FY1: 0.050
      FY2: 0.052
      FY3: 0.053
      n_analysts: 28
    vs_consensus: |
      Above consensus FY1 by 100bps. Consensus appears to underweight services momentum;
      FY-1 services grew +14% and structural drivers (App Store, advertising, financial services)
      remain intact per 10-K Item 7. Hardware drag (-3% in FY-1) likely to moderate as iPhone 17
      cycle hits in 2026.
    rationale: |
      Historical anchor: 1.0% CAGR over last 3 years, but this masks divergent segment trends.
      Services CAGR ~14%, hardware roughly flat. Going forward, services becomes a larger share
      of mix (currently ~24%, growing), pulling consolidated growth toward 5-6% near-term.
      Decay to 2.5% terminal as installed base saturation continues and services growth normalizes.

      Sources:
      - 10-K FY2025 Item 7 (services segment narrative, p. 27-29)
      - Notion AAPL page → Historical Analysis → Revenue Drivers
      - Broker consensus from extract_broker_estimates output
    confidence: medium

  gross_margin:
    values: [0.46, 0.47, 0.47, 0.47, 0.47, 0.47, 0.47, 0.47, 0.47, 0.47]
    historical_anchor: "3Y avg 45.7%, FY-1 47.1% (expanding)"
    consensus:
      FY1: 0.464
    vs_consensus: |
      Roughly in line with consensus. Mix shift toward services (higher gross margin) supports
      modest expansion to 47% steady state. No specific catalyst for stronger expansion.
    rationale: |
      Historical 3Y avg: 45.7%. FY-1 at 47.1% reflects services mix shift continuing.
      47% steady state assumes mix continues but stabilizes. Hardware GM around 35-37%,
      services GM around 70%+. As services grows from ~24% to ~30% of mix, blended GM expands ~150bps.
    confidence: high

  opex_pct_rev:
    values: [0.13, 0.13, 0.13, 0.13, 0.13, 0.13, 0.13, 0.13, 0.13, 0.13]
    historical_anchor: "3Y avg 12.9%, stable"
    consensus:
      FY1_implied: 0.131  # back-solved from EBITDA margin and gross margin estimates
    vs_consensus: |
      In line with consensus. No operating leverage thesis. R&D growing ~8% offset by S&M leverage.
    rationale: |
      Stable across cycles. Apple maintains R&D investment through cycles and S&M is structurally
      light (brand-driven). 13% holds as the steady state.
    confidence: high

  capex_pct_rev:
    values: [0.025, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025]
    historical_anchor: "3Y avg 2.7%, FY-1 2.4%"
    consensus:
      FY1: 0.024
    vs_consensus: |
      Slightly above consensus. FY-1 capex was unusually low; expecting normalization.
      AI infrastructure investment may push capex modestly higher in next 2-3 years before
      normalizing. Holding 2.5% across the period as a reasonable mean.
    rationale: |
      Apple's capex is supplier-funded for most manufacturing, so reported capex is mostly
      data centers, retail, and corporate. ~2.5% of revenue is the long-run norm.
    confidence: high

  da_pct_capex:
    values: [1.10, 1.05, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00]
    historical_anchor: "3Y avg 1.10x, FY-1 1.21x"
    consensus: null
    vs_consensus: null
    rationale: |
      Above 1.0x suggests harvesting prior investments. Converge to 1.0x by year 3 as steady state.
    confidence: medium

  exit_multiple:
    values: [25.0, 24.0, 23.0, 22.0, 21.0, 20.0, 19.0, 19.0, 19.0, 19.0]
    historical_anchor: "Current EV/FY-1 EBITDA: ~19x. 5Y avg trading multiple ~22x."
    consensus:
      avg_price_target_implied_multiple_FY1: 23.5  # back-solved from price targets
    vs_consensus: |
      Trajectory below consensus implied multiple. Consensus seems to extrapolate current premium
      multiple. Compression to ~19x reflects view that quality premium narrows as growth normalizes
      and AAPL converges toward broader large-cap tech multiples.
    rationale: |
      Current trading: ~24x EV/LTM EBITDA. Premium to S&P given quality, brand, capital returns.
      Gradual compression as growth fades to mid-single-digits. 19x terminal reflects retained
      quality premium (vs. ~17x S&P average) but converging.
    confidence: low

single_drivers:
  dps_growth: 0.05
  cash_sweep_pct: 0.0
  min_cash_balance: 30000

# Implied vs. consensus summary (computed for quick reference, optional)
consensus_comparison:
  fy1_revenue_growth: {mine: 0.06, consensus: 0.05, delta_bps: 100}
  fy1_gross_margin: {mine: 0.46, consensus: 0.464, delta_bps: -40}
  fy1_ebitda: null  # implied from drivers; see rationale.md
```

If consensus is unavailable for a driver (e.g., D&A % of CapEx), set `consensus: null` and `vs_consensus: null`.

## Step 8 — Write rationale markdown

File: `companies/output/<TICKER>/drivers_rationale.md`

Long-form research note. Structure:

```markdown
# AAPL — Driver Research

**Date:** 2026-04-27 | **Analyst:** Aubrey

## Executive Summary

[3-4 sentences: company, core thesis, where you differ from consensus, key risks]

## Historical Context

[2-3 paragraphs: what the historicals show, key trends, what changed recently]

## Driver Discussion

For each driver:
### [Driver Name]
- **Historical:** [anchor]
- **Consensus:** [consensus values + n analysts]
- **Proposal:** [your numbers]
- **Variant view:** [where you differ and why]
- **Trajectory:** [shape and reasoning]
- **Sources cited:** [10-K sections, 10-Q quarters, Notion refs]
- **Confidence:** [high/medium/low]

## Cross-driver consistency check

[Verify that proposed drivers compose to a coherent story. E.g., if you have 6% revenue growth
with stable margins and stable capex %, your implied EBITDA growth is also ~6%. Is that the
implied growth you intended? If not, revisit drivers.]

## Open questions / monitoring items

[Things to watch in next earnings, regulatory developments, sector shifts that would invalidate
the thesis]

## Sources

[Bibliography: filings, Notion pages, broker estimates date, etc.]
```

This document is the audit trail. Write it as if a colleague (or future-you) will need to review the model in 6 months.

## Step 9 — Sync to Notion

If Notion MCP available:

### If company page exists:

1. Update **Historical Analysis section** if you made refinements in Step 5 (already done — verify).
2. Append a new section: `## Driver Research — 2026-04-27` (use today's date).
3. In that section, include:
   - Link to model file path: `companies/output/<TICKER>/<TICKER>_model.xlsx`
   - Link to rationale path: `companies/output/<TICKER>/drivers_rationale.md`
   - Summary table of proposed drivers with consensus comparison
   - 1-paragraph thesis recap

Driver Research entries are **append-only** — never modify or delete prior entries. They're snapshots of analytical state at a point in time.

### If company page does not exist:

1. Create a new page under the appropriate parent (search for "Investing" → "Companies" or sector-specific parent).
2. Page structure:

```
# [Company Name] ([TICKER])

## Snapshot
- Sector: [...]
- Last researched: [date]
- Model: [path]

## Thesis
[Empty — user fills in]

## Historical Analysis
### Revenue Drivers
[populated from Step 5]
### Margin Drivers
[...]
### Capital Allocation
[...]
### Notable Events
[...]

## Driver Research — 2026-04-27
[populated from this session]
```

3. Update YAML's `notion_page_url` field with the new page URL.

## Step 10 — Report back to user

Print a clean summary:

```
Driver research complete: AAPL

Historical anchor:
  3Y revenue CAGR: 1.0%
  3Y avg gross margin: 45.7%
  3Y avg EBITDA margin: 33.8%

Consensus (FY1):
  Revenue growth: 5.0% (28 analysts)
  Gross margin: 46.4%
  Avg price target: $195.50 (16.5% upside)

My proposals (FY1):
  Revenue growth: 6.0% (+100bps vs consensus)
  Gross margin: 46.0% (-40bps vs consensus)
  Exit multiple: 25.0x (terminal 19.0x)

Variant view: Above-consensus revenue growth driven by stronger services momentum than
broker models capture. Multiple compression view contrasts with consensus extrapolation
of current premium.

Files written:
  companies/configs/AAPL.yaml
  companies/output/AAPL/drivers_rationale.md
  Notion page: <url>

To populate the model with these assumptions:
  python -m companies.scripts.populate_drivers AAPL
```
