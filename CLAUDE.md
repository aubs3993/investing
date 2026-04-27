# CLAUDE.md

Conventions for this repo. Future Claude Code sessions should read this at startup and follow these patterns.

## Project purpose
Personal investing research repo. `macro/` holds individual scripts that pull macro data series from FRED. `companies/` will hold company-specific filings analysis. `shared/` holds reusable helpers used across both.

## Folder structure conventions
- `macro/[topic]_pull.py` — one Python script per macro data series; underscores in filenames
- `macro/output/[topic]/` — each script writes outputs to its own subfolder, never to `macro/output/` root
- `companies/configs/[ticker].yaml` — per-company input file (ticker, CIK, assumptions)
- `companies/templates/` — blank Excel templates with formulas; tracked in git; never modified directly
- `companies/output/[ticker]/` — filled Excel files per company; gitignored
- `shared/` — reusable helpers (`fred_helpers.py`, `excel_helpers.py`); refactor here as soon as code is duplicated across two scripts

## Naming
- Python files: underscores, not hyphens (`yield_curve_pull.py`, not `yield-curve-pull.py`)
- Output files: descriptive, no dates in filename (put timestamps in the data, not the name)

## Path handling
All scripts must use `Path(__file__).resolve().parent` to anchor file paths to the script's location, so scripts work regardless of working directory.

## Macro data conventions
- All macro/ scripts default to a start date of 2006-01-01. This window includes the GFC and COVID — the two stress events most relevant for current investing decisions — and keeps x-axes comparable across charts.
- Use end = datetime.today() for the end date.
- If a FRED series doesn't go back to 2006, start at the earliest available date and document the exception in a comment at the top of the script.
- For derived series (e.g. YoY % change) that need lookback data, fetch with a buffer (e.g. 2005-01-01) but filter the displayed/charted/saved data to start at 2006-01-01 so all macro charts have aligned x-axes.
- When computing reference statistics (median, percentile rank, min/max) for chart annotations, compute over the 2006+ window (what's on the chart), not the full series history.

## Secrets and environment
- `.env` at project root holds all API keys; always loaded via `python-dotenv`
- `.env` is gitignored; never commit, never paste in chat, never hardcode
- API keys belong only in `.env` or a password manager

## .gitignore conventions
- `macro/output/**/*` and `companies/output/**/*` are ignored
- `!macro/output/**/` and `!companies/output/**/` re-include directories so git can descend into them
- `!**/.gitkeep` tracks the `.gitkeep` markers
- `!macro/output/**/*.png` tracks chart PNGs under `macro/output/` so they can be referenced from Notion summaries and reviewed in git history; xlsx data files under `macro/output/` remain ignored
- `companies/output/` remains fully ignored (no PNG exception) — filled Excel files there are per-run artifacts, not shared assets
- When creating any new output subfolder, always add a `.gitkeep` file inside it

## Excel workflow (companies side)
Python populates templates but does NOT evaluate formulas. Workflow:
1. Build template in Excel with formulas and named ranges (e.g. `WACC`, `TerminalGrowth`, `Revenue_Y1`)
2. Python opens template, writes inputs to named ranges via `openpyxl`, saves to `companies/output/[ticker]/`
3. Open the filled file in real Excel to see calculated values

Always use named ranges in code (`ws["WACC"] = 0.085`), never raw cell references (`ws["B5"]`). This lets templates be visually reorganized without breaking the Python.

## Working conventions
- Always show file changes before applying them
- Always run `git status` and confirm `.env` is not staged before committing
- Never set git config without explicit instruction
- Always add a `.gitkeep` when creating an empty subfolder under `output/`
- Pause and ask if anything looks unexpected, especially around staging files that look like secrets or large binaries
- At the start of every local session, run `git status` and `git fetch --prune` to assess branch state. Then:
  - If on main with no uncommitted changes: run `git pull`
  - If on a feature branch with uncommitted or unpushed work: report the state to me and ask whether I want to continue this work or switch to main for something new
  - If on a stale feature branch (merged and deleted on origin): offer to switch to main, pull, and delete the local branch
  Do not auto-switch branches without my explicit instruction.

## Notion integration
The Notion MCP is connected for local sessions. When asked to update Notion:
- Investing pages live under the "Investing" parent page (id: 347440ca-0d18-8116-b9fa-cbb37af18b54)
- Macro analysis goes in the "Macro" sub-page (id: 34e440ca-0d18-8103-abb9-fe57c62ad904)
- Make routine additions and updates freely (new sections, appended content, refreshed summaries, new sub-pages under Investing). Briefly tell me what was changed after each batch of edits so I can scan for issues.
- Never delete existing Notion content without explicit approval
- Only use replace_content (full-page replacement) with explicit instruction. For in-place edits, use update_content with targeted search-and-replace.
- Reference chart files via their GitHub raw URLs, not local paths

## Project Conventions

### Folder structure
- `templates/` — read-only master `.xlsx` files. Scripts copy and fill, never modify in place.
- `companies/output/<TICKER>/` — per-ticker generated artifacts. Disposable, regenerable from template + config.
- `companies/configs/<ticker>.yaml` — per-ticker inputs that drive template fills.
- `shared/` — reusable Python helpers imported by `macro/` and `companies/` scripts.

### Library preferences
- Default to `openpyxl` for `.xlsx` reads/writes.
- Use `xlwings` only when the workflow requires live Excel calc, formula evaluation, or refreshing
  data connections (Power Query, external links).

### Excel template workflow
- Scripts that fill templates must:
  1. Copy the master template from `templates/` to the destination path.
  2. Open the copy and write to **named ranges** (e.g., `inp_ticker`), never cell coordinates.
  3. Save the copy. Never modify the master.
- Master template `.xlsx` files **are committed** to the repo. Binary diffs aren't useful — review
  template changes by opening the file in Excel before merging.
- Input cells in templates use the `inp_` prefix for named ranges and yellow fill with a dotted blue border.

### Modeling conventions
- Year columns combine historical (suffix `A`) and projected (suffix `E`).
- Banker color coding: blue font for hardcoded inputs, black for formulas, green for cross-tab links.
- Subtotal/total rows are bold with a single top border.

### CapIQ integration
- Historical financials and current-state inputs flow from `templates/capiq_fetcher.xlsx` (live CapIQ formulas) into `templates/company_model.xlsx` → hidden `_CapIQ_Data` tab (hardcoded values).
- IS, CF, and Inputs tabs link to `_CapIQ_Data` via formulas. They never reference CapIQ functions directly.
- The fetcher and `_CapIQ_Data` tab share an identical row/column layout — adding a field requires editing both files. `shared/capiq_layout.py` is the shared source of truth for row positions and field names.
- Refresh via `python -m shared.fetch_capiq <TICKER>` (uses xlwings to drive live Excel + CapIQ plugin).
- The main template must be openable on machines without CapIQ access — never break this invariant.

### Broker estimates integration
- Broker consensus forecasts flow from `templates/broker_fetcher.xlsx` (live `IQ_EST_*` formulas) into `templates/company_model.xlsx` → hidden `_Broker_Data` tab.
- `shared/broker_layout.py` is the shared layout source of truth (mirror of `capiq_layout.py`).
- Implied growth/margin rows (B20:B25) and implied upside (B30) live as formulas in the main template, not in the fetcher — they reference `_CapIQ_Data` historicals so they update when those refresh.
- Refresh via `python -m shared.fetch_broker_estimates <TICKER>`.

### Per-ticker workflow
1. **Bootstrap:** `python -m companies.scripts.new_ticker <TICKER>` (copies master template to `companies/output/<TICKER>/<TICKER>_model.xlsx`, creates skeleton YAML)
2. **Fetch historicals:** `python -m shared.fetch_capiq <TICKER>`
3. **Fetch broker estimates:** `python -m shared.fetch_broker_estimates <TICKER>`
4. **Extract briefs:** `python -m companies.scripts.extract_historicals <TICKER>` and `python -m companies.scripts.extract_broker_estimates <TICKER>`
5. **Research drivers:** Open Claude Code session. Ask Claude to research drivers for the ticker using `companies/scripts/driver_research_playbook.md`. Claude reads historicals, consensus, Notion, and SEC filings; writes to YAML config and rationale markdown; updates Historical Analysis in Notion (durable) and appends Driver Research entry (dated).
6. **Review:** User reviews YAML and rationale, makes edits.
7. **Populate model:** `python -m companies.scripts.populate_drivers <TICKER>`
8. **Final review:** Open model in Excel.

The split between research (Claude-driven) and populate (deterministic) means values can be regenerated without re-running research. Both fetch scripts auto-detect the per-ticker model copy if it exists, so the master template stays clean.

### Notion structure per company
Each company page in Notion has two distinct sections:
- **Historical Analysis** (durable, refined session-over-session) — captures *why* historicals look the way they do. Updated, not appended.
- **Driver Research — [date]** (one entry per session, append-only) — dated forward views. Snapshot of analytical state.

This separation prevents historical knowledge from getting buried under successive forward-looking sessions.

### Broker estimates as framing, not anchor
Broker estimates from `_Broker_Data` are a data point, not a target. The driver research playbook requires explicit articulation of variant view (above/below/in-line vs. consensus) for every driver. If variant view can't be articulated with specificity, default to consensus.
