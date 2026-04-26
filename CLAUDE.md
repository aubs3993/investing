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
