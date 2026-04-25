# companies

Company-specific filings analysis.

## Layout
- `configs/[ticker].yaml` — per-company input file (ticker, CIK, assumptions)
- `templates/` — blank Excel models with formulas and named ranges. **Canonical source — never modified directly.** New versions go in by replacing the file with a new commit.
- `output/[ticker]/` — filled Excel files per company; gitignored

## Workflow
Python populates templates via `openpyxl` writing to named ranges, then saves the filled file to `output/[ticker]/`. Excel itself evaluates the formulas when the file is opened. Never reference raw cells in Python — always use named ranges.
