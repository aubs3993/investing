# Investing Research

Personal investing research repo. Code is organized by what it analyzes.

## Layout

- `macro/` — standalone scripts that pull macro data series (yield curve, CPI, employment, etc.). Each script writes outputs to `macro/output/<topic>/`. See `macro/README.md` for the convention.
- `companies/` — company-specific filings analysis. Configs in `companies/configs/`, generated artifacts in `companies/output/`.
- `templates/` — master Excel templates copied and filled by per-ticker scripts. See `templates/README.md`.
- `shared/` — reusable helpers used across `macro/` and `companies/`.

## Setup

1. Copy `.env` and add your FRED API key:
   ```
   FRED_API_KEY=your_key_here
   ```
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

## Dependencies

- `pandas` — data manipulation
- `openpyxl` — Excel export
- `python-dotenv` — environment variable loading
- `fredapi` — FRED API client
- `matplotlib` — charting
