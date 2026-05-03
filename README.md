# Investing Research

Personal investing research repo. Code is organized by what it analyzes.

## Layout

- `macro/` — standalone scripts that pull macro data series (yield curve, CPI, employment, etc.). Each script writes outputs to `macro/output/<topic>/`. See `macro/README.md` for the convention.
- `companies/` — company-specific filings analysis. Configs in `companies/configs/`, generated artifacts in `companies/output/`.
- `templates/` — master Excel templates copied and filled by per-ticker scripts. See `templates/README.md`.
- `social/` — sentiment and attention data from social platforms. Currently contains `wsb_momentum/`, which snapshots r/wallstreetbets posts every 30 minutes and pairs them with intraday price data and short-interest fundamentals to study correlation between retail attention and stock price movement. See `social/wsb_momentum/README.md`.
- `shared/` — reusable helpers used across `macro/`, `companies/`, and `social/`.

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
