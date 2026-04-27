# Templates

Master templates used by scripts across the repo.

## Workflow

1. Master templates live here as `.xlsx` files.
2. Scripts **read** from these files but never write back to them.
3. Per-ticker scripts copy a template to `companies/output/<TICKER>/<TICKER>_model.xlsx`,
   then write inputs to **named ranges** (never cell coordinates) and save.
4. After initial scaffolding, master templates are hand-edited in Excel.
   The scaffolding scripts in `shared/` are one-shot and not re-run after manual edits.

## Files

- `company_model.xlsx` — base financial model for general corporates (industrials, consumer, tech, healthcare).
  Generated initially by `shared/scaffold_template.py`. Edit in Excel after first build.

## Important: iterative calculation

This model contains an interest ↔ taxes ↔ debt balance circularity (deliberate, not a bug).
Before opening the workbook, in Excel: **File → Options → Formulas → Enable iterative calculation**
(Maximum iterations: 100, Maximum change: 0.001).

## CapIQ data flow

The model uses S&P Capital IQ for historical data and current state inputs (price, shares, cash, debt, etc.). The data flow:

1. `capiq_fetcher.xlsx` — standalone workbook with live CapIQ formulas tied to a single ticker input. Open in Excel with CapIQ plugin loaded to refresh manually, or driven programmatically by the fetch script.
2. `company_model.xlsx` → hidden tab `_CapIQ_Data` — mirror of the fetcher's data layout, but holds **hardcoded values**. Populated by `shared/fetch_capiq.py`.
3. `company_model.xlsx` → IS, CF, Inputs tabs — link to `_CapIQ_Data` via formulas (green font convention). Never reference CapIQ directly, so the main template stays usable on machines without CapIQ access.

### To refresh data for a ticker

```
python -m shared.fetch_capiq <TICKER>
```

This will (1) open the fetcher, (2) set the ticker, (3) wait for CapIQ formulas to resolve, (4) write values into `_CapIQ_Data` in `company_model.xlsx`, (5) save and close.

### To add a new CapIQ field

1. Open `capiq_fetcher.xlsx`. Add a new row with a label in column A and CapIQ formula(s) in the appropriate column(s).
2. Open `company_model.xlsx`. Unhide `_CapIQ_Data`. Add a row at the **same position** with the same label in column A. Leave value cells empty (the fetch script will populate).
3. Optional: add a formula link from IS, CF, or Inputs to the new cell in `_CapIQ_Data`.
4. Run `python -m shared.fetch_capiq <TICKER>` to confirm the new field flows through.

The fetcher's `Fetcher` tab and the model's `_CapIQ_Data` tab **must have identical row/column structures**. The fetch script validates this and errors if they drift. The shared layout source of truth lives in `shared/capiq_layout.py` — editing it and rerunning both scaffolders is the cleanest way to add fields if you don't mind regenerating both files.

## Broker estimates

`templates/broker_fetcher.xlsx` is a separate workbook pulling broker consensus forecasts from CapIQ (FY1/FY2/FY3 mean, FY1 high/low/count, analyst sentiment). The data flow:

1. `broker_fetcher.xlsx` — live CapIQ broker estimate formulas
2. `_Broker_Data` hidden tab in `company_model.xlsx` — hardcoded values, populated by `shared/fetch_broker_estimates.py`
3. The driver research playbook compares proposed driver values against broker consensus and forces articulation of variant view

To refresh broker estimates: `python -m shared.fetch_broker_estimates <TICKER>`

Broker estimates change frequently (post-earnings, mid-quarter revisions). Refresh independently of historicals.

The shared layout source of truth lives in `shared/broker_layout.py` — same convention as the CapIQ layer.

## Driver assumptions

The Inputs tab driver rows (revenue growth, gross margin, OpEx %, CapEx %, D&A % of CapEx, exit multiple) have named ranges (`drv_*`) populated by `companies/scripts/populate_drivers.py` reading from `companies/configs/<TICKER>.yaml`.
