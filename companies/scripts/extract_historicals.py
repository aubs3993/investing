"""Extract historicals brief from a ticker's _CapIQ_Data tab.

Usage:
    python -m companies.scripts.extract_historicals <TICKER>
    python -m companies.scripts.extract_historicals <TICKER> --format markdown
    python -m companies.scripts.extract_historicals <TICKER> --output PATH
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]


def _model_path_for(ticker: str) -> Path:
    return REPO_ROOT / "companies" / "output" / ticker / f"{ticker}_model.xlsx"


def _safe_div(num, den):
    try:
        if den in (None, 0) or num is None:
            return None
        return num / den
    except (TypeError, ZeroDivisionError):
        return None


def _cagr(start, end, years):
    if start is None or end is None or start <= 0 or end <= 0 or years <= 0:
        return None
    return (end / start) ** (1 / years) - 1


def _yoy(prev, curr):
    if prev in (None, 0) or curr is None:
        return None
    return curr / prev - 1


def _avg(vals):
    clean = [v for v in vals if isinstance(v, (int, float))]
    if not clean:
        return None
    return sum(clean) / len(clean)


def _get(ws, row, col):
    return ws.cell(row, col).value


def extract(ticker: str) -> dict:
    model = _model_path_for(ticker)
    if not model.exists():
        raise SystemExit(
            f"Missing {model}. Bootstrap with `python -m companies.scripts.new_ticker {ticker}` "
            f"and then run `python -m shared.fetch_capiq {ticker}`."
        )
    wb = load_workbook(model, data_only=True)
    if "_CapIQ_Data" not in wb.sheetnames:
        raise SystemExit("Model is missing _CapIQ_Data tab. Regenerate via scaffold_template.")
    cap = wb["_CapIQ_Data"]

    # Historical rows on _CapIQ_Data: B/C/D = FY-3, FY-2, FY-1
    def hist(row):
        return [_get(cap, row, 2), _get(cap, row, 3), _get(cap, row, 4)]

    revenue = hist(25)
    cogs = hist(26)
    gp = hist(27)
    sga = hist(28)
    rd = hist(29)
    opex = hist(30)
    da = hist(31)
    ebitda = hist(32)
    ebit = hist(33)
    int_exp = hist(34)
    int_inc = hist(35)
    pretax = hist(36)
    taxes = hist(37)
    ni = hist(38)
    diluted_shares = hist(39)
    capex = hist(40)

    def block(values):
        return {
            "FY-3": values[0],
            "FY-2": values[1],
            "FY-1": values[2],
            "cagr_3y": _cagr(values[0], values[2], 2),  # FY-3 to FY-1 = 2 periods
            "yoy_latest": _yoy(values[1], values[2]),
        }

    historicals = {
        "revenue": block(revenue),
        "cogs": block(cogs),
        "gross_profit": block(gp),
        "sga": block(sga),
        "rd_expense": block(rd),
        "total_opex": block(opex),
        "d_and_a": block(da),
        "ebitda": block(ebitda),
        "ebit": block(ebit),
        "interest_expense": block(int_exp),
        "interest_income": block(int_inc),
        "pretax_income": block(pretax),
        "taxes": block(taxes),
        "net_income": block(ni),
        "diluted_weighted_avg_shares": block(diluted_shares),
        "capex": block(capex),
    }

    def ratio_block(num, den):
        vals = [_safe_div(n, d) for n, d in zip(num, den)]
        return {
            "FY-3": vals[0], "FY-2": vals[1], "FY-1": vals[2],
            "avg_3y": _avg(vals),
        }

    ratios = {
        "gross_margin":     ratio_block(gp, revenue),
        "ebitda_margin":    ratio_block(ebitda, revenue),
        "ebit_margin":      ratio_block(ebit, revenue),
        "opex_pct_rev":     ratio_block(opex, revenue),
        "capex_pct_rev":    ratio_block(capex, revenue),
        "da_pct_capex":     ratio_block(da, capex),
        "effective_tax":    ratio_block(taxes, pretax),
    }

    current_state = {
        "company_name":          _get(cap, 7, 5),
        "sector":                _get(cap, 8, 5),
        "currency":              _get(cap, 9, 5),
        "filing_status":         _get(cap, 10, 5),
        "price":                 _get(cap, 13, 5),
        "diluted_shares_mm":     _get(cap, 14, 5),
        "quarterly_dps":         _get(cap, 15, 5),
        "cash_mm":               _get(cap, 16, 5),
        "debt_mm":               _get(cap, 17, 5),
        "minority_interest":     _get(cap, 18, 5),
        "equity_investments":    _get(cap, 19, 5),
        "effective_tax_rate":    _get(cap, 20, 5),
        "market_cap_mm":         _get(cap, 21, 5),
        "net_debt_mm":           _get(cap, 22, 5),
    }
    price = current_state["price"]
    qdps = current_state["quarterly_dps"]
    if isinstance(qdps, (int, float)):
        current_state["annual_dps"] = qdps * 4
    if isinstance(price, (int, float)) and isinstance(current_state["diluted_shares_mm"], (int, float)):
        current_state.setdefault("market_cap_mm", price * current_state["diluted_shares_mm"])
    debt = current_state.get("debt_mm")
    cash = current_state.get("cash_mm")
    if isinstance(debt, (int, float)) and isinstance(cash, (int, float)):
        current_state["enterprise_value_mm"] = (
            (current_state.get("market_cap_mm") or 0) + debt - cash
            + (current_state.get("minority_interest") or 0)
            - (current_state.get("equity_investments") or 0)
        )

    return {
        "ticker": ticker,
        "company_name": current_state.get("company_name"),
        "sector": current_state.get("sector"),
        "currency": current_state.get("currency"),
        "fetch_timestamp": _get(cap, 2, 2),
        "historicals": historicals,
        "ratios": ratios,
        "current_state": current_state,
    }


def _fmt_pct(v): return f"{v*100:.1f}%" if isinstance(v, (int, float)) else "—"
def _fmt_num(v): return f"{v:,.0f}" if isinstance(v, (int, float)) else "—"
def _fmt_money(v): return f"${v:,.2f}" if isinstance(v, (int, float)) else "—"


def to_markdown(data: dict) -> str:
    h = data["historicals"]
    r = data["ratios"]
    cs = data["current_state"]
    lines = []
    lines.append(f"# {data['ticker']} — Historicals Brief")
    lines.append("")
    lines.append(f"**Company:** {data.get('company_name')}  ")
    lines.append(f"**Sector:** {data.get('sector')}  ")
    lines.append(f"**Currency:** {data.get('currency')}  ")
    lines.append(f"**Fetched:** {data.get('fetch_timestamp')}")
    lines.append("")
    lines.append("## Historicals")
    lines.append("")
    lines.append("| Metric | FY-3 | FY-2 | FY-1 | 3Y CAGR | YoY latest |")
    lines.append("|---|---:|---:|---:|---:|---:|")
    for label, key in [
        ("Revenue", "revenue"), ("Gross Profit", "gross_profit"),
        ("Total OpEx", "total_opex"), ("EBITDA", "ebitda"),
        ("EBIT", "ebit"), ("D&A", "d_and_a"),
        ("Net Income", "net_income"), ("CapEx", "capex"),
    ]:
        b = h[key]
        lines.append(
            f"| {label} | {_fmt_num(b['FY-3'])} | {_fmt_num(b['FY-2'])} | {_fmt_num(b['FY-1'])} | "
            f"{_fmt_pct(b['cagr_3y'])} | {_fmt_pct(b['yoy_latest'])} |"
        )
    lines.append("")
    lines.append("## Ratios")
    lines.append("")
    lines.append("| Ratio | FY-3 | FY-2 | FY-1 | 3Y avg |")
    lines.append("|---|---:|---:|---:|---:|")
    for label, key in [
        ("Gross Margin", "gross_margin"),
        ("EBITDA Margin", "ebitda_margin"),
        ("EBIT Margin", "ebit_margin"),
        ("OpEx % of Revenue", "opex_pct_rev"),
        ("CapEx % of Revenue", "capex_pct_rev"),
        ("D&A % of CapEx", "da_pct_capex"),
        ("Effective Tax Rate", "effective_tax"),
    ]:
        b = r[key]
        lines.append(
            f"| {label} | {_fmt_pct(b['FY-3'])} | {_fmt_pct(b['FY-2'])} | {_fmt_pct(b['FY-1'])} | "
            f"{_fmt_pct(b['avg_3y'])} |"
        )
    lines.append("")
    lines.append("## Current State")
    lines.append("")
    lines.append(f"- Price: {_fmt_money(cs.get('price'))}")
    lines.append(f"- Diluted shares (M): {_fmt_num(cs.get('diluted_shares_mm'))}")
    lines.append(f"- Market cap (M): {_fmt_num(cs.get('market_cap_mm'))}")
    lines.append(f"- Cash (M): {_fmt_num(cs.get('cash_mm'))}")
    lines.append(f"- Debt (M): {_fmt_num(cs.get('debt_mm'))}")
    lines.append(f"- Net debt (M): {_fmt_num(cs.get('net_debt_mm'))}")
    lines.append(f"- Enterprise value (M): {_fmt_num(cs.get('enterprise_value_mm'))}")
    lines.append(f"- Quarterly DPS: {_fmt_money(cs.get('quarterly_dps'))}")
    lines.append(f"- Annual DPS: {_fmt_money(cs.get('annual_dps'))}")
    return "\n".join(lines)


def main(argv=None):
    parser = argparse.ArgumentParser(description="Extract historicals brief from _CapIQ_Data.")
    parser.add_argument("ticker")
    parser.add_argument("--format", choices=["json", "markdown"], default="json")
    parser.add_argument("--output", default=None,
                        help="Write to file instead of stdout.")
    args = parser.parse_args(argv)
    ticker = args.ticker.strip().upper()
    data = extract(ticker)
    if args.format == "json":
        out = json.dumps(data, indent=2, default=str)
    else:
        out = to_markdown(data)
    if args.output:
        Path(args.output).write_text(out, encoding="utf-8")
        print(f"Wrote {args.output}")
    else:
        sys.stdout.write(out + "\n")


if __name__ == "__main__":
    main()
