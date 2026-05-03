"""Tests for fundamentals_fetcher.fetch_for_ticker.

Mocks yfinance so we can probe the computed-ratio + clamp-negative logic
without hitting the network.
"""

from __future__ import annotations

from social.wsb_momentum import fundamentals_fetcher as ff


class _FakeTicker:
    def __init__(self, info):
        self.info = info


def _patch_yf(monkeypatch, info_dict):
    monkeypatch.setattr(ff.yf, "Ticker", lambda _t: _FakeTicker(info_dict))


def test_fetch_full_row_computes_all_ratios(monkeypatch):
    _patch_yf(monkeypatch, {
        "sharesOutstanding":       1_000_000_000,
        "sharesShort":               200_000_000,
        "floatShares":               800_000_000,
        "shortRatio":                3.5,
        "heldPercentInstitutions":   0.65,
        "heldPercentInsiders":       0.10,
    })
    f = ff.fetch_for_ticker("XYZ")
    assert f["shares_outstanding"] == 1_000_000_000
    assert f["shares_short"] == 200_000_000
    assert f["float_shares"] == 800_000_000
    assert f["short_ratio"] == 3.5
    assert f["short_pct_float"] == 200_000_000 / 800_000_000
    assert f["float_pct_outstanding"] == 800_000_000 / 1_000_000_000
    assert f["held_pct_institutions"] == 0.65
    assert f["held_pct_insiders"] == 0.10
    assert abs(f["held_pct_retail"] - (1 - 0.65 - 0.10)) < 1e-12


def test_fetch_clamps_negative_retail_to_null(monkeypatch, capsys):
    # Inst+insiders > 100% — yfinance occasionally produces this for foreign / dual listings.
    _patch_yf(monkeypatch, {
        "sharesOutstanding":         100_000_000,
        "heldPercentInstitutions":   0.95,
        "heldPercentInsiders":       0.20,
    })
    f = ff.fetch_for_ticker("FOREIGN")
    assert f["held_pct_retail"] is None
    out = capsys.readouterr().out
    assert "clamping to NULL" in out
    assert "FOREIGN" in out


def test_fetch_retail_null_when_inputs_missing(monkeypatch):
    _patch_yf(monkeypatch, {
        "sharesOutstanding": 100_000_000,
        # No institutions / insiders fields at all.
    })
    f = ff.fetch_for_ticker("ETF")
    assert f["held_pct_institutions"] is None
    assert f["held_pct_insiders"] is None
    assert f["held_pct_retail"] is None
    # Other derived ratios also NULL because their inputs were missing.
    assert f["short_pct_float"] is None
    assert f["float_pct_outstanding"] is None


def test_fetch_retail_zero_is_kept(monkeypatch):
    # A genuine 0% retail (institutions+insiders sum to exactly 1.0) — should
    # store 0.0, not NULL.
    _patch_yf(monkeypatch, {
        "heldPercentInstitutions": 0.80,
        "heldPercentInsiders":     0.20,
    })
    f = ff.fetch_for_ticker("LOCKEDUP")
    assert f["held_pct_retail"] == 0.0


def test_fetch_handles_yf_exception(monkeypatch):
    def _boom(_t):
        raise RuntimeError("yfinance flapped")
    monkeypatch.setattr(ff.yf, "Ticker", _boom)
    assert ff.fetch_for_ticker("ANY") is None
