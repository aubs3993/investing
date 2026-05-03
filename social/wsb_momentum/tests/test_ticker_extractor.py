"""Tests for ticker_extractor.

Run from repo root:
    pytest social/wsb_momentum/tests/test_ticker_extractor.py
"""

from __future__ import annotations

from social.wsb_momentum.ticker_extractor import extract_tickers, load_ticker_set


def test_load_ticker_set_has_known_tickers():
    tickers = load_ticker_set()
    # Spot-check: TSLA, AAPL, GME should always be in NASDAQ; A and AA in NYSE.
    for sym in ("TSLA", "AAPL", "GME", "AA"):
        assert sym in tickers, f"{sym} missing from reference set"
    assert len(tickers) > 5000  # we loaded ~11k


def test_cashtag_match():
    assert extract_tickers("Loaded up on $TSLA today") == {"TSLA"}


def test_bare_ticker_match():
    assert extract_tickers("GME to the moon") == {"GME"}


def test_multiple_tickers():
    out = extract_tickers("Buying TSLA, $AAPL, and GME")
    assert out == {"TSLA", "AAPL", "GME"}


def test_blacklist_rejects_bare_jargon():
    # "DD" exists as a real NYSE ticker (DuPont) but the blacklist rejects it bare.
    # "YOLO" is also in the ticker list as an ETF and must be rejected bare.
    assert extract_tickers("Did my own DD on this YOLO play") == set()


def test_cashtag_bypasses_blacklist():
    # Explicit cashtag means the author meant the ticker — bypass the blacklist.
    assert extract_tickers("Bought $DD calls") == {"DD"}


def test_ambiguous_single_letter_A():
    # 'A' is Agilent on NYSE and is in our reference set, but it's overwhelmingly
    # the English article on WSB. Policy: blacklisted bare, but cashtag bypasses.
    assert extract_tickers("A great trade today") == set()
    assert extract_tickers("$A is undervalued") == {"A"}


def test_empty_string():
    assert extract_tickers("") == set()
    assert extract_tickers(None) == set()  # type: ignore[arg-type]


def test_mixed_case_rejected():
    # Lower or mixed case is not extracted — would trigger massive false positives.
    assert extract_tickers("tsla and Tsla and TsLa") == set()


def test_punctuation_and_word_boundaries():
    # Tickers attached to common punctuation should still match.
    assert extract_tickers("AAPL.") == {"AAPL"}
    assert extract_tickers("(NVDA)") == {"NVDA"}
    assert extract_tickers("MSFT, GOOG!") == {"MSFT", "GOOG"}


def test_unknown_uppercase_word_dropped():
    # Random all-caps word that isn't a real ticker should be dropped silently.
    assert extract_tickers("ZZZZX is a fake ticker") == set()


def test_too_long_token_ignored():
    # 6+ char all-caps tokens are not even regex candidates.
    assert extract_tickers("ABCDEF is too long") == set()
