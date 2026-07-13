from io import BytesIO

import pandas as pd

from covariance_matrix.data import load_custom_returns, merge_return_sources, monthly_simple_returns, parse_tickers


def test_parse_tickers_deduplicates():
    assert parse_tickers("spy, GLD\nSPY") == ["SPY", "GLD"]


def test_monthly_simple_returns():
    prices = pd.DataFrame({"A": [100, 110, 121]}, index=pd.to_datetime(["2024-01-31", "2024-02-29", "2024-03-31"]))
    returns = monthly_simple_returns(prices)
    assert returns["A"].round(6).tolist() == [0.1, 0.1]


def test_load_custom_returns_csv():
    payload = BytesIO(b"Date,Fund A\n2024-01-31,0.01\n2024-02-29,-0.02\n")
    result = load_custom_returns(payload, "returns.csv")
    assert list(result.columns) == ["Fund A"]
    assert result.iloc[1, 0] == -0.02


def test_merge_rejects_duplicate_asset_names():
    index = pd.date_range("2024-01-31", periods=2, freq="ME")
    left = pd.DataFrame({"A": [0.1, 0.2]}, index=index)
    right = pd.DataFrame({"A": [0.3, 0.4]}, index=index)
    try:
        merge_return_sources(left, right)
    except ValueError as exc:
        assert "Duplicate asset names" in str(exc)
    else:
        raise AssertionError("Expected duplicate asset names to fail")
