from covariance_matrix.core import build_covariance_excel, parse_tickers


def test_public_api_exports():
    assert callable(build_covariance_excel)
    assert parse_tickers("spy") == ["SPY"]
