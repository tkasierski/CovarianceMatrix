import inspect

import numpy as np
import pandas as pd

import covariance_matrix.simple_core as simple_core
from covariance_matrix.core import (
    calculate_downside_risk_metrics,
    calculate_drawdown_metrics,
    parse_tickers,
)


def test_parse_tickers_removes_duplicates_and_accepts_mixed_separators():
    assert parse_tickers("aapl, msft\nAAPL spy") == ["AAPL", "MSFT", "SPY"]


def test_drawdown_metrics_identifies_largest_drawdown():
    returns = pd.DataFrame(
        {
            "A": [0.10, -0.20, 0.05, 0.25],
        },
        index=pd.date_range("2024-01-31", periods=4, freq="ME"),
    )

    summary, series = calculate_drawdown_metrics(returns)

    assert np.isclose(summary.loc["A", "Largest Peak-to-Trough Drawdown"], -0.20)
    assert summary.loc["A", "Longest Drawdown (Months)"] == 2
    assert "A" in series.columns


def test_downside_risk_metrics_reports_positive_var_loss_numbers():
    returns = pd.DataFrame(
        {
            "A": [-0.10, -0.05, 0.00, 0.03, 0.07],
        },
        index=pd.date_range("2024-01-31", periods=5, freq="ME"),
    )

    metrics, _ = calculate_downside_risk_metrics(returns)

    assert metrics.loc["A", "Observations"] == 5
    assert metrics.loc["A", "Historical 95% VaR"] > 0
    assert metrics.loc["A", "Historical 95% CVaR"] > 0


def test_simple_return_pipeline_does_not_reference_log_returns():
    source = inspect.getsource(simple_core)

    assert "Monthly_Log_Returns" not in source
    assert "monthly_log_returns" not in source
    assert "np.log" not in source
    assert 'returns_sheet = "Monthly_Simple_Returns"' in source
