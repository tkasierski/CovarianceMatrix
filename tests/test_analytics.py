import numpy as np
import pandas as pd

from covariance_matrix.analytics import analyze_returns, prepare_returns


def sample_returns():
    return pd.DataFrame(
        {"A": [0.01, 0.02, np.nan, 0.04], "B": [0.02, np.nan, 0.03, 0.05]},
        index=pd.date_range("2024-01-31", periods=4, freq="ME"),
    )


def test_listwise_uses_common_history():
    prepared = prepare_returns(sample_returns(), "listwise")
    assert len(prepared) == 2
    assert not prepared.isna().any().any()


def test_pairwise_preserves_partial_history():
    prepared = prepare_returns(sample_returns(), "pairwise")
    assert len(prepared) == 4
    assert prepared.isna().any().any()


def test_observation_counts_are_pair_specific():
    result = analyze_returns(sample_returns(), missing_data_method="pairwise", min_observations=2)
    assert result.observation_counts.loc["A", "A"] == 3
    assert result.observation_counts.loc["A", "B"] == 2


def test_historical_var_and_cvar_use_realized_rolling_annual_returns():
    monthly = pd.Series(
        [0.01] * 12 + [-0.10],
        index=pd.date_range("2024-01-31", periods=13, freq="ME"),
    )
    returns = pd.DataFrame({"A": monthly})
    result = analyze_returns(returns, min_observations=2)

    rolling_annual = (1 + monthly).rolling(12, min_periods=12).apply(np.prod, raw=True) - 1
    rolling_annual = rolling_annual.dropna()
    q95 = rolling_annual.quantile(0.05)
    q99 = rolling_annual.quantile(0.01)
    metrics = result.downside_metrics.loc["A"]

    assert np.isclose(metrics["Historical Annual 95% VaR"], -q95)
    assert np.isclose(
        metrics["Historical Annual 95% CVaR"],
        -rolling_annual[rolling_annual <= q95].mean(),
    )
    assert np.isclose(metrics["Historical Annual 99% VaR"], -q99)
    assert np.isclose(
        metrics["Historical Annual 99% CVaR"],
        -rolling_annual[rolling_annual <= q99].mean(),
    )
