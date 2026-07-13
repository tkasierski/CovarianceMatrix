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
