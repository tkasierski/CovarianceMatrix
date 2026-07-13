from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np
import pandas as pd

MissingDataMethod = Literal["listwise", "pairwise"]


@dataclass(frozen=True)
class AnalysisResult:
    returns: pd.DataFrame
    covariance: pd.DataFrame
    correlation: pd.DataFrame
    observation_counts: pd.DataFrame
    downside_metrics: pd.DataFrame
    drawdown_series: pd.DataFrame
    warnings: list[str]
    missing_data_method: MissingDataMethod
    min_observations: int


def validate_returns(returns: pd.DataFrame, min_observations: int = 24) -> list[str]:
    warnings: list[str] = []
    if returns.index.duplicated().any():
        warnings.append("Duplicate return dates were detected.")
    for asset in returns.columns:
        series = returns[asset].dropna()
        if (series < -1).any():
            warnings.append(f"{asset}: return below -100% detected.")
        if len(series) < min_observations:
            warnings.append(f"{asset}: only {len(series)} observations; minimum is {min_observations}.")
        if len(series) and np.isclose(series.std(ddof=1), 0):
            warnings.append(f"{asset}: zero or near-zero volatility.")
    return warnings


def prepare_returns(returns: pd.DataFrame, method: MissingDataMethod) -> pd.DataFrame:
    if method not in {"listwise", "pairwise"}:
        raise ValueError("missing_data_method must be 'listwise' or 'pairwise'.")
    cleaned = returns.sort_index().replace([np.inf, -np.inf], np.nan)
    return cleaned.dropna(how="any") if method == "listwise" else cleaned.dropna(how="all")


def observation_counts(returns: pd.DataFrame) -> pd.DataFrame:
    columns = returns.columns
    counts = pd.DataFrame(index=columns, columns=columns, dtype=int)
    for left in columns:
        for right in columns:
            counts.loc[left, right] = int(returns[[left, right]].dropna().shape[0])
    return counts


def calculate_drawdown_metrics(returns: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    summary: list[dict[str, object]] = []
    series_out = pd.DataFrame(index=returns.index)
    for asset in returns.columns:
        series = returns[asset].dropna()
        if series.empty:
            continue
        wealth = (1 + series).cumprod()
        peak = wealth.cummax()
        drawdown = wealth / peak - 1
        series_out[asset] = drawdown
        duration = 0
        longest = 0
        for value in drawdown:
            duration = duration + 1 if value < 0 else 0
            longest = max(longest, duration)
        summary.append({
            "Asset": asset,
            "Largest Peak-to-Trough Drawdown": drawdown.min(),
            "Longest Drawdown (Months)": longest,
            "Current Drawdown": drawdown.iloc[-1],
        })
    return pd.DataFrame(summary).set_index("Asset"), series_out


def calculate_downside_metrics(returns: pd.DataFrame, minimum_acceptable_return: float = 0.0) -> tuple[pd.DataFrame, pd.DataFrame]:
    rows: list[dict[str, object]] = []
    drawdown_summary, drawdown_series = calculate_drawdown_metrics(returns)
    for asset in returns.columns:
        series = returns[asset].dropna()
        if series.empty:
            continue
        downside = np.minimum(series - minimum_acceptable_return, 0)
        q95 = series.quantile(0.05)
        q99 = series.quantile(0.01)
        rows.append({
            "Asset": asset,
            "Observations": len(series),
            "Monthly Volatility": series.std(),
            "Annualized Volatility": series.std() * np.sqrt(12),
            "Monthly Downside Deviation": np.sqrt(np.mean(downside**2)),
            "Annualized Downside Deviation": np.sqrt(np.mean(downside**2)) * np.sqrt(12),
            "Historical 95% VaR": -q95,
            "Historical 95% CVaR": -series[series <= q95].mean(),
            "Historical 99% VaR": -q99,
            "Historical 99% CVaR": -series[series <= q99].mean(),
        })
    metrics = pd.DataFrame(rows).set_index("Asset")
    return metrics.join(drawdown_summary, how="left"), drawdown_series


def analyze_returns(
    returns: pd.DataFrame,
    missing_data_method: MissingDataMethod = "listwise",
    min_observations: int = 24,
    minimum_acceptable_return: float = 0.0,
) -> AnalysisResult:
    prepared = prepare_returns(returns, missing_data_method)
    if prepared.empty:
        raise ValueError("No usable monthly returns remain after missing-data treatment.")
    warnings = validate_returns(prepared, min_observations)
    covariance = prepared.cov(min_periods=min_observations)
    correlation = prepared.corr(min_periods=min_observations)
    counts = observation_counts(prepared)
    downside, drawdown = calculate_downside_metrics(prepared, minimum_acceptable_return)
    finite_cov = covariance.fillna(0).to_numpy(dtype=float)
    if finite_cov.size and np.linalg.eigvalsh((finite_cov + finite_cov.T) / 2).min() < -1e-10:
        warnings.append("Covariance matrix is not positive semidefinite; portfolio risk may be unstable.")
    return AnalysisResult(prepared, covariance, correlation, counts, downside, drawdown, warnings, missing_data_method, min_observations)
