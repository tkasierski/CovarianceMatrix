from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
import yfinance as yf


def parse_tickers(ticker_text: str | Iterable[str]) -> list[str]:
    """Parse tickers from commas, spaces, new lines, or an iterable of strings."""
    if isinstance(ticker_text, str):
        raw_items = ticker_text.replace(",", "\n").replace(" ", "\n").split("\n")
    else:
        raw_items = list(ticker_text)

    tickers = [str(item).strip().upper() for item in raw_items if str(item).strip()]

    seen: set[str] = set()
    unique_tickers: list[str] = []
    for ticker in tickers:
        if ticker not in seen:
            seen.add(ticker)
            unique_tickers.append(ticker)

    return unique_tickers


def calculate_drawdown_metrics(
    monthly_simple_returns: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Calculate drawdown summary metrics and drawdown time series by asset."""
    drawdown_summary: list[dict[str, object]] = []
    drawdown_series = pd.DataFrame(index=monthly_simple_returns.index)

    for asset in monthly_simple_returns.columns:
        returns = monthly_simple_returns[asset].dropna()

        if returns.empty:
            drawdown_summary.append(
                {
                    "Asset": asset,
                    "Largest Peak-to-Trough Drawdown": np.nan,
                    "Longest Drawdown (Months)": np.nan,
                    "Current Drawdown": np.nan,
                    "Drawdown Start": pd.NaT,
                    "Drawdown Trough": pd.NaT,
                    "Drawdown Recovery": pd.NaT,
                }
            )
            continue

        wealth_index = (1 + returns).cumprod()
        running_peak = wealth_index.cummax()
        drawdown = wealth_index / running_peak - 1
        drawdown_series[asset] = drawdown

        max_drawdown = drawdown.min()
        trough_date = drawdown.idxmin()

        pre_trough = wealth_index.loc[:trough_date]
        peak_value_before_trough = pre_trough.cummax().loc[trough_date]
        peak_dates = pre_trough[pre_trough == peak_value_before_trough].index
        drawdown_start = peak_dates[-1] if len(peak_dates) > 0 else pd.NaT

        post_trough = wealth_index.loc[trough_date:]
        recovery_candidates = post_trough[post_trough >= peak_value_before_trough]
        drawdown_recovery = recovery_candidates.index[0] if len(recovery_candidates) > 0 else pd.NaT

        in_drawdown = drawdown < 0
        longest_duration = 0
        current_duration = 0
        for is_drawdown in in_drawdown:
            if is_drawdown:
                current_duration += 1
                longest_duration = max(longest_duration, current_duration)
            else:
                current_duration = 0

        drawdown_summary.append(
            {
                "Asset": asset,
                "Largest Peak-to-Trough Drawdown": max_drawdown,
                "Longest Drawdown (Months)": longest_duration,
                "Current Drawdown": drawdown.iloc[-1],
                "Drawdown Start": drawdown_start,
                "Drawdown Trough": trough_date,
                "Drawdown Recovery": drawdown_recovery,
            }
        )

    return pd.DataFrame(drawdown_summary).set_index("Asset"), drawdown_series


def calculate_downside_risk_metrics(
    monthly_simple_returns: pd.DataFrame,
    minimum_acceptable_return: float = 0.0,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Calculate volatility, downside deviation, VaR, CVaR, and drawdown metrics."""
    metrics: list[dict[str, object]] = []
    drawdown_summary, drawdown_series = calculate_drawdown_metrics(monthly_simple_returns)

    for asset in monthly_simple_returns.columns:
        returns = monthly_simple_returns[asset].dropna()

        if returns.empty:
            metrics.append(
                {
                    "Asset": asset,
                    "Observations": 0,
                    "Monthly Volatility": np.nan,
                    "Annualized Volatility": np.nan,
                    "Monthly Downside Deviation": np.nan,
                    "Annualized Downside Deviation": np.nan,
                    "Historical 95% VaR": np.nan,
                    "Historical 95% CVaR": np.nan,
                    "Historical 99% VaR": np.nan,
                    "Historical 99% CVaR": np.nan,
                }
            )
            continue

        monthly_vol = returns.std()
        annualized_vol = monthly_vol * np.sqrt(12)

        downside_returns = np.minimum(returns - minimum_acceptable_return, 0)
        monthly_downside_deviation = np.sqrt(np.mean(downside_returns**2))
        annualized_downside_deviation = monthly_downside_deviation * np.sqrt(12)

        var_95_threshold = returns.quantile(0.05)
        cvar_95_returns = returns[returns <= var_95_threshold]
        var_99_threshold = returns.quantile(0.01)
        cvar_99_returns = returns[returns <= var_99_threshold]

        metrics.append(
            {
                "Asset": asset,
                "Observations": len(returns),
                "Monthly Volatility": monthly_vol,
                "Annualized Volatility": annualized_vol,
                "Monthly Downside Deviation": monthly_downside_deviation,
                "Annualized Downside Deviation": annualized_downside_deviation,
                "Historical 95% VaR": -var_95_threshold,
                "Historical 95% CVaR": -cvar_95_returns.mean() if len(cvar_95_returns) > 0 else np.nan,
                "Historical 99% VaR": -var_99_threshold,
                "Historical 99% CVaR": -cvar_99_returns.mean() if len(cvar_99_returns) > 0 else np.nan,
            }
        )

    return pd.DataFrame(metrics).set_index("Asset").join(drawdown_summary, how="left"), drawdown_series


def _extract_adjusted_close(raw: pd.DataFrame, tickers: list[str]) -> tuple[pd.DataFrame, list[str], list[str]]:
    adj_close = pd.DataFrame()

    if len(tickers) == 1:
        ticker = tickers[0]
        if "Adj Close" in raw.columns:
            adj_close[ticker] = raw["Adj Close"]
    else:
        for ticker in tickers:
            if isinstance(raw.columns, pd.MultiIndex) and (ticker, "Adj Close") in raw.columns:
                adj_close[ticker] = raw[(ticker, "Adj Close")]

    valid_tickers = adj_close.columns.tolist()
    failed_tickers = [ticker for ticker in tickers if ticker not in valid_tickers]
    return adj_close, valid_tickers, failed_tickers


def build_covariance_excel(
    tickers: Iterable[str],
    start_date: str,
    end_date: str,
    output_prefix: str = "covariance_matrix",
    minimum_acceptable_return: float = 0.0,
    output_dir: str | Path = ".",
) -> dict[str, object]:
    """Build covariance analysis tables and write a formatted Excel workbook."""
    parsed_tickers = parse_tickers(tickers)
    if not parsed_tickers:
        raise ValueError("Ticker list is empty.")

    raw = yf.download(
        tickers=parsed_tickers,
        start=start_date,
        end=end_date,
        auto_adjust=False,
        progress=False,
        group_by="ticker",
        threads=True,
    )

    if raw.empty:
        raise ValueError("No data returned from Yahoo Finance.")

    adj_close, valid_tickers, failed_tickers = _extract_adjusted_close(raw, parsed_tickers)
    if adj_close.empty:
        raise ValueError("No valid adjusted close data found for any ticker.")

    adj_close = adj_close.dropna(how="all")
    adj_close.index = pd.to_datetime(adj_close.index)

    monthly_prices = adj_close.resample("ME").last()
    monthly_log_returns = np.log(monthly_prices / monthly_prices.shift(1)).iloc[1:]
    monthly_simple_returns = monthly_prices.pct_change().iloc[1:]
    monthly_log_returns_clean = monthly_log_returns.dropna(how="any")

    if monthly_log_returns_clean.empty:
        raise ValueError("No complete monthly return rows remain after dropping missing data.")

    covariance_matrix = monthly_log_returns_clean.cov()
    correlation_matrix = monthly_log_returns_clean.corr()
    downside_metrics, drawdown_series = calculate_downside_risk_metrics(
        monthly_simple_returns=monthly_simple_returns,
        minimum_acceptable_return=minimum_acceptable_return,
    )

    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = output_path / f"{output_prefix}_{timestamp}.xlsx"

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        covariance_matrix.to_excel(writer, sheet_name="Covariance_Matrix")
        correlation_matrix.to_excel(writer, sheet_name="Correlation_Matrix")
        downside_metrics.to_excel(writer, sheet_name="Downside_Risk_Metrics")
        drawdown_series.to_excel(writer, sheet_name="Drawdown_Series")
        adj_close.to_excel(writer, sheet_name="Adj_Close_Daily")
        monthly_log_returns.to_excel(writer, sheet_name="Monthly_Log_Returns")
        monthly_simple_returns.to_excel(writer, sheet_name="Monthly_Simple_Returns")

        workbook = writer.book
        percent_format = workbook.add_format({"num_format": "0.00%"})
        number_format = workbook.add_format({"num_format": "0.00"})
        integer_format = workbook.add_format({"num_format": "0"})
        date_format = workbook.add_format({"num_format": "mm/dd/yyyy"})

        ws = writer.sheets["Downside_Risk_Metrics"]
        ws.set_column("A:A", 18)
        ws.set_column("B:B", 14, integer_format)
        ws.set_column("C:J", 18, percent_format)
        ws.set_column("K:M", 18, percent_format)
        ws.set_column("N:N", 18, integer_format)
        ws.set_column("O:Q", 18, date_format)

        for sheet_name in [
            "Covariance_Matrix",
            "Correlation_Matrix",
            "Drawdown_Series",
            "Monthly_Log_Returns",
            "Monthly_Simple_Returns",
        ]:
            ws = writer.sheets[sheet_name]
            ws.set_column("A:A", 14, date_format)
            ws.set_column("B:ZZ", 14, percent_format)

        ws = writer.sheets["Adj_Close_Daily"]
        ws.set_column("A:A", 14, date_format)
        ws.set_column("B:ZZ", 14, number_format)

    return {
        "adj_close_daily": adj_close,
        "monthly_log_returns": monthly_log_returns,
        "monthly_simple_returns": monthly_simple_returns,
        "monthly_log_returns_clean": monthly_log_returns_clean,
        "covariance_matrix": covariance_matrix,
        "correlation_matrix": correlation_matrix,
        "downside_metrics": downside_metrics,
        "drawdown_series": drawdown_series,
        "valid_tickers": valid_tickers,
        "failed_tickers": failed_tickers,
        "output_file": str(output_file),
    }
