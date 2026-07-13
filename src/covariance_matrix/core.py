from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import BinaryIO, Iterable

import pandas as pd

from .analytics import MissingDataMethod, analyze_returns
from .data import download_adjusted_close, load_custom_returns, merge_return_sources, monthly_simple_returns, parse_tickers
from .excel import build_workbook


def build_covariance_excel(
    tickers: Iterable[str] = (),
    start_date: str = "2018-01-01",
    end_date: str = "2025-12-31",
    output_prefix: str = "covariance_matrix",
    minimum_acceptable_return: float = 0.0,
    risk_free_rate: float = 0.0,
    missing_data_method: MissingDataMethod = "listwise",
    min_observations: int = 24,
    custom_returns_source: str | Path | BinaryIO | bytes | None = None,
    custom_returns_filename: str | None = None,
    output_dir: str | Path = ".",
) -> dict[str, object]:
    parsed = parse_tickers(tickers)
    prices = pd.DataFrame()
    valid: list[str] = []
    failed: list[str] = []
    public_returns = pd.DataFrame()
    if parsed:
        prices, valid, failed = download_adjusted_close(parsed, start_date, end_date)
        public_returns = monthly_simple_returns(prices)
    custom_returns = pd.DataFrame()
    if custom_returns_source is not None:
        custom_returns = load_custom_returns(custom_returns_source, custom_returns_filename)
    merged = merge_return_sources(public_returns, custom_returns)
    result = analyze_returns(
        merged,
        missing_data_method=missing_data_method,
        min_observations=min_observations,
        minimum_acceptable_return=minimum_acceptable_return,
    )
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / f"{output_prefix}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    build_workbook(result, output_file, risk_free_rate, minimum_acceptable_return, prices)
    return {
        "output_file": str(output_file),
        "valid_tickers": valid,
        "failed_tickers": failed,
        "assets": list(result.returns.columns),
        "warnings": result.warnings,
        "missing_data_method": result.missing_data_method,
    }


__all__ = ["build_covariance_excel", "parse_tickers"]
