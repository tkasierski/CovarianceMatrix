from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import BinaryIO, Iterable

import pandas as pd
import yfinance as yf


def parse_tickers(ticker_text: str | Iterable[str]) -> list[str]:
    if isinstance(ticker_text, str):
        raw = ticker_text.replace(",", "\n").replace(" ", "\n").splitlines()
    else:
        raw = list(ticker_text)
    seen: set[str] = set()
    output: list[str] = []
    for item in raw:
        ticker = str(item).strip().upper()
        if ticker and ticker not in seen:
            seen.add(ticker)
            output.append(ticker)
    return output


def _extract_adjusted_close(raw: pd.DataFrame, tickers: list[str]) -> tuple[pd.DataFrame, list[str], list[str]]:
    prices = pd.DataFrame()
    if len(tickers) == 1:
        ticker = tickers[0]
        if "Adj Close" in raw.columns:
            prices[ticker] = raw["Adj Close"]
    elif isinstance(raw.columns, pd.MultiIndex):
        for ticker in tickers:
            key = (ticker, "Adj Close")
            if key in raw.columns:
                prices[ticker] = raw[key]
    valid = prices.columns.tolist()
    failed = [ticker for ticker in tickers if ticker not in valid]
    return prices, valid, failed


def download_adjusted_close(tickers: Iterable[str], start_date: str, end_date: str) -> tuple[pd.DataFrame, list[str], list[str]]:
    parsed = parse_tickers(tickers)
    if not parsed:
        return pd.DataFrame(), [], []
    raw = yf.download(
        tickers=parsed,
        start=start_date,
        end=end_date,
        auto_adjust=False,
        progress=False,
        group_by="ticker",
        threads=True,
    )
    if raw.empty:
        raise ValueError("No data returned from Yahoo Finance.")
    prices, valid, failed = _extract_adjusted_close(raw, parsed)
    prices.index = pd.to_datetime(prices.index)
    return prices.dropna(how="all"), valid, failed


def monthly_simple_returns(prices: pd.DataFrame) -> pd.DataFrame:
    if prices.empty:
        return pd.DataFrame()
    monthly_prices = prices.resample("ME").last()
    return monthly_prices.pct_change(fill_method=None).iloc[1:]


def _read_tabular(source: str | Path | BinaryIO | bytes, filename: str | None = None) -> pd.DataFrame:
    if isinstance(source, bytes):
        source = BytesIO(source)
    suffix = Path(filename or str(source)).suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(source)
    return pd.read_csv(source)


def load_custom_returns(source: str | Path | BinaryIO | bytes, filename: str | None = None) -> pd.DataFrame:
    frame = _read_tabular(source, filename=filename)
    if frame.empty:
        return pd.DataFrame()
    date_candidates = [c for c in frame.columns if str(c).strip().lower() in {"date", "month", "period"}]
    date_col = date_candidates[0] if date_candidates else frame.columns[0]
    dates = pd.to_datetime(frame.pop(date_col), errors="coerce")
    frame.index = dates
    frame = frame.loc[frame.index.notna()].sort_index()
    frame.columns = [str(c).strip() for c in frame.columns]
    frame = frame.apply(pd.to_numeric, errors="coerce")
    frame.index = frame.index.to_period("M").to_timestamp("M")
    return frame.groupby(level=0).last().dropna(how="all")


def merge_return_sources(public_returns: pd.DataFrame | None, custom_returns: pd.DataFrame | None) -> pd.DataFrame:
    frames = [frame for frame in [public_returns, custom_returns] if frame is not None and not frame.empty]
    if not frames:
        raise ValueError("No public or custom return data were provided.")
    merged = pd.concat(frames, axis=1).sort_index()
    duplicates = merged.columns[merged.columns.duplicated()].tolist()
    if duplicates:
        raise ValueError(f"Duplicate asset names across return sources: {', '.join(duplicates)}")
    return merged.dropna(how="all")
