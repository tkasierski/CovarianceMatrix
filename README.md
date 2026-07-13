# Covariance Matrix

A Python command-line and Streamlit tool for generating covariance, correlation, downside-risk, drawdown, and portfolio risk-dashboard tables for public securities using Yahoo Finance data.

This repository was migrated from a Google Colab notebook into a portable Python package.

## Return convention

The tool uses **monthly simple returns for every risk calculation**. Public-market returns therefore use the same convention as typical hedge-fund and private-investment return reports.

Using one return convention means that:

- covariance and correlation use monthly simple returns
- volatility and portfolio risk use monthly simple returns
- VaR, CVaR, downside deviation, and drawdowns use monthly simple returns
- custom hedge-fund or private-asset returns only need to be pasted into one worksheet

The generated workbook no longer includes or depends on a monthly log-return worksheet.

## Features

- Pulls daily adjusted close prices from Yahoo Finance via `yfinance`
- Resamples daily prices to month-end prices
- Computes monthly simple returns
- Exports a formula-driven Excel workbook with:
  - monthly and annualized covariance matrices
  - correlation matrix
  - downside-risk metrics
  - drawdown series
  - portfolio allocation and risk-contribution dashboards
  - daily adjusted closes
  - monthly simple returns

## Installation

Clone the repository and install it locally:

```bash
git clone https://github.com/tkasierski/CovarianceMatrix.git
cd CovarianceMatrix
python -m pip install -e .
```

## Usage

Run from the command line:

```bash
covariance-matrix \
  --tickers AAPL MSFT GOOGL AMZN META SPY \
  --start 2018-01-01 \
  --end 2025-12-31 \
  --output-prefix covariance_matrix
```

Optional minimum acceptable monthly return for downside deviation:

```bash
covariance-matrix \
  --tickers AAPL MSFT SPY \
  --start 2020-01-01 \
  --end 2024-12-31 \
  --minimum-acceptable-return 0.005
```

The command writes a timestamped `.xlsx` file to the selected output directory.

## Custom return streams

Open the generated workbook and replace values in `Monthly_Simple_Returns` with monthly simple returns for public securities, hedge funds, or private assets. Covariance, correlation, downside-risk, drawdown, and portfolio-risk formulas will recalculate from that single worksheet.

Use net returns when the objective is to measure investor-level net risk.

## Development

Install development dependencies:

```bash
python -m pip install -e '.[dev]'
```

Run tests:

```bash
pytest
```

## Notes

Yahoo Finance data can have missing values, ticker-specific gaps, and survivorship limitations. The tool requires at least one complete monthly observation across all selected assets for the covariance and correlation matrices.
