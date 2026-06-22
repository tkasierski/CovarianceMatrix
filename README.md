# Covariance Matrix

A Python command-line tool for generating covariance, correlation, downside-risk, and drawdown tables for public securities using Yahoo Finance data.

This repository was migrated from a Google Colab notebook into a portable Python package.

## Features

- Pulls daily adjusted close prices from Yahoo Finance via `yfinance`
- Resamples daily prices to month-end prices
- Computes monthly log returns for covariance and correlation matrices
- Computes monthly simple returns for downside-risk metrics
- Exports a formatted Excel workbook with:
  - covariance matrix
  - correlation matrix
  - downside-risk metrics
  - drawdown series
  - daily adjusted closes
  - monthly log returns
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

Yahoo Finance data can have missing values, ticker-specific gaps, and survivorship limitations. This tool preserves the original notebook methodology by dropping incomplete monthly log-return rows for the covariance and correlation calculations.