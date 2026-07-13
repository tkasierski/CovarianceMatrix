# Covariance Matrix

A Python and Streamlit tool for combining public-market prices with monthly simple returns from hedge funds, private assets, or other manually maintained investments.

## Methodology

All analytics use monthly **simple returns**. Public prices are resampled to month-end and converted with percentage change. Custom files must contain a date column followed by decimal return columns.

Missing data is explicit:

- `listwise` (default): uses only months with a return for every asset. This produces a consistent common-history covariance matrix.
- `pairwise`: uses all overlapping observations for each asset pair. This retains more data, but can produce a covariance matrix that is not positive semidefinite.

The workbook includes an observation-count matrix and a validation sheet. The default minimum history is 24 months.

## Excel design

The workbook contains:

- `Monthly_Simple_Returns`
- `Covariance_Matrix`
- `Correlation_Matrix`
- `Observation_Counts`
- `Downside_Risk_Metrics`
- `Drawdown_Series`
- `Portfolio_Dashboard`
- `Validation`

Covariance and correlation formulas use direct return-column ranges rather than repeated `INDEX/MATCH` lookups. The portfolio dashboard has a dedicated annual risk-free-rate input used in the Sharpe ratio.

## CLI

```bash
covariance-matrix \
  --tickers SPY GLD EWJ \
  --custom-returns hedge_funds.xlsx \
  --start 2018-01-01 \
  --end 2026-06-30 \
  --missing-data-method listwise \
  --min-observations 24 \
  --risk-free-rate 0.04
```

At least one ticker or custom-return file is required.

## Custom return file

```text
Date,Fund A,Fund B
2024-01-31,0.012,-0.004
2024-02-29,0.008,0.011
```

Returns must be decimal simple returns, so `0.012` means `1.2%`.

## Development

```bash
python -m pip install -e '.[dev]'
pytest -q
```

GitHub Actions runs the test suite on Python 3.10, 3.11, and 3.12 for every pull request and push to `main`.
