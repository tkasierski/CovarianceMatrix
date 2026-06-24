from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell
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

        longest_duration = 0
        current_duration = 0
        for is_drawdown in drawdown < 0:
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
        downside_returns = np.minimum(returns - minimum_acceptable_return, 0)
        monthly_downside_deviation = np.sqrt(np.mean(downside_returns**2))
        var_95_threshold = returns.quantile(0.05)
        cvar_95_returns = returns[returns <= var_95_threshold]
        var_99_threshold = returns.quantile(0.01)
        cvar_99_returns = returns[returns <= var_99_threshold]

        metrics.append(
            {
                "Asset": asset,
                "Observations": len(returns),
                "Monthly Volatility": monthly_vol,
                "Annualized Volatility": monthly_vol * np.sqrt(12),
                "Monthly Downside Deviation": monthly_downside_deviation,
                "Annualized Downside Deviation": monthly_downside_deviation * np.sqrt(12),
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


def _quote_sheet(sheet_name: str) -> str:
    return f"'{sheet_name.replace(chr(39), chr(39) * 2)}'"


def _cell(row: int, col: int, *, absolute: bool = False) -> str:
    return xl_rowcol_to_cell(row, col, row_abs=absolute, col_abs=absolute)


def _range(sheet: str, first_row: int, first_col: int, last_row: int, last_col: int, *, absolute: bool = True) -> str:
    return (
        f"{_quote_sheet(sheet)}!"
        f"{xl_rowcol_to_cell(first_row, first_col, row_abs=absolute, col_abs=absolute)}:"
        f"{xl_rowcol_to_cell(last_row, last_col, row_abs=absolute, col_abs=absolute)}"
    )


def _write_formula(
    worksheet,
    row: int,
    col: int,
    formula: str,
    cell_format=None,
    cached_value: str | float = "",
) -> None:
    """Write a formula with a non-zero cached value only when explicitly requested.

    xlsxwriter's default cached formula result is zero. That can look like a hardcoded
    value in downloaded files before Excel recalculates. These workbook formulas are
    meant to be live on open, so dashboard formula cells are written with a blank cached
    value unless a better cached value is known.
    """
    worksheet.write_formula(row, col, formula, cell_format, cached_value)


def _write_formula_matrix(
    worksheet,
    start_row: int,
    start_col: int,
    formulas: list[list[str]],
    cell_format=None,
) -> None:
    for row_offset, row in enumerate(formulas):
        for col_offset, formula in enumerate(row):
            _write_formula(worksheet, start_row + row_offset, start_col + col_offset, formula, cell_format)


def _configure_formats(workbook) -> dict[str, object]:
    return {
        "title": workbook.add_format({"bold": True, "font_size": 14}),
        "note": workbook.add_format({"italic": True, "text_wrap": True, "font_color": "#666666"}),
        "header": workbook.add_format({"bold": True, "bg_color": "#E7E6E6", "border": 1, "text_wrap": True}),
        "asset": workbook.add_format({"bold": True, "border": 1}),
        "number": workbook.add_format({"num_format": "0.0000", "border": 1}),
        "percent": workbook.add_format({"num_format": "0.00%", "border": 1}),
        "currency": workbook.add_format({"num_format": "$#,##0", "border": 1}),
        "integer": workbook.add_format({"num_format": "0", "border": 1}),
        "date": workbook.add_format({"num_format": "mm/dd/yyyy", "border": 1}),
        "input_percent": workbook.add_format({"num_format": "0.00%", "border": 1, "bg_color": "#E2F0D9"}),
        "input_currency": workbook.add_format({"num_format": "$#,##0", "border": 1, "bg_color": "#E2F0D9"}),
    }


def _weighted_covariance_formula(
    annual_left_col: int,
    monthly_top_row: int,
    asset_index: int,
    n_assets: int,
    first_asset_row: int,
    left_col: int,
) -> str:
    """Return scalar weighted covariance terms for one asset row."""
    annual_row = monthly_top_row + 1 + asset_index
    terms: list[str] = []
    for other_asset_index in range(n_assets):
        annual_cov_cell = _cell(annual_row, annual_left_col + 1 + other_asset_index, absolute=True)
        weight_cell = _cell(first_asset_row + other_asset_index, left_col + 2, absolute=True)
        terms.append(f"{annual_cov_cell}*{weight_cell}")
    return "+".join(terms) if terms else "0"


def _write_live_covariance_sheet(
    writer: pd.ExcelWriter,
    valid_tickers: list[str],
    returns_last_row: int,
    formats: dict[str, object],
) -> dict[str, object]:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Covariance_Matrix")
    writer.sheets["Covariance_Matrix"] = worksheet
    worksheet.activate()

    n_assets = len(valid_tickers)
    log_sheet = "Monthly_Log_Returns"
    log_headers = _range(log_sheet, 0, 1, 0, n_assets)
    log_data = _range(log_sheet, 1, 1, returns_last_row, n_assets)

    monthly_top_row = 1
    monthly_left_col = 0
    annual_left_col = n_assets + 3
    dashboard_top_row = n_assets + 7
    helper_base_col = max(annual_left_col + n_assets + 2, 20)

    worksheet.write(0, 0, "Monthly Covariance", formats["title"])
    worksheet.write(0, annual_left_col, "Annual Covariance", formats["title"])
    note = (
        "Risk calculations use the return series shown in Monthly_Log_Returns and "
        "Monthly_Simple_Returns. To use custom return streams, replace values in both "
        "return tabs. Covariance, correlation, and dashboard risk use log returns; "
        "downside-risk metrics and drawdowns use simple returns. For private assets, "
        "hedge funds, and PMEs, paste net returns if net-risk analysis is desired. "
        "Expected returns are manual inputs."
    )
    worksheet.merge_range(dashboard_top_row - 2, 0, dashboard_top_row - 1, 7, note, formats["note"])

    for offset, ticker in enumerate(valid_tickers):
        worksheet.write(monthly_top_row, monthly_left_col + 1 + offset, ticker, formats["header"])
        worksheet.write(monthly_top_row + 1 + offset, monthly_left_col, ticker, formats["asset"])
        worksheet.write(monthly_top_row, annual_left_col + 1 + offset, ticker, formats["header"])
        worksheet.write(monthly_top_row + 1 + offset, annual_left_col, ticker, formats["asset"])

    monthly_formulas: list[list[str]] = []
    annual_formulas: list[list[str]] = []
    for row_offset in range(n_assets):
        row_asset_cell = _cell(monthly_top_row + 1 + row_offset, monthly_left_col, absolute=True)
        monthly_row: list[str] = []
        annual_row: list[str] = []
        for col_offset in range(n_assets):
            col_asset_cell = _cell(monthly_top_row, monthly_left_col + 1 + col_offset, absolute=True)
            formula = (
                f'=IFERROR(_xlfn.COVARIANCE.S('
                f'INDEX({log_data},0,MATCH({row_asset_cell},{log_headers},0)),'
                f'INDEX({log_data},0,MATCH({col_asset_cell},{log_headers},0))),"")'
            )
            monthly_row.append(formula)
            monthly_cell = xl_rowcol_to_cell(monthly_top_row + 1 + row_offset, monthly_left_col + 1 + col_offset)
            annual_row.append(f'=IFERROR({monthly_cell}*12,"")')
        monthly_formulas.append(monthly_row)
        annual_formulas.append(annual_row)

    _write_formula_matrix(worksheet, monthly_top_row + 1, monthly_left_col + 1, monthly_formulas, formats["percent"])
    _write_formula_matrix(worksheet, monthly_top_row + 1, annual_left_col + 1, annual_formulas, formats["percent"])

    annual_matrix = f"${xl_col_to_name(annual_left_col + 1)}${monthly_top_row + 2}:${xl_col_to_name(annual_left_col + n_assets)}${monthly_top_row + 1 + n_assets}"

    def write_dashboard(left_col: int, helper_col: int, title: str) -> dict[str, str]:
        worksheet.write(dashboard_top_row, left_col, title, formats["title"])
        headers = ["Asset", "$ Value", "Weight", "Expected Return", "Asset Risk", "Marginal Risk", "Risk Contribution", "% of Risk"]
        worksheet.write_row(dashboard_top_row + 1, left_col, headers, formats["header"])
        worksheet.write(dashboard_top_row + 1, helper_col, "Weighted Covariance", formats["header"])

        first_asset_row = dashboard_top_row + 2
        last_asset_row = first_asset_row + n_assets - 1
        total_row = last_asset_row + 1
        return_row = total_row + 2
        risk_row = return_row + 1
        sharpe_row = risk_row + 1
        weights_range = f"${xl_col_to_name(left_col + 2)}${first_asset_row + 1}:${xl_col_to_name(left_col + 2)}${last_asset_row + 1}"
        expected_range = f"${xl_col_to_name(left_col + 3)}${first_asset_row + 1}:${xl_col_to_name(left_col + 3)}${last_asset_row + 1}"
        helper_range = f"${xl_col_to_name(helper_col)}${first_asset_row + 1}:${xl_col_to_name(helper_col)}${last_asset_row + 1}"

        for asset_index in range(n_assets):
            row = first_asset_row + asset_index
            asset_ref = _cell(monthly_top_row + 1 + asset_index, monthly_left_col, absolute=True)
            weighted_cov_formula = _weighted_covariance_formula(
                annual_left_col=annual_left_col,
                monthly_top_row=monthly_top_row,
                asset_index=asset_index,
                n_assets=n_assets,
                first_asset_row=first_asset_row,
                left_col=left_col,
            )
            _write_formula(worksheet, row, left_col, f"={asset_ref}", formats["asset"])
            worksheet.write_number(row, left_col + 1, 0, formats["input_currency"])
            _write_formula(
                worksheet,
                row,
                left_col + 2,
                f'=IF({_cell(total_row, left_col + 1, absolute=True)}=0,"",{_cell(row, left_col + 1)}/{_cell(total_row, left_col + 1, absolute=True)})',
                formats["percent"],
            )
            worksheet.write_blank(row, left_col + 3, None, formats["input_percent"])
            _write_formula(worksheet, row, left_col + 4, f'=IFERROR(SQRT(INDEX({annual_matrix},{asset_index + 1},{asset_index + 1})),"")', formats["percent"])
            _write_formula(worksheet, row, helper_col, f'=IFERROR({weighted_cov_formula},"")', formats["percent"])
            _write_formula(worksheet, row, left_col + 5, f'=IF({_cell(risk_row, left_col + 1, absolute=True)}="","",{_cell(row, helper_col)}/{_cell(risk_row, left_col + 1, absolute=True)})', formats["percent"])
            _write_formula(worksheet, row, left_col + 6, f'=IF(OR({_cell(row, left_col + 2)}="",{_cell(row, left_col + 5)}=""),"",{_cell(row, left_col + 2)}*{_cell(row, left_col + 5)})', formats["percent"])
            _write_formula(worksheet, row, left_col + 7, f'=IF({_cell(risk_row, left_col + 1, absolute=True)}="","",{_cell(row, left_col + 6)}/{_cell(risk_row, left_col + 1, absolute=True)})', formats["percent"])

        worksheet.write(total_row, left_col, "Total", formats["asset"])
        _write_formula(worksheet, total_row, left_col + 1, f"=SUM({_cell(first_asset_row, left_col + 1)}:{_cell(last_asset_row, left_col + 1)})", formats["currency"], 0)
        _write_formula(worksheet, total_row, left_col + 2, f"=SUM({_cell(first_asset_row, left_col + 2)}:{_cell(last_asset_row, left_col + 2)})", formats["percent"])
        worksheet.write(return_row, left_col, "Return", formats["asset"])
        _write_formula(worksheet, return_row, left_col + 1, f"=SUMPRODUCT({weights_range},{expected_range})", formats["percent"])
        worksheet.write(risk_row, left_col, "Risk", formats["asset"])
        _write_formula(worksheet, risk_row, left_col + 1, f'=IF(SUM({weights_range})=0,"",SQRT(SUMPRODUCT({weights_range},{helper_range})))', formats["percent"])
        worksheet.write(sharpe_row, left_col, "Sharpe", formats["asset"])
        _write_formula(worksheet, sharpe_row, left_col + 1, f'=IF({_cell(risk_row, left_col + 1)}="","",({_cell(return_row, left_col + 1)}-{_cell(first_asset_row, left_col + 3)})/{_cell(risk_row, left_col + 1)})', formats["number"])

        return {
            "weights": weights_range,
            "first_asset_row": str(first_asset_row),
            "last_asset_row": str(last_asset_row),
            "return_cell": _cell(return_row, left_col + 1),
            "risk_cell": _cell(risk_row, left_col + 1),
        }

    potential = write_dashboard(0, helper_base_col, "Potential Portfolio")
    current = write_dashboard(10, helper_base_col + 1, "Current Portfolio")

    worksheet.set_column(0, max(annual_left_col + n_assets, 18), 14)
    worksheet.set_column(0, 0, 18)
    worksheet.set_column(helper_base_col, helper_base_col + 1, 14, None, {"hidden": True})
    worksheet.freeze_panes(monthly_top_row + 1, 1)
    worksheet.autofilter(dashboard_top_row + 1, 0, dashboard_top_row + 1 + n_assets, 7)

    return {
        "monthly_matrix": f"${xl_col_to_name(monthly_left_col + 1)}${monthly_top_row + 2}:${xl_col_to_name(monthly_left_col + n_assets)}${monthly_top_row + 1 + n_assets}",
        "annual_matrix": annual_matrix,
        "potential": potential,
        "current": current,
    }


def _write_live_correlation_sheet(
    writer: pd.ExcelWriter,
    valid_tickers: list[str],
    returns_last_row: int,
    formats: dict[str, object],
) -> None:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Correlation_Matrix")
    writer.sheets["Correlation_Matrix"] = worksheet

    n_assets = len(valid_tickers)
    log_sheet = "Monthly_Log_Returns"
    log_headers = _range(log_sheet, 0, 1, 0, n_assets)
    log_data = _range(log_sheet, 1, 1, returns_last_row, n_assets)

    worksheet.write(0, 0, "Correlation Matrix", formats["title"])
    for offset, ticker in enumerate(valid_tickers):
        worksheet.write(1, 1 + offset, ticker, formats["header"])
        worksheet.write(2 + offset, 0, ticker, formats["asset"])

    formulas: list[list[str]] = []
    for row_offset in range(n_assets):
        row_asset_cell = _cell(2 + row_offset, 0, absolute=True)
        row: list[str] = []
        for col_offset in range(n_assets):
            col_asset_cell = _cell(1, 1 + col_offset, absolute=True)
            row.append(
                f'=IFERROR(CORREL('
                f'INDEX({log_data},0,MATCH({row_asset_cell},{log_headers},0)),'
                f'INDEX({log_data},0,MATCH({col_asset_cell},{log_headers},0))),"")'
            )
        formulas.append(row)
    _write_formula_matrix(worksheet, 2, 1, formulas, formats["number"])
    worksheet.set_column(0, n_assets, 14)
    worksheet.freeze_panes(2, 1)


def _write_live_drawdown_sheet(
    writer: pd.ExcelWriter,
    monthly_simple_returns: pd.DataFrame,
    valid_tickers: list[str],
    formats: dict[str, object],
) -> dict[str, dict[str, str]]:
    workbook = writer.book
    drawdown_ws = workbook.add_worksheet("Drawdown_Series")
    calc_ws = workbook.add_worksheet("_Risk_Calc")
    writer.sheets["Drawdown_Series"] = drawdown_ws
    writer.sheets["_Risk_Calc"] = calc_ws
    calc_ws.hide()

    n_assets = len(valid_tickers)
    n_obs = len(monthly_simple_returns.index)
    simple_sheet = "Monthly_Simple_Returns"
    simple_headers = _range(simple_sheet, 0, 1, 0, n_assets)

    drawdown_ws.write(0, 0, "Date", formats["header"])
    calc_ws.write(0, 0, "Date", formats["header"])
    for col_offset, ticker in enumerate(valid_tickers):
        drawdown_ws.write(0, 1 + col_offset, ticker, formats["header"])
        base_col = 1 + col_offset * 4
        calc_ws.write(0, base_col, f"{ticker} Return", formats["header"])
        calc_ws.write(0, base_col + 1, f"{ticker} Wealth", formats["header"])
        calc_ws.write(0, base_col + 2, f"{ticker} Peak", formats["header"])
        calc_ws.write(0, base_col + 3, f"{ticker} DD Duration", formats["header"])

    for row_offset in range(n_obs):
        row = 1 + row_offset
        source_date = f"={_quote_sheet(simple_sheet)}!{_cell(row, 0)}"
        _write_formula(drawdown_ws, row, 0, source_date, formats["date"])
        _write_formula(calc_ws, row, 0, source_date, formats["date"])
        for col_offset, ticker in enumerate(valid_tickers):
            base_col = 1 + col_offset * 4
            ticker_literal = ticker.replace('"', '""')
            return_formula = (
                f"=IFERROR(INDEX({_quote_sheet(simple_sheet)}!$B${row + 1}:${xl_col_to_name(n_assets)}${row + 1},"
                f"1,MATCH(\"{ticker_literal}\",{simple_headers},0)),\"\")"
            )
            _write_formula(calc_ws, row, base_col, return_formula, formats["percent"])
            if row == 1:
                _write_formula(calc_ws, row, base_col + 1, f"=IFERROR(1+{_cell(row, base_col)},\"\")", formats["number"])
                _write_formula(calc_ws, row, base_col + 2, f"=IF({_cell(row, base_col + 1)}=\"\",\"\",MAX(1,{_cell(row, base_col + 1)}))", formats["number"])
            else:
                _write_formula(calc_ws, row, base_col + 1, f"=IF({_cell(row, base_col)}=\"\",\"\",{_cell(row - 1, base_col + 1)}*(1+{_cell(row, base_col)}))", formats["number"])
                _write_formula(calc_ws, row, base_col + 2, f"=IF({_cell(row, base_col + 1)}=\"\",\"\",MAX({_cell(row - 1, base_col + 2)},{_cell(row, base_col + 1)}))", formats["number"])
            _write_formula(calc_ws, row, base_col + 3, f"=IF({_cell(row, base_col + 1)}=\"\",\"\",IF({_cell(row, base_col + 1)}<{_cell(row, base_col + 2)},IF(ROW()=2,1,{_cell(row - 1, base_col + 3)}+1),0))", formats["integer"])
            _write_formula(drawdown_ws, row, 1 + col_offset, f"=IFERROR({_quote_sheet('_Risk_Calc')}!{_cell(row, base_col + 1)}/{_quote_sheet('_Risk_Calc')}!{_cell(row, base_col + 2)}-1,\"\")", formats["percent"])

    drawdown_ws.set_column(0, 0, 14, formats["date"])
    drawdown_ws.set_column(1, n_assets, 14, formats["percent"])
    calc_ws.set_column(0, 0, 14, formats["date"])
    calc_ws.set_column(1, 1 + n_assets * 4, 14)
    drawdown_ws.freeze_panes(1, 1)

    refs: dict[str, dict[str, str]] = {}
    for col_offset, ticker in enumerate(valid_tickers):
        drawdown_col = xl_col_to_name(1 + col_offset)
        duration_col = xl_col_to_name(1 + col_offset * 4 + 3)
        refs[ticker] = {
            "drawdown_range": f"{_quote_sheet('Drawdown_Series')}!${drawdown_col}$2:${drawdown_col}${n_obs + 1}",
            "duration_range": f"{_quote_sheet('_Risk_Calc')}!${duration_col}$2:${duration_col}${n_obs + 1}",
        }
    return refs


def _write_live_downside_sheet(
    writer: pd.ExcelWriter,
    valid_tickers: list[str],
    n_obs: int,
    minimum_acceptable_return: float,
    drawdown_refs: dict[str, dict[str, str]],
    formats: dict[str, object],
) -> None:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Downside_Risk_Metrics")
    writer.sheets["Downside_Risk_Metrics"] = worksheet

    n_assets = len(valid_tickers)
    simple_sheet = "Monthly_Simple_Returns"
    simple_headers = _range(simple_sheet, 0, 1, 0, n_assets)
    simple_data = _range(simple_sheet, 1, 1, n_obs, n_assets)

    worksheet.write(0, 0, "Minimum Acceptable Monthly Return", formats["header"])
    worksheet.write_number(0, 1, minimum_acceptable_return, formats["input_percent"])
    headers = [
        "Asset",
        "Observations",
        "Monthly Volatility",
        "Annualized Volatility",
        "Monthly Downside Deviation",
        "Annualized Downside Deviation",
        "Historical 95% VaR",
        "Historical 95% CVaR",
        "Historical 99% VaR",
        "Historical 99% CVaR",
        "Largest Peak-to-Trough Drawdown",
        "Longest Drawdown (Months)",
        "Current Drawdown",
    ]
    worksheet.write_row(2, 0, headers, formats["header"])

    for row_offset, ticker in enumerate(valid_tickers):
        row = 3 + row_offset
        worksheet.write(row, 0, ticker, formats["asset"])
        asset_cell = _cell(row, 0, absolute=True)
        asset_returns = f"INDEX({simple_data},0,MATCH({asset_cell},{simple_headers},0))"
        var_95_cell = _cell(row, 6)
        var_99_cell = _cell(row, 8)
        _write_formula(worksheet, row, 1, f"=COUNT({asset_returns})", formats["integer"])
        _write_formula(worksheet, row, 2, f"=IFERROR(_xlfn.STDEV.S({asset_returns}),\"\")", formats["percent"])
        _write_formula(worksheet, row, 3, f"=IFERROR({_cell(row, 2)}*SQRT(12),\"\")", formats["percent"])
        _write_formula(worksheet, row, 4, f"=IFERROR(SQRT(SUMPRODUCT(({asset_returns}<$B$1)*({asset_returns}-$B$1)^2)/COUNT({asset_returns})),\"\")", formats["percent"])
        _write_formula(worksheet, row, 5, f"=IFERROR({_cell(row, 4)}*SQRT(12),\"\")", formats["percent"])
        _write_formula(worksheet, row, 6, f"=IFERROR(-_xlfn.PERCENTILE.INC({asset_returns},0.05),\"\")", formats["percent"])
        _write_formula(worksheet, row, 7, f"=IFERROR(-AVERAGEIF({asset_returns},\"<=\"&-{var_95_cell},{asset_returns}),\"\")", formats["percent"])
        _write_formula(worksheet, row, 8, f"=IFERROR(-_xlfn.PERCENTILE.INC({asset_returns},0.01),\"\")", formats["percent"])
        _write_formula(worksheet, row, 9, f"=IFERROR(-AVERAGEIF({asset_returns},\"<=\"&-{var_99_cell},{asset_returns}),\"\")", formats["percent"])
        _write_formula(worksheet, row, 10, f"=IFERROR(MIN({drawdown_refs[ticker]['drawdown_range']}),\"\")", formats["percent"])
        _write_formula(worksheet, row, 11, f"=IFERROR(MAX({drawdown_refs[ticker]['duration_range']}),\"\")", formats["integer"])
        _write_formula(worksheet, row, 12, f"=IFERROR(LOOKUP(2,1/({drawdown_refs[ticker]['drawdown_range']}<>\"\"),{drawdown_refs[ticker]['drawdown_range']}),\"\")", formats["percent"])

    worksheet.set_column(0, 0, 18)
    worksheet.set_column(1, 1, 14, formats["integer"])
    worksheet.set_column(2, 10, 18, formats["percent"])
    worksheet.set_column(11, 11, 18, formats["integer"])
    worksheet.set_column(12, 12, 18, formats["percent"])
    worksheet.freeze_panes(3, 1)


def build_covariance_excel(
    tickers: Iterable[str],
    start_date: str,
    end_date: str,
    output_prefix: str = "covariance_matrix",
    minimum_acceptable_return: float = 0.0,
    output_dir: str | Path = ".",
) -> dict[str, object]:
    """Build a workbook where core return data is static and risk analytics are live Excel formulas."""
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

    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = output_path / f"{output_prefix}_{timestamp}.xlsx"

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book
        workbook.set_calc_mode("auto")
        formats = _configure_formats(workbook)
        date_format = workbook.add_format({"num_format": "mm/dd/yyyy"})
        percent_format = workbook.add_format({"num_format": "0.00%"})
        number_format = workbook.add_format({"num_format": "0.00"})

        cov_refs = _write_live_covariance_sheet(
            writer=writer,
            valid_tickers=valid_tickers,
            returns_last_row=len(monthly_log_returns.index),
            formats=formats,
        )
        _write_live_correlation_sheet(
            writer=writer,
            valid_tickers=valid_tickers,
            returns_last_row=len(monthly_log_returns.index),
            formats=formats,
        )
        drawdown_refs = _write_live_drawdown_sheet(
            writer=writer,
            monthly_simple_returns=monthly_simple_returns,
            valid_tickers=valid_tickers,
            formats=formats,
        )
        _write_live_downside_sheet(
            writer=writer,
            valid_tickers=valid_tickers,
            n_obs=len(monthly_simple_returns.index),
            minimum_acceptable_return=minimum_acceptable_return,
            drawdown_refs=drawdown_refs,
            formats=formats,
        )

        adj_close.to_excel(writer, sheet_name="Adj_Close_Daily")
        monthly_log_returns.to_excel(writer, sheet_name="Monthly_Log_Returns")
        monthly_simple_returns.to_excel(writer, sheet_name="Monthly_Simple_Returns")

        ws = writer.sheets["Adj_Close_Daily"]
        ws.set_column("A:A", 14, date_format)
        ws.set_column("B:ZZ", 14, number_format)
        ws.freeze_panes(1, 1)
        for sheet_name in ["Monthly_Log_Returns", "Monthly_Simple_Returns"]:
            ws = writer.sheets[sheet_name]
            ws.set_column("A:A", 14, date_format)
            ws.set_column("B:ZZ", 14, percent_format)
            ws.freeze_panes(1, 1)

    return {
        "adj_close_daily": adj_close,
        "monthly_log_returns": monthly_log_returns,
        "monthly_simple_returns": monthly_simple_returns,
        "monthly_log_returns_clean": monthly_log_returns_clean,
        "valid_tickers": valid_tickers,
        "failed_tickers": failed_tickers,
        "output_file": str(output_file),
        "formula_driven_outputs": [
            "Covariance_Matrix",
            "Correlation_Matrix",
            "Downside_Risk_Metrics",
            "Drawdown_Series",
        ],
        "covariance_references": cov_refs,
    }
