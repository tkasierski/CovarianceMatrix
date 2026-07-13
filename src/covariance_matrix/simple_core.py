from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
import yfinance as yf
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from covariance_matrix.core import (
    _cell,
    _configure_formats,
    _extract_adjusted_close,
    _quote_sheet,
    _range,
    _weighted_covariance_formula,
    _write_formula,
    _write_formula_matrix,
    _write_live_downside_sheet,
    _write_live_drawdown_sheet,
    parse_tickers,
)


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
    returns_sheet = "Monthly_Simple_Returns"
    return_headers = _range(returns_sheet, 0, 1, 0, n_assets)
    return_data = _range(returns_sheet, 1, 1, returns_last_row, n_assets)

    monthly_top_row = 1
    monthly_left_col = 0
    annual_left_col = n_assets + 3
    dashboard_top_row = n_assets + 7
    helper_base_col = max(annual_left_col + n_assets + 2, 20)

    worksheet.write(0, 0, "Monthly Covariance", formats["title"])
    worksheet.write(0, annual_left_col, "Annual Covariance", formats["title"])
    note = (
        "All risk calculations use the simple-return series shown in Monthly_Simple_Returns. "
        "To use custom public, private-asset, or hedge-fund return streams, replace values in "
        "that tab. Covariance, correlation, volatility, downside risk, and drawdowns will then "
        "recalculate from the same return convention. Paste net returns for net-risk analysis. "
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
                f'INDEX({return_data},0,MATCH({row_asset_cell},{return_headers},0)),'
                f'INDEX({return_data},0,MATCH({col_asset_cell},{return_headers},0))),"")'
            )
            monthly_row.append(formula)
            monthly_cell = xl_rowcol_to_cell(
                monthly_top_row + 1 + row_offset,
                monthly_left_col + 1 + col_offset,
            )
            annual_row.append(f'=IFERROR({monthly_cell}*12,"")')
        monthly_formulas.append(monthly_row)
        annual_formulas.append(annual_row)

    _write_formula_matrix(
        worksheet,
        monthly_top_row + 1,
        monthly_left_col + 1,
        monthly_formulas,
        formats["percent"],
    )
    _write_formula_matrix(
        worksheet,
        monthly_top_row + 1,
        annual_left_col + 1,
        annual_formulas,
        formats["percent"],
    )

    annual_matrix = (
        f"${xl_col_to_name(annual_left_col + 1)}${monthly_top_row + 2}:"
        f"${xl_col_to_name(annual_left_col + n_assets)}${monthly_top_row + 1 + n_assets}"
    )

    def write_dashboard(left_col: int, helper_col: int, title: str) -> dict[str, str]:
        worksheet.write(dashboard_top_row, left_col, title, formats["title"])
        headers = [
            "Asset",
            "$ Value",
            "Weight",
            "Expected Return",
            "Asset Risk",
            "Marginal Risk",
            "Risk Contribution",
            "% of Risk",
        ]
        worksheet.write_row(dashboard_top_row + 1, left_col, headers, formats["header"])
        worksheet.write(dashboard_top_row + 1, helper_col, "Weighted Covariance", formats["header"])

        first_asset_row = dashboard_top_row + 2
        last_asset_row = first_asset_row + n_assets - 1
        total_row = last_asset_row + 1
        return_row = total_row + 2
        risk_row = return_row + 1
        sharpe_row = risk_row + 1
        weights_range = (
            f"${xl_col_to_name(left_col + 2)}${first_asset_row + 1}:"
            f"${xl_col_to_name(left_col + 2)}${last_asset_row + 1}"
        )
        expected_range = (
            f"${xl_col_to_name(left_col + 3)}${first_asset_row + 1}:"
            f"${xl_col_to_name(left_col + 3)}${last_asset_row + 1}"
        )
        helper_range = (
            f"${xl_col_to_name(helper_col)}${first_asset_row + 1}:"
            f"${xl_col_to_name(helper_col)}${last_asset_row + 1}"
        )

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
                f'=IF({_cell(total_row, left_col + 1, absolute=True)}=0,"",'
                f'{_cell(row, left_col + 1)}/{_cell(total_row, left_col + 1, absolute=True)})',
                formats["percent"],
            )
            worksheet.write_blank(row, left_col + 3, None, formats["input_percent"])
            _write_formula(
                worksheet,
                row,
                left_col + 4,
                f'=IFERROR(SQRT(INDEX({annual_matrix},{asset_index + 1},{asset_index + 1})),"")',
                formats["percent"],
            )
            _write_formula(
                worksheet,
                row,
                helper_col,
                f'=IFERROR({weighted_cov_formula},"")',
                formats["percent"],
            )
            _write_formula(
                worksheet,
                row,
                left_col + 5,
                f'=IF({_cell(risk_row, left_col + 1, absolute=True)}="","",'
                f'{_cell(row, helper_col)}/{_cell(risk_row, left_col + 1, absolute=True)})',
                formats["percent"],
            )
            _write_formula(
                worksheet,
                row,
                left_col + 6,
                f'=IF(OR({_cell(row, left_col + 2)}="",{_cell(row, left_col + 5)}=""),"",'
                f'{_cell(row, left_col + 2)}*{_cell(row, left_col + 5)})',
                formats["percent"],
            )
            _write_formula(
                worksheet,
                row,
                left_col + 7,
                f'=IF({_cell(risk_row, left_col + 1, absolute=True)}="","",'
                f'{_cell(row, left_col + 6)}/{_cell(risk_row, left_col + 1, absolute=True)})',
                formats["percent"],
            )

        worksheet.write(total_row, left_col, "Total", formats["asset"])
        _write_formula(
            worksheet,
            total_row,
            left_col + 1,
            f"=SUM({_cell(first_asset_row, left_col + 1)}:{_cell(last_asset_row, left_col + 1)})",
            formats["currency"],
            0,
        )
        _write_formula(
            worksheet,
            total_row,
            left_col + 2,
            f"=SUM({_cell(first_asset_row, left_col + 2)}:{_cell(last_asset_row, left_col + 2)})",
            formats["percent"],
        )
        worksheet.write(return_row, left_col, "Return", formats["asset"])
        _write_formula(
            worksheet,
            return_row,
            left_col + 1,
            f"=SUMPRODUCT({weights_range},{expected_range})",
            formats["percent"],
        )
        worksheet.write(risk_row, left_col, "Risk", formats["asset"])
        _write_formula(
            worksheet,
            risk_row,
            left_col + 1,
            f'=IF(SUM({weights_range})=0,"",SQRT(SUMPRODUCT({weights_range},{helper_range})))',
            formats["percent"],
        )
        worksheet.write(sharpe_row, left_col, "Sharpe", formats["asset"])
        _write_formula(
            worksheet,
            sharpe_row,
            left_col + 1,
            f'=IF({_cell(risk_row, left_col + 1)}="","",'
            f'({_cell(return_row, left_col + 1)}-{_cell(first_asset_row, left_col + 3)})/'
            f'{_cell(risk_row, left_col + 1)})',
            formats["number"],
        )

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
        "monthly_matrix": (
            f"${xl_col_to_name(monthly_left_col + 1)}${monthly_top_row + 2}:"
            f"${xl_col_to_name(monthly_left_col + n_assets)}${monthly_top_row + 1 + n_assets}"
        ),
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
    returns_sheet = "Monthly_Simple_Returns"
    return_headers = _range(returns_sheet, 0, 1, 0, n_assets)
    return_data = _range(returns_sheet, 1, 1, returns_last_row, n_assets)

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
                f'INDEX({return_data},0,MATCH({row_asset_cell},{return_headers},0)),'
                f'INDEX({return_data},0,MATCH({col_asset_cell},{return_headers},0))),"")'
            )
        formulas.append(row)
    _write_formula_matrix(worksheet, 2, 1, formulas, formats["number"])
    worksheet.set_column(0, n_assets, 14)
    worksheet.freeze_panes(2, 1)


def build_covariance_excel(
    tickers: Iterable[str],
    start_date: str,
    end_date: str,
    output_prefix: str = "covariance_matrix",
    minimum_acceptable_return: float = 0.0,
    output_dir: str | Path = ".",
) -> dict[str, object]:
    """Build a formula-driven workbook using simple returns for every risk calculation."""
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
    monthly_simple_returns = monthly_prices.pct_change(fill_method=None).iloc[1:]
    monthly_simple_returns_clean = monthly_simple_returns.dropna(how="any")
    if monthly_simple_returns_clean.empty:
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
            returns_last_row=len(monthly_simple_returns.index),
            formats=formats,
        )
        _write_live_correlation_sheet(
            writer=writer,
            valid_tickers=valid_tickers,
            returns_last_row=len(monthly_simple_returns.index),
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
        monthly_simple_returns.to_excel(writer, sheet_name="Monthly_Simple_Returns")

        ws = writer.sheets["Adj_Close_Daily"]
        ws.set_column("A:A", 14, date_format)
        ws.set_column("B:ZZ", 14, number_format)
        ws.freeze_panes(1, 1)

        ws = writer.sheets["Monthly_Simple_Returns"]
        ws.set_column("A:A", 14, date_format)
        ws.set_column("B:ZZ", 14, percent_format)
        ws.freeze_panes(1, 1)

    return {
        "adj_close_daily": adj_close,
        "monthly_simple_returns": monthly_simple_returns,
        "monthly_simple_returns_clean": monthly_simple_returns_clean,
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
