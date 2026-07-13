from __future__ import annotations

from pathlib import Path

import pandas as pd
from xlsxwriter.utility import xl_col_to_name

from .analytics import AnalysisResult


def _direct_range(sheet: str, col: int, rows: int) -> str:
    letter = xl_col_to_name(col)
    return f"'{sheet}'!${letter}$2:${letter}${rows + 1}"


def build_workbook(
    result: AnalysisResult,
    output_file: str | Path,
    risk_free_rate: float = 0.0,
    minimum_acceptable_return: float = 0.0,
    adjusted_close: pd.DataFrame | None = None,
) -> str:
    output_file = str(output_file)
    returns = result.returns
    assets = list(returns.columns)
    n_assets = len(assets)
    n_rows = len(returns)

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book
        workbook.set_calc_mode("auto")
        fmt_header = workbook.add_format({"bold": True, "bg_color": "#E7E6E6", "border": 1})
        fmt_pct = workbook.add_format({"num_format": "0.00%", "border": 1})
        fmt_num = workbook.add_format({"num_format": "0.0000", "border": 1})
        fmt_money = workbook.add_format({"num_format": "$#,##0", "border": 1})
        fmt_input = workbook.add_format({"num_format": "0.00%", "bg_color": "#E2F0D9", "border": 1})
        fmt_date = workbook.add_format({"num_format": "mm/dd/yyyy"})

        returns.to_excel(writer, sheet_name="Monthly_Simple_Returns")
        ws_returns = writer.sheets["Monthly_Simple_Returns"]
        ws_returns.set_column(0, 0, 14, fmt_date)
        ws_returns.set_column(1, n_assets, 14, fmt_pct)
        ws_returns.freeze_panes(1, 1)

        if adjusted_close is not None and not adjusted_close.empty:
            adjusted_close.to_excel(writer, sheet_name="Adj_Close_Daily")

        for sheet_name, function_name, number_format in [
            ("Covariance_Matrix", "COVARIANCE.S", fmt_pct),
            ("Correlation_Matrix", "CORREL", fmt_num),
        ]:
            ws = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = ws
            ws.write(0, 0, sheet_name.replace("_", " "), fmt_header)
            for i, asset in enumerate(assets):
                ws.write(1, i + 1, asset, fmt_header)
                ws.write(i + 2, 0, asset, fmt_header)
            for i in range(n_assets):
                left = _direct_range("Monthly_Simple_Returns", i + 1, n_rows)
                for j in range(n_assets):
                    right = _direct_range("Monthly_Simple_Returns", j + 1, n_rows)
                    ws.write_formula(i + 2, j + 1, f'=IFERROR({function_name}({left},{right}),"")', number_format)
            ws.freeze_panes(2, 1)

        result.observation_counts.to_excel(writer, sheet_name="Observation_Counts")
        result.downside_metrics.to_excel(writer, sheet_name="Downside_Risk_Metrics")
        result.drawdown_series.to_excel(writer, sheet_name="Drawdown_Series")

        ws = workbook.add_worksheet("Portfolio_Dashboard")
        writer.sheets["Portfolio_Dashboard"] = ws
        ws.write(0, 0, "Missing-data method", fmt_header)
        ws.write(0, 1, result.missing_data_method)
        ws.write(1, 0, "Minimum observations", fmt_header)
        ws.write(1, 1, result.min_observations)
        ws.write(2, 0, "Annual risk-free rate", fmt_header)
        ws.write_number(2, 1, risk_free_rate, fmt_input)
        ws.write(3, 0, "Minimum acceptable monthly return", fmt_header)
        ws.write_number(3, 1, minimum_acceptable_return, fmt_input)
        headers = ["Asset", "$ Value", "Weight", "Expected Return", "Asset Risk", "Marginal Risk", "Risk Contribution", "% of Risk", "Weighted Covariance"]
        ws.write_row(5, 0, headers, fmt_header)
        cov_sheet = "Covariance_Matrix"
        first_row = 6
        total_row = first_row + n_assets
        for i, asset in enumerate(assets):
            row = first_row + i
            ws.write(row, 0, asset, fmt_header)
            ws.write_number(row, 1, 0, fmt_money)
            ws.write_formula(row, 2, f'=IF($B${total_row+1}=0,"",B{row+1}/$B${total_row+1})', fmt_pct)
            ws.write_blank(row, 3, None, fmt_input)
            diag = f"'{cov_sheet}'!{xl_col_to_name(i+1)}{i+3}"
            ws.write_formula(row, 4, f'=IFERROR(SQRT({diag}*12),"")', fmt_pct)
            weighted_terms = "+".join(
                f"'{cov_sheet}'!{xl_col_to_name(j+1)}{i+3}*$C${first_row+j+1}" for j in range(n_assets)
            )
            ws.write_formula(row, 8, f'=IFERROR(({weighted_terms})*12,"")', fmt_pct)
            ws.write_formula(row, 5, f'=IFERROR(I{row+1}/$B${total_row+4},"")', fmt_pct)
            ws.write_formula(row, 6, f'=IFERROR(C{row+1}*F{row+1},"")', fmt_pct)
            ws.write_formula(row, 7, f'=IFERROR(G{row+1}/$B${total_row+4},"")', fmt_pct)
        ws.write(total_row, 0, "Total", fmt_header)
        ws.write_formula(total_row, 1, f"=SUM(B{first_row+1}:B{total_row})", fmt_money)
        ws.write_formula(total_row, 2, f"=SUM(C{first_row+1}:C{total_row})", fmt_pct)
        ws.write(total_row + 2, 0, "Portfolio Return", fmt_header)
        ws.write_formula(total_row + 2, 1, f"=SUMPRODUCT(C{first_row+1}:C{total_row},D{first_row+1}:D{total_row})", fmt_pct)
        ws.write(total_row + 3, 0, "Portfolio Risk", fmt_header)
        ws.write_formula(total_row + 3, 1, f'=IFERROR(SQRT(SUMPRODUCT(C{first_row+1}:C{total_row},I{first_row+1}:I{total_row})),"")', fmt_pct)
        ws.write(total_row + 4, 0, "Sharpe Ratio", fmt_header)
        ws.write_formula(total_row + 4, 1, f'=IFERROR((B{total_row+3}-$B$3)/B{total_row+4},"")', fmt_num)
        ws.set_column(8, 8, None, None, {"hidden": True})

        validation = workbook.add_worksheet("Validation")
        writer.sheets["Validation"] = validation
        validation.write(0, 0, "Warnings", fmt_header)
        for row, warning in enumerate(result.warnings, start=1):
            validation.write(row, 0, warning)
        if not result.warnings:
            validation.write(1, 0, "No validation warnings.")

    return output_file
