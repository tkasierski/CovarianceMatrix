from __future__ import annotations

from pathlib import Path

import pandas as pd
from xlsxwriter.utility import xl_col_to_name

from .analytics import AnalysisResult


def _direct_range(sheet: str, col: int, rows: int) -> str:
    letter = xl_col_to_name(col)
    return f"'{sheet}'!${letter}$2:${letter}${rows + 1}"


def _write_live_drawdown_sheet(
    writer: pd.ExcelWriter,
    returns: pd.DataFrame,
    assets: list[str],
    fmt_header,
    fmt_pct,
    fmt_integer,
) -> dict[str, dict[str, str]]:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Drawdown_Series")
    writer.sheets["Drawdown_Series"] = worksheet
    n_rows = len(returns)
    n_assets = len(assets)

    worksheet.write(0, 0, "Date", fmt_header)
    for index, asset in enumerate(assets):
        worksheet.write(0, index + 1, asset, fmt_header)
    for row, date in enumerate(returns.index, start=1):
        worksheet.write_datetime(row, 0, date.to_pydatetime())

    helper_start = n_assets + 2
    references: dict[str, dict[str, str]] = {}
    for index, asset in enumerate(assets):
        return_col = index + 1
        drawdown_col = index + 1
        wealth_col = helper_start + index * 3
        peak_col = wealth_col + 1
        duration_col = wealth_col + 2
        worksheet.write(0, wealth_col, f"{asset} Wealth", fmt_header)
        worksheet.write(0, peak_col, f"{asset} Peak", fmt_header)
        worksheet.write(0, duration_col, f"{asset} Duration", fmt_header)

        for row in range(1, n_rows + 1):
            return_cell = f"'Monthly_Simple_Returns'!{xl_col_to_name(return_col)}{row + 1}"
            wealth_cell = f"{xl_col_to_name(wealth_col)}{row + 1}"
            peak_cell = f"{xl_col_to_name(peak_col)}{row + 1}"
            drawdown_cell = f"{xl_col_to_name(drawdown_col)}{row + 1}"
            if row == 1:
                wealth_formula = f'=IF({return_cell}="","",1+{return_cell})'
                peak_formula = f'=IF({wealth_cell}="","",{wealth_cell})'
                drawdown_formula = (
                    f'=IF({return_cell}="","",IF({wealth_cell}="","",{wealth_cell}/{peak_cell}-1))'
                )
                duration_formula = f'=IF({return_cell}="","",IF({drawdown_cell}<0,1,0))'
            else:
                prior_wealth = f"{xl_col_to_name(wealth_col)}{row}"
                prior_peak = f"{xl_col_to_name(peak_col)}{row}"
                prior_duration = f"{xl_col_to_name(duration_col)}{row}"
                wealth_formula = (
                    f'=IF({return_cell}="",{prior_wealth},'
                    f'IF({prior_wealth}="",1+{return_cell},{prior_wealth}*(1+{return_cell})))'
                )
                peak_formula = f'=IF({wealth_cell}="",{prior_peak},MAX({prior_peak},{wealth_cell}))'
                drawdown_formula = (
                    f'=IF({return_cell}="","",IF({wealth_cell}="","",{wealth_cell}/{peak_cell}-1))'
                )
                duration_formula = (
                    f'=IF({return_cell}="",{prior_duration},'
                    f'IF({drawdown_cell}<0,{prior_duration}+1,0))'
                )
            worksheet.write_formula(row, wealth_col, wealth_formula, fmt_pct)
            worksheet.write_formula(row, peak_col, peak_formula, fmt_pct)
            worksheet.write_formula(row, drawdown_col, drawdown_formula, fmt_pct)
            worksheet.write_formula(row, duration_col, duration_formula, fmt_integer)

        references[asset] = {
            "drawdown": (
                f"'Drawdown_Series'!${xl_col_to_name(drawdown_col)}$2:"
                f"${xl_col_to_name(drawdown_col)}${n_rows + 1}"
            ),
            "duration": (
                f"'Drawdown_Series'!${xl_col_to_name(duration_col)}$2:"
                f"${xl_col_to_name(duration_col)}${n_rows + 1}"
            ),
        }

    worksheet.set_column(
        helper_start,
        helper_start + n_assets * 3 - 1,
        None,
        None,
        {"hidden": True},
    )
    worksheet.freeze_panes(1, 1)
    return references


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
        fmt_integer = workbook.add_format({"num_format": "0", "border": 1})

        returns.to_excel(writer, sheet_name="Monthly_Simple_Returns")
        ws_returns = writer.sheets["Monthly_Simple_Returns"]
        ws_returns.set_column(0, 0, 14, fmt_date)
        ws_returns.set_column(1, n_assets, 14, fmt_pct)
        ws_returns.freeze_panes(1, 1)

        if adjusted_close is not None and not adjusted_close.empty:
            adjusted_close.to_excel(writer, sheet_name="Adj_Close_Daily")

        for sheet_name, function_name, title, number_format in [
            ("Covariance_Matrix", "_xlfn.COVARIANCE.S", "Monthly Covariance Matrix", fmt_pct),
            ("Correlation_Matrix", "CORREL", "Correlation Matrix", fmt_num),
        ]:
            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet
            worksheet.write(0, 0, title, fmt_header)
            for index, asset in enumerate(assets):
                worksheet.write(1, index + 1, asset, fmt_header)
                worksheet.write(index + 2, 0, asset, fmt_header)
            for left_index in range(n_assets):
                left_range = _direct_range("Monthly_Simple_Returns", left_index + 1, n_rows)
                for right_index in range(n_assets):
                    right_range = _direct_range("Monthly_Simple_Returns", right_index + 1, n_rows)
                    worksheet.write_formula(
                        left_index + 2,
                        right_index + 1,
                        f'=IFERROR({function_name}({left_range},{right_range}),"")',
                        number_format,
                    )
            worksheet.freeze_panes(2, 1)

        result.observation_counts.to_excel(writer, sheet_name="Observation_Counts")

        dashboard = workbook.add_worksheet("Portfolio_Dashboard")
        writer.sheets["Portfolio_Dashboard"] = dashboard
        dashboard.write(0, 0, "Missing-data method", fmt_header)
        dashboard.write(0, 1, result.missing_data_method)
        dashboard.write(1, 0, "Minimum observations", fmt_header)
        dashboard.write(1, 1, result.min_observations)
        dashboard.write(2, 0, "Annual risk-free rate", fmt_header)
        dashboard.write_number(2, 1, risk_free_rate, fmt_input)
        dashboard.write(3, 0, "Minimum acceptable monthly return", fmt_header)
        dashboard.write_number(3, 1, minimum_acceptable_return, fmt_input)
        headers = [
            "Asset",
            "$ Value",
            "Weight",
            "Expected Return",
            "Asset Risk",
            "Marginal Risk",
            "Risk Contribution",
            "% of Risk",
            "Weighted Covariance",
        ]
        dashboard.write_row(5, 0, headers, fmt_header)
        first_row = 6
        total_row = first_row + n_assets
        for index, asset in enumerate(assets):
            row = first_row + index
            dashboard.write(row, 0, asset, fmt_header)
            dashboard.write_number(row, 1, 0, fmt_money)
            dashboard.write_formula(
                row,
                2,
                f'=IF($B${total_row + 1}=0,"",B{row + 1}/$B${total_row + 1})',
                fmt_pct,
            )
            dashboard.write_blank(row, 3, None, fmt_input)
            diagonal_cell = f"'Covariance_Matrix'!{xl_col_to_name(index + 1)}{index + 3}"
            dashboard.write_formula(row, 4, f'=IFERROR(SQRT({diagonal_cell}*12),"")', fmt_pct)
            weighted_terms = "+".join(
                f"'Covariance_Matrix'!{xl_col_to_name(other + 1)}{index + 3}*"
                f"$C${first_row + other + 1}"
                for other in range(n_assets)
            )
            dashboard.write_formula(row, 8, f'=IFERROR(({weighted_terms})*12,"")', fmt_pct)
            dashboard.write_formula(row, 5, f'=IFERROR(I{row + 1}/$B${total_row + 4},"")', fmt_pct)
            dashboard.write_formula(row, 6, f'=IFERROR(C{row + 1}*F{row + 1},"")', fmt_pct)
            dashboard.write_formula(row, 7, f'=IFERROR(G{row + 1}/$B${total_row + 4},"")', fmt_pct)
        dashboard.write(total_row, 0, "Total", fmt_header)
        dashboard.write_formula(total_row, 1, f"=SUM(B{first_row + 1}:B{total_row})", fmt_money)
        dashboard.write_formula(total_row, 2, f"=SUM(C{first_row + 1}:C{total_row})", fmt_pct)
        dashboard.write(total_row + 2, 0, "Portfolio Return", fmt_header)
        dashboard.write_formula(
            total_row + 2,
            1,
            f"=SUMPRODUCT(C{first_row + 1}:C{total_row},D{first_row + 1}:D{total_row})",
            fmt_pct,
        )
        dashboard.write(total_row + 3, 0, "Portfolio Risk", fmt_header)
        dashboard.write_formula(
            total_row + 3,
            1,
            f'=IFERROR(SQRT(SUMPRODUCT(C{first_row + 1}:C{total_row},'
            f'I{first_row + 1}:I{total_row})),"")',
            fmt_pct,
        )
        dashboard.write(total_row + 4, 0, "Sharpe Ratio", fmt_header)
        dashboard.write_formula(
            total_row + 4,
            1,
            f'=IFERROR((B{total_row + 3}-$B$3)/B{total_row + 4},"")',
            fmt_num,
        )
        dashboard.set_column(8, 8, None, None, {"hidden": True})

        drawdown_refs = _write_live_drawdown_sheet(
            writer,
            returns,
            assets,
            fmt_header,
            fmt_pct,
            fmt_integer,
        )

        downside = workbook.add_worksheet("Downside_Risk_Metrics")
        writer.sheets["Downside_Risk_Metrics"] = downside
        downside_headers = [
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
        downside.write_row(0, 0, downside_headers, fmt_header)
        mar_cell = "'Portfolio_Dashboard'!$B$4"
        for index, asset in enumerate(assets):
            row = index + 1
            return_range = _direct_range("Monthly_Simple_Returns", index + 1, n_rows)
            drawdown_range = drawdown_refs[asset]["drawdown"]
            duration_range = drawdown_refs[asset]["duration"]
            downside.write(row, 0, asset, fmt_header)
            downside.write_formula(row, 1, f"=COUNT({return_range})", fmt_integer)
            downside.write_formula(row, 2, f'=IFERROR(STDEV.S({return_range}),"")', fmt_pct)
            downside.write_formula(row, 3, f'=IFERROR(C{row + 1}*SQRT(12),"")', fmt_pct)
            downside.write_formula(
                row,
                4,
                f'=IFERROR(SQRT(SUMPRODUCT(--ISNUMBER({return_range}),'
                f'--({return_range}<{mar_cell}),({return_range}-{mar_cell})^2)'
                f'/COUNT({return_range})),"")',
                fmt_pct,
            )
            downside.write_formula(row, 5, f'=IFERROR(E{row + 1}*SQRT(12),"")', fmt_pct)
            downside.write_formula(row, 6, f'=IFERROR(-PERCENTILE({return_range},0.05),"")', fmt_pct)
            downside.write_formula(
                row,
                7,
                f'=IFERROR(-SUMIF({return_range},"<="&PERCENTILE({return_range},0.05),'
                f'{return_range})/COUNTIF({return_range},"<="&PERCENTILE({return_range},0.05)),"")',
                fmt_pct,
            )
            downside.write_formula(row, 8, f'=IFERROR(-PERCENTILE({return_range},0.01),"")', fmt_pct)
            downside.write_formula(
                row,
                9,
                f'=IFERROR(-SUMIF({return_range},"<="&PERCENTILE({return_range},0.01),'
                f'{return_range})/COUNTIF({return_range},"<="&PERCENTILE({return_range},0.01)),"")',
                fmt_pct,
            )
            downside.write_formula(row, 10, f'=IFERROR(MIN({drawdown_range}),"")', fmt_pct)
            downside.write_formula(row, 11, f'=IFERROR(MAX({duration_range}),"")', fmt_integer)
            downside.write_formula(
                row,
                12,
                f'=IFERROR(INDEX({drawdown_range},'
                f'MATCH(9.99999999999999E+307,{return_range})),"")',
                fmt_pct,
            )
        downside.freeze_panes(1, 1)

        validation = workbook.add_worksheet("Validation")
        writer.sheets["Validation"] = validation
        validation.write(0, 0, "Warnings", fmt_header)
        for row, warning in enumerate(result.warnings, start=1):
            validation.write(row, 0, warning)
        if not result.warnings:
            validation.write(1, 0, "No validation warnings.")

    return output_file
