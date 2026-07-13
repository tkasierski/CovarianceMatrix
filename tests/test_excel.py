from pathlib import Path

import openpyxl
import pandas as pd

from covariance_matrix.analytics import analyze_returns
from covariance_matrix.excel import build_workbook


def test_workbook_uses_direct_ranges_and_dedicated_risk_free_rate(tmp_path: Path):
    returns = pd.DataFrame(
        {"A": [0.01, 0.02, -0.01], "B": [0.02, 0.01, 0.00]},
        index=pd.date_range("2024-01-31", periods=3, freq="ME"),
    )
    result = analyze_returns(returns, min_observations=2)
    output = tmp_path / "test.xlsx"
    build_workbook(result, output, risk_free_rate=0.04)
    wb = openpyxl.load_workbook(output, data_only=False)
    cov_formula = wb["Covariance_Matrix"]["B3"].value
    assert "INDEX" not in cov_formula
    assert "MATCH" not in cov_formula
    assert "Monthly_Simple_Returns" in cov_formula
    dashboard = wb["Portfolio_Dashboard"]
    assert dashboard["B3"].value == 0.04
    sharpe_formulas = [cell.value for row in dashboard.iter_rows() for cell in row if isinstance(cell.value, str) and "IFERROR((B" in cell.value]
    assert sharpe_formulas and "$B$3" in sharpe_formulas[0]


def test_validation_sheet_exists(tmp_path: Path):
    returns = pd.DataFrame({"A": [0.01, 0.02]}, index=pd.date_range("2024-01-31", periods=2, freq="ME"))
    output = tmp_path / "test.xlsx"
    build_workbook(analyze_returns(returns, min_observations=2), output)
    wb = openpyxl.load_workbook(output)
    assert "Validation" in wb.sheetnames
    assert "Observation_Counts" in wb.sheetnames
