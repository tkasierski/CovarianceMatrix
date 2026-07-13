from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
import streamlit as st

from covariance_matrix.core import build_covariance_excel
from covariance_matrix.data import parse_tickers

st.set_page_config(page_title="Covariance Matrix Tool", layout="wide")
st.title("Covariance Matrix + Portfolio Risk Tool")
st.write("Combine public-market prices with uploaded monthly simple returns and generate a formula-driven Excel workbook.")

with st.sidebar:
    start_date = st.date_input("Start date", value=pd.Timestamp("2018-01-01"))
    end_date = st.date_input("End date", value=pd.Timestamp.today())
    ticker_text = st.text_area("Public tickers", value="SPY\nGLD", help="Optional. Separate tickers with spaces, commas, or new lines.")
    custom_file = st.file_uploader("Custom monthly returns", type=["csv", "xlsx", "xls"], help="First column must be Date; remaining columns must contain decimal simple returns.")
    missing_method = st.selectbox("Missing-data method", ["listwise", "pairwise"], help="Listwise uses only months complete for every asset. Pairwise uses all overlapping months for each asset pair.")
    min_observations = st.number_input("Minimum observations", min_value=2, value=24, step=1)
    risk_free_rate = st.number_input("Annual risk-free rate", value=0.0, step=0.001, format="%.4f")
    minimum_acceptable_return = st.number_input("Minimum acceptable monthly return", value=0.0, step=0.001, format="%.4f")
    output_prefix = st.text_input("Output prefix", value="covariance_matrix")
    run = st.button("Build workbook", type="primary")

if run:
    tickers = parse_tickers(ticker_text)
    if not tickers and custom_file is None:
        st.error("Provide at least one ticker or a custom return file.")
    elif start_date >= end_date:
        st.error("Start date must precede end date.")
    else:
        try:
            with TemporaryDirectory() as temp_dir:
                payload = custom_file.getvalue() if custom_file is not None else None
                result = build_covariance_excel(
                    tickers=tickers,
                    start_date=start_date.strftime("%Y-%m-%d"),
                    end_date=end_date.strftime("%Y-%m-%d"),
                    custom_returns_source=payload,
                    custom_returns_filename=custom_file.name if custom_file is not None else None,
                    missing_data_method=missing_method,
                    min_observations=int(min_observations),
                    risk_free_rate=risk_free_rate,
                    minimum_acceptable_return=minimum_acceptable_return,
                    output_prefix=output_prefix,
                    output_dir=temp_dir,
                )
                output = Path(result["output_file"])
                workbook_bytes = output.read_bytes()
            st.success("Workbook created.")
            st.write("Assets: " + ", ".join(result["assets"]))
            if result["warnings"]:
                st.warning("\n".join(f"- {warning}" for warning in result["warnings"]))
            st.download_button("Download Excel workbook", workbook_bytes, file_name=output.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as exc:
            st.error(str(exc))
