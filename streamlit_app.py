from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
import streamlit as st

from covariance_matrix.core import build_covariance_excel, parse_tickers


st.set_page_config(page_title="Covariance Matrix Tool", layout="wide")

st.title("Covariance Matrix + Downside Risk Tool")
st.write(
    "Generate a formula-driven Excel workbook for covariance, correlation, downside-risk, "
    "drawdown, and portfolio risk-dashboard analysis using Yahoo Finance data."
)
st.info(
    "The downloaded workbook writes adjusted daily close, monthly simple returns, and monthly "
    "log returns as static data. Covariance, correlation, downside-risk metrics, drawdowns, "
    "and dashboard outputs calculate live in Excel from the return tabs."
)

with st.sidebar:
    st.header("Inputs")
    start_date = st.date_input("Start date", value=pd.Timestamp("2018-01-01"))
    end_date = st.date_input("End date", value=pd.Timestamp("2025-12-31"))
    ticker_text = st.text_area(
        "Tickers",
        value="AAPL\nMSFT\nGOOGL\nAMZN\nMETA\nSPY",
        help="Enter tickers separated by commas, spaces, or new lines.",
        height=180,
    )
    minimum_acceptable_return = st.number_input(
        "Minimum acceptable monthly return",
        min_value=-1.0,
        max_value=1.0,
        value=0.0,
        step=0.001,
        format="%.4f",
    )
    output_prefix = st.text_input("Output file prefix", value="covariance_matrix")
    run_button = st.button("Run analysis", type="primary")

if run_button:
    tickers = parse_tickers(ticker_text)

    if not tickers:
        st.error("Enter at least one ticker.")
    elif start_date >= end_date:
        st.error("Start date must be before end date.")
    else:
        with st.spinner("Pulling market data and building workbook..."):
            try:
                with TemporaryDirectory() as temp_dir:
                    results = build_covariance_excel(
                        tickers=tickers,
                        start_date=start_date.strftime("%Y-%m-%d"),
                        end_date=end_date.strftime("%Y-%m-%d"),
                        output_prefix=output_prefix,
                        minimum_acceptable_return=minimum_acceptable_return,
                        output_dir=temp_dir,
                    )
                    output_file = Path(results["output_file"])
                    workbook_bytes = output_file.read_bytes()

                st.success("Workbook created.")

                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Valid tickers")
                    st.write(", ".join(results["valid_tickers"]) or "None")
                with col2:
                    st.subheader("Failed tickers")
                    failed_tickers = results["failed_tickers"]
                    st.write(", ".join(failed_tickers) if failed_tickers else "None")

                st.subheader("Workbook outputs")
                st.write(
                    "Open the downloaded workbook in Excel to use the live covariance matrix, "
                    "annualized covariance matrix, correlation matrix, downside-risk metrics, "
                    "drawdown series, and allocation dashboard."
                )
                formula_outputs = results.get("formula_driven_outputs", [])
                if formula_outputs:
                    st.write("Formula-driven sheets: " + ", ".join(formula_outputs))

                st.download_button(
                    label="Download Excel workbook",
                    data=workbook_bytes,
                    file_name=output_file.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as exc:
                st.error(f"Error: {exc}")
else:
    st.info("Enter inputs in the sidebar and click Run analysis.")
