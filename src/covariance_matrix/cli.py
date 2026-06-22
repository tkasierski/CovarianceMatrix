from __future__ import annotations

import argparse
import sys
from pathlib import Path

from covariance_matrix.core import build_covariance_excel, parse_tickers


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="covariance-matrix",
        description="Generate covariance, correlation, downside-risk, and drawdown tables for public securities.",
    )
    parser.add_argument(
        "--tickers",
        nargs="+",
        required=True,
        help="Ticker symbols separated by spaces. Comma-separated values are also accepted within each argument.",
    )
    parser.add_argument("--start", required=True, help="Start date in YYYY-MM-DD format.")
    parser.add_argument("--end", required=True, help="End date in YYYY-MM-DD format.")
    parser.add_argument(
        "--output-prefix",
        default="covariance_matrix",
        help="Prefix for the generated Excel filename.",
    )
    parser.add_argument(
        "--output-dir",
        default=".",
        help="Directory where the generated Excel workbook should be written.",
    )
    parser.add_argument(
        "--minimum-acceptable-return",
        type=float,
        default=0.0,
        help="Minimum acceptable monthly return used for downside deviation. Example: 0.005 for 0.50%%.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    tickers = parse_tickers(args.tickers)

    try:
        results = build_covariance_excel(
            tickers=tickers,
            start_date=args.start,
            end_date=args.end,
            output_prefix=args.output_prefix,
            output_dir=Path(args.output_dir),
            minimum_acceptable_return=args.minimum_acceptable_return,
        )
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Excel file created: {results['output_file']}")
    print(f"Valid tickers: {', '.join(results['valid_tickers'])}")
    failed_tickers = results["failed_tickers"]
    if failed_tickers:
        print(f"Failed tickers: {', '.join(failed_tickers)}")
    else:
        print("Failed tickers: None")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
