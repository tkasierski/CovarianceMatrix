from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .core import build_covariance_excel


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate covariance and portfolio-risk analysis.")
    parser.add_argument("--tickers", nargs="*", default=[])
    parser.add_argument("--custom-returns", help="CSV or Excel file with Date in the first column and simple returns in remaining columns.")
    parser.add_argument("--start", default="2018-01-01")
    parser.add_argument("--end", default="2025-12-31")
    parser.add_argument("--output-prefix", default="covariance_matrix")
    parser.add_argument("--output-dir", default=".")
    parser.add_argument("--minimum-acceptable-return", type=float, default=0.0)
    parser.add_argument("--risk-free-rate", type=float, default=0.0)
    parser.add_argument("--missing-data-method", choices=["listwise", "pairwise"], default="listwise")
    parser.add_argument("--min-observations", type=int, default=24)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if not args.tickers and not args.custom_returns:
        print("Error: provide --tickers, --custom-returns, or both.", file=sys.stderr)
        return 1
    try:
        result = build_covariance_excel(
            tickers=args.tickers,
            start_date=args.start,
            end_date=args.end,
            output_prefix=args.output_prefix,
            output_dir=Path(args.output_dir),
            minimum_acceptable_return=args.minimum_acceptable_return,
            risk_free_rate=args.risk_free_rate,
            missing_data_method=args.missing_data_method,
            min_observations=args.min_observations,
            custom_returns_source=args.custom_returns,
            custom_returns_filename=args.custom_returns,
        )
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(f"Excel file created: {result['output_file']}")
    if result["warnings"]:
        print("Warnings:")
        for warning in result["warnings"]:
            print(f"- {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
