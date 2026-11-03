"""Command line entry point: python -m finreport"""

from __future__ import annotations

import argparse

from .load import load_budget, load_categories, load_transactions
from .report import build_report


def main() -> None:
    parser = argparse.ArgumentParser(description="Build an Excel financial report from transaction CSV files.")
    parser.add_argument("--transactions", default="data/sample_transactions.csv")
    parser.add_argument("--budget", default="data/sample_budget.csv")
    parser.add_argument("--categories", default="data/categories.csv")
    parser.add_argument("--output", default="output/financial_report.xlsx")
    parser.add_argument("--company", default="Sample Co.")
    args = parser.parse_args()

    categories = load_categories(args.categories)
    data = load_transactions(args.transactions, categories)
    budget = load_budget(args.budget, categories)
    path = build_report(data, budget, categories, args.output, args.company)

    print(f"Rows read: {data.rows_read}")
    print(f"Rows loaded: {len(data.valid)}")
    print(f"Rows rejected: {len(data.rejected)}")
    for item in data.rejected.itertuples(index=False):
        print(f"  row {item.source_row}: {item.reason}")
    print(f"Report written to {path}")


if __name__ == "__main__":
    main()
