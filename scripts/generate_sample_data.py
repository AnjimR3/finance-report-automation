"""Generate the fictional sample data in data/. Results repeat for a given seed.

The company and every figure are invented. A few bad rows are added on purpose
to show the validation step.
"""

from __future__ import annotations

import csv
import random
from datetime import date
from pathlib import Path

DATA = Path(__file__).resolve().parents[1] / "data"
YEAR = 2025

# category: (monthly budget, transactions per month, descriptions)
PLAN = {
    "Product Sales": (82000, 6, ["Online store payout", "Wholesale order", "Marketplace payout"]),
    "Service Revenue": (24000, 3, ["Support contract", "Implementation project"]),
    "Materials": (29000, 4, ["Supplier invoice", "Packaging order"]),
    "Shipping": (7500, 3, ["Carrier invoice", "Freight charge"]),
    "Salaries": (31000, 2, ["Payroll run"]),
    "Rent": (5500, 1, ["Office rent"]),
    "Marketing": (6500, 3, ["Online ads", "Trade show", "Design agency"]),
    "Software": (2200, 2, ["Accounting software", "Cloud services"]),
    "Travel": (1800, 2, ["Client visit", "Conference travel"]),
    "Professional Fees": (2000, 1, ["Accounting fees", "Legal fees"]),
}
SEASONALITY = [0.90, 0.88, 0.97, 1.00, 1.03, 1.06, 0.98, 0.96, 1.04, 1.08, 1.17, 1.25]
FIXED = {"Salaries", "Rent", "Software"}
OVER_BUDGET = {"Marketing": 1.22, "Shipping": 1.11}


def split(total: float, parts: int, rng: random.Random) -> list[float]:
    weights = [rng.uniform(0.7, 1.3) for _ in range(parts)]
    return [round(total * weight / sum(weights), 2) for weight in weights]


def main() -> None:
    rng = random.Random(42)
    transactions, budget = [], []
    for month, season in enumerate(SEASONALITY, start=1):
        for category, (monthly, count, descriptions) in PLAN.items():
            factor = 1.0 if category in FIXED else season
            budget.append([f"{YEAR}-{month:02d}", category, round(monthly * factor, -1)])
            actual = monthly * factor * rng.uniform(0.94, 1.06) * OVER_BUDGET.get(category, 1.0)
            for amount in split(actual, count, rng):
                day = date(YEAR, month, rng.randint(1, 28)).isoformat()
                transactions.append([day, rng.choice(descriptions), category, amount])
    transactions.sort(key=lambda row: row[0])

    transactions += [
        ["2025-13-05", "Online ads", "Marketing", 410.00],
        ["2025-06-14", "Carrier invoice", "Shipping", "n/a"],
        ["2025-08-02", "Office snacks", "Refreshments", 86.50],
        ["2025-09-19", "", "Travel", 220.00],
        list(transactions[20]),
    ]

    with open(DATA / "sample_transactions.csv", "w", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["date", "description", "category", "amount"])
        writer.writerows(transactions)
    with open(DATA / "sample_budget.csv", "w", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["month", "category", "amount"])
        writer.writerows(budget)


if __name__ == "__main__":
    main()
