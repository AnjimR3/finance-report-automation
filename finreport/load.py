"""Load and validate the input CSV files.

Bad transactions never stop a run. They are set aside with a reason so they
can be reviewed on the report's Data Checks sheet.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

GROUPS = ("Revenue", "Cost of Goods Sold", "Operating Expenses")
TRANSACTION_COLUMNS = ["date", "description", "category", "amount"]


@dataclass
class TransactionData:
    """Result of loading a transactions file."""

    valid: pd.DataFrame
    rejected: pd.DataFrame
    rows_read: int


def _read_text_csv(path: str | Path, required: list[str]) -> pd.DataFrame:
    frame = pd.read_csv(path, dtype=str, keep_default_na=False)
    frame.columns = [column.strip().lower() for column in frame.columns]
    missing = [column for column in required if column not in frame.columns]
    if missing:
        raise ValueError(f"{path} is missing required columns: {', '.join(missing)}")
    return frame[required].apply(lambda column: column.str.strip())


def load_categories(path: str | Path) -> pd.DataFrame:
    """Load the mapping of each category to its P&L group."""
    categories = _read_text_csv(path, ["category", "group"])
    unknown = sorted(set(categories["group"]) - set(GROUPS))
    if unknown:
        raise ValueError(f"Unknown groups in {path}: {', '.join(unknown)}")
    duplicated = categories.loc[categories["category"].str.lower().duplicated(), "category"]
    if not duplicated.empty:
        raise ValueError(f"Duplicate categories in {path}: {', '.join(duplicated)}")
    return categories


def load_transactions(path: str | Path, categories: pd.DataFrame) -> TransactionData:
    """Load transactions, separating valid rows from rejected ones."""
    raw = _read_text_csv(path, TRANSACTION_COLUMNS)
    canonical = {name.lower(): name for name in categories["category"]}

    dates = pd.to_datetime(raw["date"], format="%Y-%m-%d", errors="coerce")
    amounts = pd.to_numeric(raw["amount"], errors="coerce")
    names = raw["category"].str.lower().map(canonical)

    reasons = pd.Series("", index=raw.index, dtype=object)
    checks = [
        (raw[TRANSACTION_COLUMNS].eq("").any(axis=1), "Missing required value"),
        (dates.isna(), "Invalid date"),
        (amounts.isna(), "Invalid amount"),
        (names.isna(), "Unknown category"),
        (raw.duplicated(keep="first"), "Possible duplicate"),
    ]
    for failed, reason in checks:
        reasons = reasons.where(~(failed & reasons.eq("")), reason)

    is_valid = reasons.eq("")
    valid = pd.DataFrame(
        {
            "date": dates[is_valid],
            "description": raw.loc[is_valid, "description"],
            "category": names[is_valid],
            "amount": amounts[is_valid],
        }
    ).sort_values("date", kind="stable").reset_index(drop=True)

    rejected = raw.loc[~is_valid].copy()
    rejected.insert(0, "source_row", rejected.index + 2)  # +2: header row, 1-based
    rejected["reason"] = reasons[~is_valid]
    return TransactionData(valid, rejected.reset_index(drop=True), len(raw))


def load_budget(path: str | Path, categories: pd.DataFrame) -> pd.DataFrame:
    """Load the monthly budget. Unlike transactions, a bad budget stops the run."""
    raw = _read_text_csv(path, ["month", "category", "amount"])
    months = pd.to_datetime(raw["month"], format="%Y-%m", errors="coerce")
    amounts = pd.to_numeric(raw["amount"], errors="coerce")
    if months.isna().any() or amounts.isna().any():
        raise ValueError(f"{path} has an invalid month or amount")
    unknown = sorted(set(raw["category"]) - set(categories["category"]))
    if unknown:
        raise ValueError(f"Unknown categories in {path}: {', '.join(unknown)}")
    return pd.DataFrame({"month": months, "category": raw["category"], "amount": amounts})
