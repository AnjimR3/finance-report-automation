import pandas as pd
import pytest

from finreport.load import load_budget, load_categories, load_transactions

CATEGORIES = pd.DataFrame(
    {"category": ["Sales", "Rent"], "group": ["Revenue", "Operating Expenses"]}
)


def write(tmp_path, text):
    path = tmp_path / "input.csv"
    path.write_text(text)
    return path


def test_valid_rows_are_loaded_and_sorted(tmp_path):
    path = write(tmp_path, "date,description,category,amount\n2025-02-01,B,Sales,20\n2025-01-01,A,Rent,10.5\n")
    data = load_transactions(path, CATEGORIES)
    assert list(data.valid["description"]) == ["A", "B"]
    assert data.valid["amount"].sum() == pytest.approx(30.5)
    assert data.rejected.empty


@pytest.mark.parametrize(
    ("row", "reason"),
    [
        ("2025-13-01,A,Sales,10", "Invalid date"),
        ("2025-01-01,A,Sales,abc", "Invalid amount"),
        ("2025-01-01,A,Unknown,10", "Unknown category"),
        ("2025-01-01,,Sales,10", "Missing required value"),
    ],
)
def test_bad_rows_are_rejected_with_a_reason(tmp_path, row, reason):
    data = load_transactions(write(tmp_path, f"date,description,category,amount\n{row}\n"), CATEGORIES)
    assert data.valid.empty
    assert data.rejected.loc[0, "reason"] == reason
    assert data.rejected.loc[0, "source_row"] == 2


def test_only_the_repeat_of_a_duplicate_is_rejected(tmp_path):
    row = "2025-01-01,A,Sales,10"
    data = load_transactions(write(tmp_path, f"date,description,category,amount\n{row}\n{row}\n"), CATEGORIES)
    assert len(data.valid) == 1
    assert data.rejected.loc[0, "reason"] == "Possible duplicate"


def test_category_matching_ignores_case_and_spaces(tmp_path):
    data = load_transactions(write(tmp_path, "date,description,category,amount\n2025-01-01,A, sales ,10\n"), CATEGORIES)
    assert data.valid.loc[0, "category"] == "Sales"


def test_refunds_are_kept_as_negative_amounts(tmp_path):
    data = load_transactions(write(tmp_path, "date,description,category,amount\n2025-01-01,Refund,Sales,-10\n"), CATEGORIES)
    assert data.valid.loc[0, "amount"] == -10


def test_missing_column_stops_the_run(tmp_path):
    with pytest.raises(ValueError, match="amount"):
        load_transactions(write(tmp_path, "date,description,category\n"), CATEGORIES)


def test_unknown_group_stops_the_run(tmp_path):
    with pytest.raises(ValueError, match="Unknown groups"):
        load_categories(write(tmp_path, "category,group\nSales,Income\n"))


def test_unknown_budget_category_stops_the_run(tmp_path):
    with pytest.raises(ValueError, match="Unknown categories"):
        load_budget(write(tmp_path, "month,category,amount\n2025-01,Other,10\n"), CATEGORIES)
