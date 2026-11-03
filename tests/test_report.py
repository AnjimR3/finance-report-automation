"""Recalculate the workbook and compare its formulas with pandas on the same data."""

import shutil
import subprocess
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

from finreport.load import load_budget, load_categories, load_transactions
from finreport.report import build_report

ROOT = Path(__file__).resolve().parents[1]
OFFICE = shutil.which("soffice") or shutil.which("libreoffice")
pytestmark = pytest.mark.skipif(OFFICE is None, reason="LibreOffice is needed to recalculate formulas")


@pytest.fixture(scope="module")
def report(tmp_path_factory):
    folder = tmp_path_factory.mktemp("report")
    categories = load_categories(ROOT / "data/categories.csv")
    data = load_transactions(ROOT / "data/sample_transactions.csv", categories)
    budget = load_budget(ROOT / "data/sample_budget.csv", categories)
    built = build_report(data, budget, categories, folder / "built.xlsx")
    converted = folder / "converted"
    subprocess.run(
        [OFFICE, "--headless", f"-env:UserInstallation=file://{folder}/profile",
         "--convert-to", "xlsx", "--outdir", str(converted), str(built)],
        check=True, capture_output=True, timeout=180,
    )
    workbook = load_workbook(converted / "built.xlsx", data_only=True)
    return workbook, data, budget, categories


def find_row(sheet, label):
    return next(row for row in range(1, sheet.max_row + 1) if sheet.cell(row, 1).value == label)


def totals_by_group(data, categories):
    merged = data.valid.merge(categories, on="category")
    return merged.groupby("group")["amount"].sum()


def test_no_formula_errors(report):
    workbook, *_ = report
    for sheet in workbook:
        for row in sheet.iter_rows():
            for cell in row:
                assert not (isinstance(cell.value, str) and cell.value.startswith("#")), f"{sheet.title}!{cell.coordinate}"


def test_pnl_totals_match_pandas(report):
    workbook, data, _, categories = report
    sheet = workbook["Monthly P&L"]
    totals = totals_by_group(data, categories)
    column = sheet.max_column
    assert sheet.cell(find_row(sheet, "Total Revenue"), column).value == pytest.approx(totals["Revenue"])
    profit = totals["Revenue"] - totals["Cost of Goods Sold"] - totals["Operating Expenses"]
    assert sheet.cell(find_row(sheet, "Operating Profit"), column).value == pytest.approx(profit)


def test_every_monthly_cell_matches_pandas(report):
    workbook, data, _, categories = report
    sheet = workbook["Monthly P&L"]
    expected = data.valid.groupby([data.valid["date"].dt.to_period("M"), "category"])["amount"].sum()
    for name in categories["category"]:
        row = find_row(sheet, name)
        for column in range(2, sheet.max_column):
            month = pd.Timestamp(sheet.cell(4, column).value).to_period("M")
            assert sheet.cell(row, column).value == pytest.approx(expected.get((month, name), 0))


def test_budget_variance_sign(report):
    workbook, data, budget, categories = report
    sheet = workbook["Budget vs Actual"]
    actual = data.valid.groupby("category")["amount"].sum()
    planned = budget.groupby("category")["amount"].sum()
    revenue = find_row(sheet, "Product Sales")
    assert sheet.cell(revenue, 5).value == pytest.approx(actual["Product Sales"] - planned["Product Sales"])
    cost = find_row(sheet, "Marketing")
    assert sheet.cell(cost, 5).value == pytest.approx(planned["Marketing"] - actual["Marketing"])
    assert sheet.cell(cost, 7).value == "Unfavourable"


def test_data_checks_pass(report):
    workbook, data, *_ = report
    sheet = workbook["Data Checks"]
    assert sheet["B7"].value == "OK"
    assert sheet["B11"].value == "OK"
    assert sheet["B6"].value == len(data.rejected) == 5
    assert workbook["Summary"].cell(find_row(workbook["Summary"], "Data checks"), 2).value == "Passed"
