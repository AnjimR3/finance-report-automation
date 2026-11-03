"""Build the Excel report.

Every figure in the analysis sheets is a live Excel formula that traces back to
the Transactions and Budget sheets, so a reviewer can audit any number.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.worksheet import Worksheet

from .load import GROUPS, TransactionData

FONT = "Arial"
NAVY = "1F3864"
LIGHT = "DCE6F1"
MONEY = "$#,##0;($#,##0);-"
PERCENT = "0.0%;(0.0%);-"
MONTH = "mmm yyyy"
BLUE_INPUT = "0000FF"
GREEN_LINK = "008000"
PNL = "'Monthly P&L'"
BVA = "'Budget vs Actual'"


def _font(**kwargs) -> Font:
    return Font(name=FONT, size=kwargs.pop("size", 10), **kwargs)


def _title(ws: Worksheet, title: str, subtitle: str) -> None:
    ws["A1"] = title
    ws["A1"].font = _font(size=14, bold=True, color=NAVY)
    ws["A2"] = subtitle
    ws["A2"].font = _font(italic=True, color="595959")
    ws.sheet_view.showGridLines = False


def _header(ws: Worksheet, row: int, labels: list) -> None:
    for index, label in enumerate(labels, start=1):
        cell = ws.cell(row, index, label)
        cell.font = _font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(
            horizontal="left" if index == 1 else "right", vertical="center", wrap_text=True
        )


def _emphasise(ws: Worksheet, row: int, columns: int) -> None:
    for index in range(1, columns + 1):
        cell = ws.cell(row, index)
        color = cell.font.color.rgb if cell.font.color else None
        cell.font = _font(bold=True, color=color[-6:] if isinstance(color, str) else "000000")
        cell.fill = PatternFill("solid", fgColor=LIGHT)
        cell.border = Border(top=Side(style="thin", color=NAVY))


def _widths(ws: Worksheet, widths: dict[str, float]) -> None:
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width


def _page_setup(ws: Worksheet) -> None:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def _write_transactions(ws: Worksheet, valid: pd.DataFrame) -> None:
    _header(ws, 1, ["Date", "Description", "Category", "Amount", "Month"])
    for row, item in enumerate(valid.itertuples(index=False), start=2):
        ws.cell(row, 1, item.date.to_pydatetime()).number_format = "yyyy-mm-dd"
        ws.cell(row, 2, item.description)
        ws.cell(row, 3, item.category)
        ws.cell(row, 4, float(item.amount)).number_format = "$#,##0.00;($#,##0.00);-"
        ws.cell(row, 5, f"=DATE(YEAR(A{row}),MONTH(A{row}),1)").number_format = MONTH
        for column in range(1, 5):
            ws.cell(row, column).font = _font(color=BLUE_INPUT)
        ws.cell(row, 5).font = _font()
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{max(len(valid) + 1, 2)}"
    _widths(ws, {"A": 12, "B": 40, "C": 20, "D": 14, "E": 12})


def _write_budget(ws: Worksheet, budget: pd.DataFrame) -> None:
    _header(ws, 1, ["Month", "Category", "Budget"])
    for row, item in enumerate(budget.itertuples(index=False), start=2):
        ws.cell(row, 1, item.month.to_pydatetime()).number_format = MONTH
        ws.cell(row, 2, item.category)
        ws.cell(row, 3, float(item.amount)).number_format = MONEY
        for column in range(1, 4):
            ws.cell(row, column).font = _font(color=BLUE_INPUT)
    ws.freeze_panes = "A2"
    _widths(ws, {"A": 12, "B": 22, "C": 14})


def _write_pnl(ws: Worksheet, company: str, months: list, categories: pd.DataFrame) -> dict:
    """Write the monthly P&L and return the rows and columns other sheets use."""
    _title(ws, f"{company}: Monthly P&L", "All amounts in USD. Each cell sums the Transactions sheet.")
    header_row = 4
    _header(ws, header_row, ["Line item", *[m.to_pydatetime() for m in months], "Total"])
    first, last = 2, 1 + len(months)
    total_column = last + 1
    for column in range(first, last + 1):
        ws.cell(header_row, column).number_format = MONTH

    rows: dict[str, int] = {}
    row = header_row + 1

    def add_formula_row(label: str, build, number_format: str = MONEY, emphasise: bool = True) -> None:
        nonlocal row
        ws.cell(row, 1, label).font = _font()
        for column in range(first, total_column + 1):
            cell = ws.cell(row, column, build(col(column)))
            cell.number_format = number_format
            cell.font = _font()
        if emphasise:
            _emphasise(ws, row, total_column)
        rows[label] = row
        row += 1

    for group in GROUPS:
        ws.cell(row, 1, group).font = _font(bold=True, color=NAVY)
        row += 1
        names = categories.loc[categories["group"] == group, "category"]
        start = row
        for name in names:
            cell = ws.cell(row, 1, name)
            cell.font = _font()
            cell.alignment = Alignment(indent=1)
            for column in range(first, last + 1):
                letter = col(column)
                formula = (
                    f"=SUMIFS(Transactions!$D:$D,Transactions!$C:$C,$A{row},"
                    f"Transactions!$E:$E,{letter}${header_row})"
                )
                target = ws.cell(row, column, formula)
                target.number_format = MONEY
                target.font = _font()
            total = ws.cell(row, total_column, f"=SUM({col(first)}{row}:{col(last)}{row})")
            total.number_format = MONEY
            total.font = _font(bold=True)
            rows[name] = row
            row += 1
        end = row - 1
        add_formula_row(f"Total {group}", lambda c, s=start, e=end: f"=SUM({c}{s}:{c}{e})")
        if group == "Cost of Goods Sold":
            add_formula_row(
                "Gross Profit",
                lambda c: f"={c}{rows['Total Revenue']}-{c}{rows['Total Cost of Goods Sold']}",
            )
        row += 1

    add_formula_row(
        "Operating Profit",
        lambda c: f"={c}{rows['Gross Profit']}-{c}{rows['Total Operating Expenses']}",
    )
    row += 1
    margin = lambda name: lambda c: f"=IF({c}{rows['Total Revenue']}=0,0,{c}{rows[name]}/{c}{rows['Total Revenue']})"
    add_formula_row("Gross Margin", margin("Gross Profit"), PERCENT, emphasise=False)
    add_formula_row("Operating Margin", margin("Operating Profit"), PERCENT, emphasise=False)

    ws.freeze_panes = ws.cell(header_row + 1, 2)
    _widths(ws, {"A": 28, **{col(c): 12 for c in range(first, total_column + 1)}})
    _page_setup(ws)
    return {"rows": rows, "first": first, "last": last, "total": total_column, "header": header_row}


def _write_bva(ws: Worksheet, company: str, categories: pd.DataFrame, pnl: dict) -> dict:
    _title(
        ws,
        f"{company}: Budget vs Actual",
        "Positive variance is favourable: revenue above budget, or costs below budget.",
    )
    header_row = 4
    _header(ws, header_row, ["Category", "Group", "Budget", "Actual", "Variance", "Variance %", "Status"])
    ws.cell(header_row, 2).alignment = Alignment(horizontal="left")
    ws.cell(header_row, 7).alignment = Alignment(horizontal="left", indent=1)
    total_letter = col(pnl["total"])

    ordered = pd.concat(categories[categories["group"] == group] for group in GROUPS)
    first = header_row + 1
    last = first + len(ordered) - 1

    def add_measures(row: int) -> None:
        ws.cell(row, 5).number_format = MONEY
        ws.cell(row, 6, f"=IF(C{row}=0,0,E{row}/ABS(C{row}))").number_format = PERCENT
        ws.cell(row, 7, f'=IF(E{row}>=0,"Favourable","Unfavourable")').alignment = Alignment(indent=1)
        for column in (3, 4):
            ws.cell(row, column).number_format = MONEY

    for row, item in enumerate(ordered.itertuples(index=False), start=first):
        ws.cell(row, 1, item.category)
        ws.cell(row, 2, item.group)
        ws.cell(row, 3, f"=SUMIFS(Budget!$C:$C,Budget!$B:$B,$A{row})")
        link = ws.cell(row, 4, f"={PNL}!{total_letter}{pnl['rows'][item.category]}")
        link.font = _font(color=GREEN_LINK)
        sign = "D{0}-C{0}" if item.group == "Revenue" else "C{0}-D{0}"
        ws.cell(row, 5, "=" + sign.format(row))
        add_measures(row)
        for column in (1, 2, 3, 5, 6, 7):
            ws.cell(row, column).font = _font()
        ws.cell(row, 7).alignment = Alignment(indent=1)

    rows: dict[str, int] = {}
    row = last + 1
    for group in GROUPS:
        ws.cell(row, 1, f"Total {group}")
        for letter in "CD":
            ws[f"{letter}{row}"] = f'=SUMIFS({letter}${first}:{letter}${last},$B${first}:$B${last},"{group}")'
        ws.cell(row, 5, f"=D{row}-C{row}" if group == "Revenue" else f"=C{row}-D{row}")
        add_measures(row)
        _emphasise(ws, row, 7)
        rows[group] = row
        row += 1
    ws.cell(row, 1, "Operating Profit")
    for letter in "CD":
        ws[f"{letter}{row}"] = (
            f"={letter}{rows['Revenue']}-{letter}{rows['Cost of Goods Sold']}"
            f"-{letter}{rows['Operating Expenses']}"
        )
    ws.cell(row, 5, f"=D{row}-C{row}")
    add_measures(row)
    _emphasise(ws, row, 7)
    rows["Operating Profit"] = row

    status_range = f"G{first}:G{row}"
    ws.conditional_formatting.add(status_range, CellIsRule(operator="equal", formula=['"Unfavourable"'], font=Font(color="C00000")))
    ws.conditional_formatting.add(status_range, CellIsRule(operator="equal", formula=['"Favourable"'], font=Font(color="007A33")))
    ws.freeze_panes = ws.cell(first, 2)
    _widths(ws, {"A": 28, "B": 22, "C": 13, "D": 13, "E": 13, "F": 12, "G": 14})
    _page_setup(ws)
    return {"first": first, "last": last, "rows": rows}


def _write_checks(ws: Worksheet, company: str, data: TransactionData, pnl: dict) -> None:
    _title(ws, f"{company}: Data Checks", "Confirms that nothing was lost between the source file and the report.")
    total = col(pnl["total"])
    pnl_rows = pnl["rows"]
    source_total = float(data.valid["amount"].sum())
    lines = [
        ("Rows read from source file", data.rows_read, "#,##0", True),
        ("Rows loaded", "=COUNT(Transactions!$D:$D)", "#,##0", False),
        ("Rows rejected", len(data.rejected), "#,##0;;0", True),
        ("Row count check", '=IF(B4=B5+B6,"OK","CHECK")', None, False),
        ("Sum of valid amounts in source file", source_total, "$#,##0.00;($#,##0.00);-", True),
        ("Sum of Transactions sheet", "=SUM(Transactions!$D:$D)", "$#,##0.00;($#,##0.00);-", False),
        (
            "Sum of all P&L lines",
            f"={PNL}!{total}{pnl_rows['Total Revenue']}+{PNL}!{total}{pnl_rows['Total Cost of Goods Sold']}"
            f"+{PNL}!{total}{pnl_rows['Total Operating Expenses']}",
            "$#,##0.00;($#,##0.00);-",
            False,
        ),
        ("Amount check", '=IF(AND(ABS(B8-B9)<0.005,ABS(B9-B10)<0.005),"OK","CHECK")', None, False),
    ]
    for row, (label, value, number_format, is_input) in enumerate(lines, start=4):
        ws.cell(row, 1, label).font = _font()
        cell = ws.cell(row, 2, value)
        cell.font = _font(color=BLUE_INPUT if is_input else "000000")
        cell.alignment = Alignment(horizontal="right")
        if number_format:
            cell.number_format = number_format
    ws.cell(12, 1, "Blue values are measured by Python when the report is built.").font = _font(italic=True, color="595959")
    ws.conditional_formatting.add("B4:B11", CellIsRule(operator="equal", formula=['"OK"'], font=Font(color="007A33", bold=True)))
    ws.conditional_formatting.add("B4:B11", CellIsRule(operator="equal", formula=['"CHECK"'], font=Font(color="C00000", bold=True)))

    ws.cell(14, 1, "Rejected rows").font = _font(bold=True, color=NAVY)
    _header(ws, 15, ["Reason", "Source row", "Date", "Description", "Category", "Amount"])
    for column in (3, 4, 5):
        ws.cell(15, column).alignment = Alignment(horizontal="left", indent=1)
    for row, item in enumerate(data.rejected.itertuples(index=False), start=16):
        values = [item.reason, int(item.source_row), item.date, item.description, item.category, item.amount]
        for column, value in enumerate(values, start=1):
            cell = ws.cell(row, column, value)
            cell.font = _font()
            if column == 6:
                cell.alignment = Alignment(horizontal="right")
            elif column in (3, 4, 5):
                cell.alignment = Alignment(indent=1)
    _widths(ws, {"A": 36, "B": 16, "C": 14, "D": 36, "E": 22, "F": 14})
    _page_setup(ws)


def _write_summary(ws: Worksheet, company: str, months: list, pnl: dict, bva: dict) -> None:
    first_month, last_month = months[0], months[-1]
    _title(ws, f"{company}: Financial Summary", f"{first_month:%b %Y} to {last_month:%b %Y}. All amounts in USD.")
    rows = pnl["rows"]
    total = col(pnl["total"])
    budget_rows = bva["rows"]

    _header(ws, 4, ["Key figures", "Actual", "Budget", "Variance"])
    figures = [
        ("Revenue", rows["Total Revenue"], budget_rows["Revenue"]),
        ("Cost of Goods Sold", rows["Total Cost of Goods Sold"], budget_rows["Cost of Goods Sold"]),
        ("Gross Profit", rows["Gross Profit"], None),
        ("Operating Expenses", rows["Total Operating Expenses"], budget_rows["Operating Expenses"]),
        ("Operating Profit", rows["Operating Profit"], budget_rows["Operating Profit"]),
    ]
    row = 5
    positions = {}
    for label, pnl_row, budget_row in figures:
        ws.cell(row, 1, label).font = _font()
        actual = ws.cell(row, 2, f"={PNL}!{total}{pnl_row}")
        actual.font = _font(color=GREEN_LINK)
        actual.number_format = MONEY
        if budget_row is not None:
            budget = ws.cell(row, 3, f"={BVA}!C{budget_row}")
            variance = ws.cell(row, 4, f"={BVA}!E{budget_row}")
        else:
            revenue_row, cost_row = budget_rows["Revenue"], budget_rows["Cost of Goods Sold"]
            budget = ws.cell(row, 3, f"={BVA}!C{revenue_row}-{BVA}!C{cost_row}")
            variance = ws.cell(row, 4, f"=B{row}-C{row}")
        budget.font = _font(color=GREEN_LINK)
        budget.number_format = MONEY
        variance.font = _font(color=GREEN_LINK)
        variance.number_format = MONEY
        positions[label] = row
        row += 1
    _emphasise(ws, positions["Operating Profit"], 4)

    ws.cell(row, 1, "Gross Margin").font = _font()
    ws.cell(row, 2, f"={PNL}!{total}{rows['Gross Margin']}")
    ws.cell(row + 1, 1, "Operating Margin").font = _font()
    ws.cell(row + 1, 2, f"={PNL}!{total}{rows['Operating Margin']}")
    for margin_row, source in ((row, "Gross Profit"), (row + 1, "Operating Profit")):
        cell = ws.cell(margin_row, 2)
        cell.font = _font(color=GREEN_LINK)
        cell.number_format = PERCENT
        budget = ws.cell(margin_row, 3, f"=IF(C{positions['Revenue']}=0,0,C{positions[source]}/C{positions['Revenue']})")
        budget.number_format = PERCENT
        budget.font = _font()
        variance = ws.cell(margin_row, 4, f"=B{margin_row}-C{margin_row}")
        variance.number_format = PERCENT
        variance.font = _font()
    row += 3

    header = pnl["header"]
    span = f"{col(pnl['first'])}{{0}}:{col(pnl['last'])}{{0}}"
    profit = f"{PNL}!${col(pnl['first'])}${rows['Operating Profit']}:${col(pnl['last'])}${rows['Operating Profit']}"
    labels = f"{PNL}!${col(pnl['first'])}${header}:${col(pnl['last'])}${header}"
    bva_first, bva_last = bva["first"], bva["last"]
    variances = f"{BVA}!$E${bva_first}:$E${bva_last}"
    names = f"{BVA}!$A${bva_first}:$A${bva_last}"

    _header(ws, row, ["Highlights", "Result", "Amount", ""])
    highlights = [
        ("Best month by operating profit", f"=INDEX({labels},1,MATCH(MAX({profit}),{profit},0))", f"=MAX({profit})", MONTH),
        ("Weakest month by operating profit", f"=INDEX({labels},1,MATCH(MIN({profit}),{profit},0))", f"=MIN({profit})", MONTH),
        ("Largest unfavourable variance", f"=INDEX({names},MATCH(MIN({variances}),{variances},0))", f"=MIN({variances})", None),
        ("Largest favourable variance", f"=INDEX({names},MATCH(MAX({variances}),{variances},0))", f"=MAX({variances})", None),
    ]
    for offset, (label, result, amount, number_format) in enumerate(highlights, start=1):
        target = row + offset
        ws.cell(target, 1, label).font = _font()
        result_cell = ws.cell(target, 2, result)
        result_cell.font = _font()
        result_cell.alignment = Alignment(horizontal="right")
        if number_format:
            result_cell.number_format = number_format
        amount_cell = ws.cell(target, 3, amount)
        amount_cell.font = _font()
        amount_cell.number_format = MONEY
    row += len(highlights) + 2

    ws.cell(row, 1, "Data checks").font = _font(bold=True)
    checks = ws.cell(row, 2, "=IF(AND('Data Checks'!B7=\"OK\",'Data Checks'!B11=\"OK\"),\"Passed\",\"Review Data Checks\")")
    checks.font = _font(bold=True)
    checks.alignment = Alignment(horizontal="right")
    ws.cell(row + 1, 1, "Rows rejected").font = _font()
    rejected = ws.cell(row + 1, 2, "='Data Checks'!B6")
    rejected.font = _font(color=GREEN_LINK)
    rejected.number_format = "#,##0;;0"
    ws.conditional_formatting.add(f"B{row}", FormulaRule(formula=[f'$B${row}="Passed"'], font=Font(color="007A33", bold=True)))
    ws.conditional_formatting.add(f"B{row}", FormulaRule(formula=[f'$B${row}<>"Passed"'], font=Font(color="C00000", bold=True)))

    pnl_sheet = ws.parent["Monthly P&L"]
    bva_sheet = ws.parent["Budget vs Actual"]
    months_ref = Reference(pnl_sheet, min_col=pnl["first"], max_col=pnl["last"], min_row=header)

    trend = BarChart()
    trend.type = "col"
    trend.title = "Revenue and operating profit by month (USD)"
    revenue_ref = Reference(pnl_sheet, min_col=1, max_col=pnl["last"], min_row=rows["Total Revenue"])
    trend.add_data(revenue_ref, from_rows=True, titles_from_data=True)
    trend.set_categories(months_ref)
    trend.y_axis.numFmt = "$#,##0"
    trend.y_axis.majorGridlines = None
    trend.x_axis.number_format = "mmm"
    trend.x_axis.delete = False
    trend.y_axis.delete = False
    trend.series[0].graphicalProperties.solidFill = NAVY

    line = LineChart()
    line_ref = Reference(pnl_sheet, min_col=1, max_col=pnl["last"], min_row=rows["Operating Profit"])
    line.add_data(line_ref, from_rows=True, titles_from_data=True)
    line.series[0].graphicalProperties.line.solidFill = "ED7D31"
    line.series[0].graphicalProperties.line.width = 28000
    line.series[0].smooth = False
    trend += line
    trend.legend.position = "b"
    trend.height, trend.width = 8.5, 19
    ws.add_chart(trend, "F4")

    costs = BarChart()
    costs.type = "bar"
    costs.title = "Operating expenses: budget vs actual (USD)"
    opex_rows = [
        bva_first + index
        for index in range(bva_last - bva_first + 1)
        if bva_sheet.cell(bva_first + index, 2).value == "Operating Expenses"
    ]
    costs.add_data(
        Reference(bva_sheet, min_col=3, max_col=4, min_row=opex_rows[0], max_row=opex_rows[-1]),
        titles_from_data=False,
    )
    for series, column in zip(costs.series, "CD"):
        series.tx = SeriesLabel(strRef=StrRef(f"{BVA}!${column}${bva_first - 1}"))
    costs.series[0].graphicalProperties.solidFill = "A6A6A6"
    costs.series[1].graphicalProperties.solidFill = NAVY
    costs.set_categories(Reference(bva_sheet, min_col=1, min_row=opex_rows[0], max_row=opex_rows[-1]))
    costs.y_axis.numFmt = "$#,##0"
    costs.y_axis.majorGridlines = None
    costs.x_axis.scaling.orientation = "maxMin"
    costs.x_axis.delete = False
    costs.y_axis.delete = False
    costs.legend.position = "b"
    costs.height, costs.width = 8.5, 19
    ws.add_chart(costs, "F22")

    _widths(ws, {"A": 34, "B": 15, "C": 14, "D": 14, "E": 3})
    _page_setup(ws)


def build_report(
    data: TransactionData,
    budget: pd.DataFrame,
    categories: pd.DataFrame,
    output: str | Path,
    company: str = "Sample Co.",
) -> Path:
    """Write the report and return its path."""
    if data.valid.empty:
        raise ValueError("No valid transactions to report on")

    period = pd.date_range(
        data.valid["date"].min().to_period("M").to_timestamp(),
        data.valid["date"].max().to_period("M").to_timestamp(),
        freq="MS",
    )
    months = list(period)
    budget = budget[budget["month"].isin(period)].reset_index(drop=True)

    workbook = Workbook()
    summary = workbook.active
    summary.title = "Summary"
    pnl_sheet = workbook.create_sheet("Monthly P&L")
    bva_sheet = workbook.create_sheet("Budget vs Actual")
    checks_sheet = workbook.create_sheet("Data Checks")
    transactions_sheet = workbook.create_sheet("Transactions")
    budget_sheet = workbook.create_sheet("Budget")

    _write_transactions(transactions_sheet, data.valid)
    _write_budget(budget_sheet, budget)
    pnl = _write_pnl(pnl_sheet, company, months, categories)
    bva = _write_bva(bva_sheet, company, categories, pnl)
    _write_checks(checks_sheet, company, data, pnl)
    _write_summary(summary, company, months, pnl, bva)

    workbook.calculation.fullCalcOnLoad = True
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    return output
