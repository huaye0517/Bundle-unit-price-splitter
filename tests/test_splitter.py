from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from splitter import (
    SplitterError,
    _enable_formula_recalculation,
    _freeze_cached_amounts,
    process_workbooks,
    ratio_file_info,
    update_ratio_data,
)


SALES_HEADERS = [
    "标记", "订单编号", "订单状态", "结算状态", "销售渠道", "处理时间", "付款时间", "发货仓库",
    "物流公司", "物流单号", "网店订单号", "发货时间", "订单类型", "应收合计", "货品数量", "货品摘要",
    "客户账号", "收货人", "手机", "收货地址", "合并备注", "货品编号", "货品名称", "规格", "数量", "单价",
    "优惠", "折扣", "金额", "锁定待发", "备注", "达人ID", "达人名称", "赠品",
]


def make_ratio(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet2"
    ws.append(["组合装编号", "单品编号", "金额", "分摊比例", "单价", "数量"])
    ws.append(["A", "SKU1", 100, 0.4, 40, 1])
    ws.append(["B", "SKU1", 200, 0.5, 100, 1])  # must be ignored: VLOOKUP first match
    ws.append(["A", "SKU2", 100, 0.6, 60, 2])
    wb.save(path)


def make_new_ratio(path: Path) -> None:
    wb = Workbook()
    wb.active.title = "sheet"
    ws = wb.create_sheet("sheet1")
    headers = ["母件编号", "编号", "名称", "规格", "条码", "长(cm)", "宽(cm)", "高(cm)", "重量(g)", "体积(cm³)", "单位", "数量", "执行价格", "分摊金额", "分摊比例", "上传SN", "赠品", "换算率"]
    ws.append(headers)
    ws.append(["A", "SKU1", "商品1", "", "", "", "", "", "", "", "件", 2, 40, 80, 0.4, "否", "否", ""])
    ws.append(headers)  # merged exports can contain repeated header rows
    ws.append(["A", "SKU2", "商品2", "", "", "", "", "", "", "", "件", 2, 60, 120, 0.6, "否", "否", ""])
    wb.save(path)


def make_sales(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "销售明细"
    ws.append(SALES_HEADERS)
    base = [""] * len(SALES_HEADERS)
    for index, (code, qty) in enumerate([("MISS", 1), ("SKU1", 1), ("SKU2", 2)]):
        row = base.copy()
        row[1] = "ORDER-1"
        row[10] = "WEB-1"
        row[13] = 200 if index == 0 else None
        row[21] = code
        row[24] = qty
        row[25] = 10
        row[28] = qty * 10
        ws.append(row)
    wb.create_sheet("其他工作表")["A1"] = "保留"
    wb.save(path)


def add_order_amounts(path: Path, amounts: list[tuple[str, float]]) -> None:
    wb = load_workbook(path)
    if "Sheet1" in wb.sheetnames:
        del wb["Sheet1"]
    ws = wb.create_sheet("Sheet1")
    headers = [""] * 14
    headers[11] = "网店订单号"
    headers[13] = "订单金额"
    ws.append(headers)
    for web_order, amount in amounts:
        row = [""] * 14
        row[11] = web_order
        row[13] = amount
        ws.append(row)
    wb.save(path)
    wb.close()


class SplitterTests(unittest.TestCase):
    def test_missing_calculation_properties_are_initialized(self):
        workbook = Workbook()
        workbook.calculation = None

        _enable_formula_recalculation(workbook)

        self.assertEqual(workbook.calculation.calcMode, "auto")
        self.assertTrue(workbook.calculation.fullCalcOnLoad)
        self.assertTrue(workbook.calculation.forceFullCalc)
        workbook.close()

    def test_cached_receivable_formulas_are_frozen_as_values(self):
        formula_book = Workbook()
        formula_sheet = formula_book.active
        formula_sheet.title = "销售明细"
        formula_sheet.append(SALES_HEADERS)
        formula_sheet.append([None] * 10 + ["WEB-1", None, None, "=SUMIF(...)" ] + [None] * 20)
        formula_sheet.append([None] * 10 + ["WEB-1", None, None, "=SUMIF(...)" ] + [None] * 20)

        cached_book = Workbook()
        cached_sheet = cached_book.active
        cached_sheet.title = "销售明细"
        cached_sheet.append(SALES_HEADERS)
        cached_sheet.append([None] * 10 + ["WEB-1", None, None, 187.5] + [None] * 20)
        cached_sheet.append([None] * 10 + ["WEB-1", None, None, None] + [None] * 20)

        frozen = _freeze_cached_amounts(formula_book, cached_book)

        self.assertEqual(frozen, 1)
        self.assertEqual(formula_sheet["N2"].value, 187.5)
        self.assertIsNone(formula_sheet["N3"].value)
        formula_book.close()
        cached_book.close()

    def test_repeated_receivable_total_is_counted_once_and_fee_is_deducted(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ratio, sales, output = root / "ratio.xlsx", root / "sales.xlsx", root / "output.xlsx"
            make_ratio(ratio)
            wb = Workbook()
            ws = wb.active
            ws.title = "销售明细"
            ws.append(SALES_HEADERS)
            for index in range(10):
                row = [""] * len(SALES_HEADERS)
                row[1] = f"ORDER-{index + 1}"
                row[10] = "PO-260629-323800286823316"
                row[13] = 187.5
                row[21] = "SKU1"
                row[24] = 1
                row[25] = 18.75
                row[28] = 18.75
                ws.append(row)
            wb.save(sales)
            wb.close()

            stats = process_workbooks(ratio, sales, output)

            self.assertEqual(stats.orders, 1)
            self.assertEqual(stats.exceptional_orders, 0)
            result_book = load_workbook(output, data_only=False)
            result = result_book["拆分结果"]
            self.assertEqual([result.cell(row, 14).value for row in range(2, 12)], [187.5] * 10)
            net_total = 187.5 - 187.5 * 0.01
            self.assertAlmostEqual(sum(result.cell(row, 30).value for row in range(2, 12)), net_total, places=8)
            self.assertAlmostEqual(sum(result.cell(row, 36).value for row in range(2, 12)), net_total, places=8)
            for row in range(2, 12):
                self.assertAlmostEqual(result.cell(row, 30).value, result.cell(row, 36).value, places=8)
                self.assertEqual(result.cell(row, 34).value, f"=AG{row}*1%")
                self.assertEqual(result.cell(row, 35).value, f"=N{row}-AH{row}")
            self.assertEqual(result["AH1"].value, "手续费1%")
            self.assertEqual(result["AI1"].value, "最终金额")
            self.assertEqual(result["AJ1"].value, "摊后金额")
            self.assertEqual(result["AO1"].value, "赠品")
            self.assertEqual(result.auto_filter.ref, "A1:AO11")
            summary = result_book["透视表"]
            self.assertEqual(
                [summary.cell(1, col).value for col in range(1, 5)],
                ["网店订单号", "平均值项:应收合计", "求和项:金额", "差异"],
            )
            self.assertEqual(summary["A2"].value, "PO-260629-323800286823316")
            self.assertEqual(summary["B2"].value, net_total)
            self.assertAlmostEqual(summary["C2"].value, net_total, places=8)
            self.assertEqual(summary["D2"].value, "=C2-B2")
            self.assertEqual(summary["A3"].value, "总计")
            self.assertEqual(summary["B3"].value, "=SUM(B2:B2)")
            self.assertEqual(summary["C3"].value, "=SUM(C2:C2)")
            self.assertEqual(summary["D3"].value, "=C3-B3")
            self.assertEqual(len(getattr(summary, "_pivots", [])), 0)
            self.assertEqual(result_book.calculation.calcMode, "auto")
            self.assertTrue(result_book.calculation.fullCalcOnLoad)
            self.assertTrue(result_book.calculation.forceFullCalc)
            result_book.close()

    def test_summary_uses_fee_net_order_amounts(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ratio, sales, output = root / "ratio.xlsx", root / "sales.xlsx", root / "output.xlsx"
            make_ratio(ratio)
            wb = Workbook()
            ws = wb.active
            ws.title = "销售明细"
            ws.append(SALES_HEADERS)
            for order, gross, original_amount in (
                ("260531-420353175082566", 24.0, 14.75),
                ("260531-420353175082566", 24.0, 14.75),
                ("260531-421947010581809", 73.5, 43.065),
                ("260531-421947010581809", 73.5, 43.065),
            ):
                row = [""] * len(SALES_HEADERS)
                row[1] = order
                row[10] = order
                row[13] = gross
                row[21] = "SKU1"
                row[24] = 1
                row[25] = original_amount
                row[28] = original_amount
                ws.append(row)
            wb.save(sales)
            wb.close()

            process_workbooks(ratio, sales, output)

            result = load_workbook(output, data_only=False)
            summary = result["透视表"]
            self.assertEqual(summary["B2"].value, 23.705)
            self.assertAlmostEqual(summary["C2"].value, 23.705, places=8)
            self.assertAlmostEqual(summary["B3"].value, 72.6387, places=8)
            self.assertAlmostEqual(summary["C3"].value, 72.6387, places=8)
            self.assertEqual(summary["D2"].value, "=C2-B2")
            self.assertEqual(summary["D3"].value, "=C3-B3")
            result.close()

    def test_conflicting_receivable_totals_mark_only_that_order_exceptional(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ratio, sales, output = root / "ratio.xlsx", root / "sales.xlsx", root / "output.xlsx"
            make_ratio(ratio)
            make_sales(sales)
            wb = load_workbook(sales)
            ws = wb["销售明细"]
            ws["N2"] = 187.5
            ws["N3"] = 188
            ws["N4"] = 187.5
            wb.save(sales)
            wb.close()

            stats = process_workbooks(ratio, sales, output)

            self.assertEqual(stats.exceptional_orders, 1)
            self.assertTrue(any("多个不同的应收合计" in warning for warning in stats.warnings))
            result = load_workbook(output, data_only=True)["拆分结果"]
            for row in range(2, 5):
                self.assertEqual(result.cell(row, 30).value, 0)
                self.assertEqual(result.cell(row, 36).value, 0)

    def test_blank_web_order_falls_back_to_internal_order_number(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ratio, sales, output = root / "ratio.xlsx", root / "sales.xlsx", root / "output.xlsx"
            make_ratio(ratio)
            make_sales(sales)
            wb = load_workbook(sales)
            ws = wb["销售明细"]
            for row in range(2, 5):
                ws.cell(row, 11, None)
                ws.cell(row, 14, 187.5)
            wb.save(sales)
            wb.close()

            stats = process_workbooks(ratio, sales, output)

            self.assertEqual(stats.orders, 1)
            self.assertEqual(stats.exceptional_orders, 0)
            result = load_workbook(output, data_only=True)["拆分结果"]
            self.assertAlmostEqual(sum(result.cell(row, 30).value for row in range(2, 5)), 187.1, places=8)
            self.assertAlmostEqual(sum(result.cell(row, 36).value for row in range(2, 5)), 187.1, places=8)

    def test_append_ratio_data_adds_new_bundle_items_without_overwriting(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            current, incoming = root / "current.xlsx", root / "incoming.xlsx"
            make_new_ratio(current)
            wb = Workbook()
            ws = wb.active
            ws.title = "总表"
            headers = ["母件编号", "编号", "名称", "规格", "条码", "长(cm)", "宽(cm)", "高(cm)", "重量(g)", "体积(cm³)", "单位", "数量", "执行价格", "分摊金额", "分摊比例", "上传SN", "赠品", "换算率"]
            ws.append(headers)
            ws.append(["A", "SKU1", "重复商品", "", "", "", "", "", "", "", "件", 1, 999, 999, 1, "否", "否", ""])
            ws.append(["B", "SKU3", "新增商品", "", "", "", "", "", "", "", "件", 1, 30, 30, 1, "否", "否", ""])
            wb.save(incoming)
            wb.close()

            stats = update_ratio_data(current, incoming, "append")

            self.assertEqual(stats.added_rows, 1)
            self.assertEqual(stats.source_rows, 3)
            self.assertEqual(stats.unique_items, 3)
            unique_items, source_rows = ratio_file_info(current)
            self.assertEqual((unique_items, source_rows), (3, 3))
            result = load_workbook(current, data_only=True)
            data_sheet = max(result.worksheets, key=lambda sheet: sheet.max_row)
            rows = list(data_sheet.iter_rows(min_row=2, values_only=True))
            self.assertEqual(sum(1 for row in rows if row[0] == "A" and row[1] == "SKU1"), 1)
            self.assertTrue(any(row[0] == "B" and row[1] == "SKU3" and row[12] == 30 for row in rows))
            result.close()

    def test_replace_ratio_data_uses_incoming_workbook(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            current, incoming = root / "current.xlsx", root / "incoming.xlsx"
            make_new_ratio(current)
            make_ratio(incoming)

            stats = update_ratio_data(current, incoming, "replace")

            self.assertEqual(stats.added_rows, 3)
            self.assertEqual(stats.source_rows, 3)
            self.assertEqual(stats.unique_items, 2)
            result = load_workbook(current, data_only=True)
            self.assertIn("Sheet2", result.sheetnames)
            self.assertNotIn("sheet1", result.sheetnames)
            result.close()

    def test_supports_new_ratio_export_format(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ratio, sales, output = root / "ratio.xlsx", root / "sales.xlsx", root / "output.xlsx"
            make_new_ratio(ratio)
            make_sales(sales)

            stats = process_workbooks(ratio, sales, output)

            self.assertEqual(stats.matched_rows, 2)
            result = load_workbook(output, data_only=True)["拆分结果"]
            self.assertEqual(result["AA3"].value, 40)
            self.assertEqual(result["AA4"].value, 60)
            self.assertAlmostEqual(result["AJ3"].value + result["AJ4"].value, 199.6, places=12)

    def test_formula_chain_and_first_match(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ratio, sales, output = root / "ratio.xlsx", root / "sales.xlsx", root / "output.xlsx"
            make_ratio(ratio)
            make_sales(sales)
            progress_events = []
            stats = process_workbooks(ratio, sales, output, progress=lambda value, message: progress_events.append((value, message)))

            self.assertEqual(stats.orders, 1)
            self.assertEqual(stats.matched_rows, 2)
            self.assertEqual(stats.unmatched_rows, 1)
            self.assertEqual(progress_events[-1][0], 100)
            self.assertTrue(any("正在拆分订单 1/1" in message for _, message in progress_events))
            wb = load_workbook(output, data_only=True)
            ws = wb["拆分结果"]
            self.assertEqual(ws["AA2"].value, 0)
            self.assertEqual(ws["AD2"].value, 0)
            self.assertEqual(ws["AA3"].value, 40)  # first SKU1, not 100
            self.assertEqual(ws["AA3"].number_format, "0.00")
            self.assertEqual(ws["AB3"].number_format, "0.00")
            self.assertEqual(ws["AJ3"].number_format, "0.00")
            self.assertAlmostEqual(ws["AB3"].value, 40 / 100 / 1, places=14)
            self.assertAlmostEqual(ws["AC3"].value, 79.84, places=12)
            self.assertAlmostEqual(ws["AD3"].value, 79.84, places=12)
            self.assertAlmostEqual(ws["AA4"].value, 60, places=12)
            self.assertAlmostEqual(ws["AC4"].value, 59.88, places=12)
            self.assertAlmostEqual(ws["AD4"].value, 119.76, places=12)
            self.assertAlmostEqual(ws["AD3"].value + ws["AD4"].value, 199.6, places=12)
            self.assertAlmostEqual(ws["AJ3"].value + ws["AJ4"].value, 199.6, places=12)
            self.assertIn("销售明细", wb.sheetnames)
            self.assertIn("其他工作表", wb.sheetnames)
            self.assertEqual(wb["其他工作表"]["A1"].value, "保留")
            wb.close()

    def test_all_unmatched_order_allocates_receivable_by_original_amount(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ratio, sales, output = root / "ratio.xlsx", root / "sales.xlsx", root / "output.xlsx"
            make_ratio(ratio)
            make_sales(sales)
            wb = load_workbook(sales)
            ws = wb["销售明细"]
            ws["V3"] = "MISS2"
            ws["V4"] = "MISS3"
            wb.save(sales)
            wb.close()
            stats = process_workbooks(ratio, sales, output)
            self.assertEqual(stats.exceptional_orders, 0)
            self.assertFalse(stats.warnings)
            result = load_workbook(output, data_only=True)["拆分结果"]
            for row in range(2, 5):
                self.assertEqual(result.cell(row, 27).value, 0)
                self.assertEqual(result.cell(row, 29).value, 49.9)
            self.assertEqual(result.cell(2, 30).value, 49.9)
            self.assertEqual(result.cell(3, 30).value, 49.9)
            self.assertEqual(result.cell(4, 30).value, 99.8)
            self.assertEqual(result.cell(2, 36).value, 49.9)
            self.assertEqual(result.cell(3, 36).value, 49.9)
            self.assertEqual(result.cell(4, 36).value, 99.8)
            self.assertEqual(sum(result.cell(row, 36).value for row in range(2, 5)), 199.6)

    def test_internal_orders_are_balanced_independently(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ratio, sales, output = root / "ratio.xlsx", root / "sales.xlsx", root / "output.xlsx"
            make_ratio(ratio)
            make_sales(sales)
            wb = load_workbook(sales)
            ws = wb["销售明细"]
            ws["K2"] = "WEB-A"
            ws["K3"] = "WEB-A"
            ws["K4"] = "WEB-B"
            ws["B4"] = "ORDER-2"
            ws["N2"] = 50
            ws["N3"] = None
            ws["N4"] = 80
            wb.save(sales)
            wb.close()

            stats = process_workbooks(ratio, sales, output)
            self.assertEqual(stats.orders, 2)
            result = load_workbook(output, data_only=True)["拆分结果"]
            self.assertAlmostEqual(result["AD2"].value + result["AD3"].value, 49.8, places=12)
            self.assertAlmostEqual(result["AD4"].value, 79.8, places=12)
            self.assertAlmostEqual(result["AJ2"].value + result["AJ3"].value, 49.8, places=12)
            self.assertAlmostEqual(result["AJ4"].value, 79.8, places=12)

    def test_sheet1_order_amounts_are_allocated_by_web_order(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ratio, sales, output = root / "ratio.xlsx", root / "sales.xlsx", root / "output.xlsx"
            make_ratio(ratio)
            make_sales(sales)
            wb = load_workbook(sales)
            ws = wb["销售明细"]
            ws["K2"] = "WEB-A"
            ws["K3"] = "WEB-A"
            ws["K4"] = "WEB-B"
            ws["N2"] = 130
            ws["N3"] = 130
            ws["N4"] = 130
            wb.save(sales)
            wb.close()
            add_order_amounts(sales, [("WEB-A", 50), ("WEB-B", 80)])

            stats = process_workbooks(ratio, sales, output)

            self.assertEqual(stats.orders, 2)
            self.assertEqual(stats.exceptional_orders, 0)
            result = load_workbook(output, data_only=True)["拆分结果"]
            self.assertAlmostEqual(sum(result.cell(row, 30).value for row in range(2, 5)), 130, places=12)
            self.assertAlmostEqual(sum(result.cell(row, 36).value for row in range(2, 5)), 130, places=12)

    def test_zero_original_amount_does_not_participate_in_allocation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ratio, sales, output = root / "ratio.xlsx", root / "sales.xlsx", root / "output.xlsx"
            make_ratio(ratio)
            make_sales(sales)
            wb = load_workbook(sales)
            ws = wb["销售明细"]
            ws["AC3"] = 0
            wb.save(sales)
            wb.close()

            process_workbooks(ratio, sales, output)

            result = load_workbook(output, data_only=True)["拆分结果"]
            self.assertEqual(result["AA3"].value, 0)
            self.assertEqual(result["AD3"].value, 0)
            self.assertEqual(result["AJ3"].value, 0)
            self.assertAlmostEqual(sum(result.cell(row, 30).value for row in range(2, 5)), 199.7, places=12)
            self.assertAlmostEqual(sum(result.cell(row, 36).value for row in range(2, 5)), 199.7, places=12)

    def test_sheet1_web_order_amount_overrides_repeated_source_totals(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ratio, sales, output = root / "ratio.xlsx", root / "sales.xlsx", root / "output.xlsx"
            make_ratio(ratio)
            make_sales(sales)
            wb = load_workbook(sales)
            ws = wb["销售明细"]
            ws["K2"] = "WEB-SHARED"
            ws["K3"] = "WEB-SHARED"
            ws["K4"] = "WEB-SHARED"
            ws["B4"] = "ORDER-2"
            ws["N2"] = 50
            ws["N3"] = None
            ws["N4"] = 80
            wb.save(sales)
            wb.close()
            add_order_amounts(sales, [("WEB-SHARED", 130)])

            stats = process_workbooks(ratio, sales, output)

            self.assertEqual(stats.orders, 1)
            self.assertEqual(stats.exceptional_orders, 0)
            result = load_workbook(output, data_only=True)["拆分结果"]
            self.assertAlmostEqual(sum(result.cell(row, 30).value for row in range(2, 5)), 130, places=12)
            self.assertAlmostEqual(sum(result.cell(row, 36).value for row in range(2, 5)), 130, places=12)

    def test_uses_sales_sheet_with_most_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ratio, sales, output = root / "ratio.xlsx", root / "sales.xlsx", root / "output.xlsx"
            make_ratio(ratio)
            make_sales(sales)
            wb = load_workbook(sales)
            source = wb["销售明细"]
            corrected = wb.copy_worksheet(source)
            corrected.title = "销售明细修正版"
            row = [""] * len(SALES_HEADERS)
            row[1] = "ORDER-2"
            row[10] = "WEB-2"
            row[13] = 30
            row[21] = "SKU1"
            row[24] = 1
            row[25] = 30
            row[28] = 30
            corrected.append(row)
            wb.save(sales)
            wb.close()

            stats = process_workbooks(ratio, sales, output)
            self.assertEqual(stats.orders, 2)
            result = load_workbook(output, data_only=True)["拆分结果"]
            self.assertEqual(result.max_row, 5)
            self.assertEqual(result["K5"].value, "WEB-2")
            self.assertAlmostEqual(result["AD5"].value, 29.7, places=12)

    def test_source_files_are_not_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ratio, sales = root / "ratio.xlsx", root / "sales.xlsx"
            make_ratio(ratio)
            make_sales(sales)
            with self.assertRaises(SplitterError):
                process_workbooks(ratio, sales, sales)


if __name__ == "__main__":
    unittest.main()
