from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from splitter import SplitterError, process_workbooks


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


class SplitterTests(unittest.TestCase):
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
            self.assertEqual(ws["AH3"].number_format, "0.00")
            self.assertAlmostEqual(ws["AB3"].value, 40 / 100 / 1, places=14)
            self.assertAlmostEqual(ws["AC3"].value, 80, places=12)
            self.assertAlmostEqual(ws["AD3"].value, 80, places=12)
            self.assertAlmostEqual(ws["AA4"].value, 60, places=12)
            self.assertAlmostEqual(ws["AC4"].value, 60, places=12)
            self.assertAlmostEqual(ws["AD4"].value, 120, places=12)
            self.assertAlmostEqual(ws["AH3"].value + ws["AH4"].value, 200, places=12)
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
                self.assertEqual(result.cell(row, 29).value, 50)
            self.assertEqual(result.cell(2, 34).value, 50)
            self.assertEqual(result.cell(3, 34).value, 50)
            self.assertEqual(result.cell(4, 34).value, 100)
            self.assertEqual(sum(result.cell(row, 34).value for row in range(2, 5)), 200)

    def test_web_orders_are_balanced_independently(self):
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
            ws["N2"] = 50
            ws["N3"] = None
            ws["N4"] = 80
            wb.save(sales)
            wb.close()

            stats = process_workbooks(ratio, sales, output)
            self.assertEqual(stats.orders, 2)
            result = load_workbook(output, data_only=True)["拆分结果"]
            self.assertAlmostEqual(result["AD2"].value + result["AD3"].value, 50, places=12)
            self.assertAlmostEqual(result["AD4"].value, 80, places=12)

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
            self.assertAlmostEqual(result["AD5"].value, 30, places=12)

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
