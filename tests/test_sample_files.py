from __future__ import annotations

import tempfile
import unittest
import os
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

from splitter import process_workbooks


RATIO = Path(os.environ.get("BUNDLE_SPLITTER_RATIO_SAMPLE", "__missing_ratio_sample__.xlsx"))
SALES = Path(os.environ.get("BUNDLE_SPLITTER_SALES_SAMPLE", "__missing_sales_sample__.xlsx"))


@unittest.skipUnless(RATIO.exists() and SALES.exists(), "样例 Excel 不在当前电脑")
class RealSampleTests(unittest.TestCase):
    def test_every_web_order_balances_to_receivable_total(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "sample-output.xlsx"
            stats = process_workbooks(RATIO, SALES, output)
            self.assertGreater(stats.matched_rows, 0)
            wb = load_workbook(output, data_only=True, read_only=True)
            source = max(
                (
                    sheet
                    for sheet in wb.worksheets
                    if sheet.title != "拆分结果"
                    and sheet.cell(1, 11).value == "网店订单号"
                    and sheet.cell(1, 29).value == "金额"
                ),
                key=lambda sheet: sheet.max_row,
            )
            result = wb["拆分结果"]
            receivable_totals = {}
            split_totals = defaultdict(float)
            for row in source.iter_rows(min_row=2, values_only=True):
                if row[10] and row[13] is not None:
                    order = str(row[10])
                    total = float(row[13])
                    if order in receivable_totals:
                        self.assertAlmostEqual(receivable_totals[order], total, places=8)
                    else:
                        receivable_totals[order] = total
            for row in result.iter_rows(min_row=2, values_only=True):
                if row[10]:
                    split_totals[str(row[10])] += float(row[29] or 0)
            self.assertEqual(set(receivable_totals), set(split_totals))
            for order, total in receivable_totals.items():
                self.assertAlmostEqual(split_totals[order], total, places=8)
            wb.close()


if __name__ == "__main__":
    unittest.main()
