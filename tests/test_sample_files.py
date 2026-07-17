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
    def test_sheet1_order_amount_drives_ad_and_ah(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "sample-output.xlsx"
            original = load_workbook(SALES, data_only=True, read_only=True)
            original_order_amounts = [
                row[0]
                for row in original["Sheet1"].iter_rows(
                    min_row=2, min_col=14, max_col=14, values_only=True
                )
            ]
            sheet1_amounts = {
                str(row[0]): float(row[2])
                for row in original["Sheet1"].iter_rows(
                    min_row=2, min_col=12, max_col=14, values_only=True
                )
                if row[0] and row[2] is not None
            }
            original.close()
            stats = process_workbooks(RATIO, SALES, output)
            self.assertGreater(stats.matched_rows, 0)
            wb = load_workbook(output, data_only=True, read_only=True)
            result = wb["拆分结果"]
            output_order_amounts = [
                row[0]
                for row in wb["Sheet1"].iter_rows(
                    min_row=2, min_col=14, max_col=14, values_only=True
                )
            ]
            self.assertEqual(output_order_amounts, original_order_amounts)
            ad_totals = defaultdict(float)
            ah_totals = defaultdict(float)
            for row in result.iter_rows(min_row=2, values_only=True):
                self.assertAlmostEqual(float(row[29] or 0), float(row[33] or 0), places=8)
                if float(row[32] or 0) == 0:
                    self.assertEqual(float(row[26] or 0), 0)
                    self.assertEqual(float(row[29] or 0), 0)
                    self.assertEqual(float(row[33] or 0), 0)
                if row[10]:
                    order = str(row[10])
                    ad_totals[order] += float(row[29] or 0)
                    ah_totals[order] += float(row[33] or 0)
            self.assertEqual(set(ad_totals), set(ah_totals))
            for order in ad_totals:
                total = sheet1_amounts.get(order, 0.0)
                self.assertAlmostEqual(ad_totals[order], total, places=8)
                self.assertAlmostEqual(ah_totals[order], total, places=8)
            wb.close()


if __name__ == "__main__":
    unittest.main()
