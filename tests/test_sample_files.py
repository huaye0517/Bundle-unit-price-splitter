from __future__ import annotations

import tempfile
import unittest
import os
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

from splitter import SALES_HEADERS, _find_sheet, _header_map, process_workbooks


RATIO = Path(os.environ.get("BUNDLE_SPLITTER_RATIO_SAMPLE", "__missing_ratio_sample__.xlsx"))
SALES = Path(os.environ.get("BUNDLE_SPLITTER_SALES_SAMPLE", "__missing_sales_sample__.xlsx"))


@unittest.skipUnless(RATIO.exists() and SALES.exists(), "样例 Excel 不在当前电脑")
class RealSampleTests(unittest.TestCase):
    def test_cached_order_total_drives_ad_and_aj(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "sample-output.xlsx"
            original = load_workbook(SALES, data_only=True, read_only=True)
            source, header_row = _find_sheet(original, SALES_HEADERS)
            self.assertIsNotNone(source)
            source_headers = _header_map(source, header_row)
            source_title = source.title
            receivable_col = source_headers["应收合计"]
            original_receivables = [
                values[0]
                for values in source.iter_rows(
                    min_row=header_row + 1,
                    min_col=receivable_col,
                    max_col=receivable_col,
                    values_only=True,
                )
            ]

            expected_totals = {}
            for sheet in original.worksheets:
                for candidate_header_row in range(1, min(sheet.max_row, 5) + 1):
                    headers = _header_map(sheet, candidate_header_row)
                    if not {"网店订单号", "订单金额"}.issubset(headers):
                        continue
                    for values in sheet.iter_rows(min_row=candidate_header_row + 1, values_only=True):
                        order = str(values[headers["网店订单号"] - 1] or "").strip()
                        amount = values[headers["订单金额"] - 1]
                        if order and amount is not None:
                            expected_totals[order] = float(amount)
                    break
                if expected_totals:
                    break

            if not expected_totals:
                web_order_col = source_headers["网店订单号"]
                order_col = source_headers["订单编号"]
                amount_col = source_headers["金额"]
                fee_bases = defaultdict(float)
                for values in source.iter_rows(min_row=header_row + 1, values_only=True):
                    order = str(
                        values[web_order_col - 1] or values[order_col - 1] or ""
                    ).strip()
                    amount = values[receivable_col - 1]
                    if not order:
                        continue
                    fee_bases[order] += float(values[amount_col - 1] or 0)
                    if amount is not None:
                        amount = float(amount)
                        if order in expected_totals:
                            self.assertAlmostEqual(expected_totals[order], amount, places=8)
                        else:
                            expected_totals[order] = amount
                for order in expected_totals:
                    expected_totals[order] -= fee_bases[order] * 0.01
            original.close()

            stats = process_workbooks(RATIO, SALES, output)
            self.assertGreater(stats.matched_rows, 0)
            self.assertEqual(stats.exceptional_orders, 0)
            wb = load_workbook(output, data_only=True, read_only=True)
            result = wb["拆分结果"]
            output_receivables = [
                values[0]
                for values in wb[source_title].iter_rows(
                    min_row=header_row + 1,
                    min_col=receivable_col,
                    max_col=receivable_col,
                    values_only=True,
                )
            ]
            self.assertEqual(output_receivables, original_receivables)

            ad_totals = defaultdict(float)
            ah_totals = defaultdict(float)
            for row in result.iter_rows(min_row=2, values_only=True):
                self.assertAlmostEqual(float(row[29] or 0), float(row[35] or 0), places=8)
                if float(row[32] or 0) == 0:
                    self.assertEqual(float(row[26] or 0), 0)
                    self.assertEqual(float(row[29] or 0), 0)
                    self.assertEqual(float(row[35] or 0), 0)
                order = str(row[10] or row[1] or "").strip()
                if order:
                    ad_totals[order] += float(row[29] or 0)
                    ah_totals[order] += float(row[35] or 0)
            self.assertEqual(set(ad_totals), set(ah_totals))
            for order in ad_totals:
                total = expected_totals.get(order, 0.0)
                self.assertAlmostEqual(ad_totals[order], total, places=8)
                self.assertAlmostEqual(ah_totals[order], total, places=8)
            wb.close()


if __name__ == "__main__":
    unittest.main()
