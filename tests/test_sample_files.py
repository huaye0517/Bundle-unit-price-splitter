from __future__ import annotations

import tempfile
import unittest
import os
from pathlib import Path

from openpyxl import load_workbook

from splitter import process_workbooks


RATIO = Path(os.environ.get("BUNDLE_SPLITTER_RATIO_SAMPLE", "__missing_ratio_sample__.xlsx"))
SALES = Path(os.environ.get("BUNDLE_SPLITTER_SALES_SAMPLE", "__missing_sales_sample__.xlsx"))


@unittest.skipUnless(RATIO.exists() and SALES.exists(), "样例 Excel 不在当前电脑")
class RealSampleTests(unittest.TestCase):
    def test_reference_order(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "sample-output.xlsx"
            stats = process_workbooks(RATIO, SALES, output)
            self.assertGreater(stats.matched_rows, 0)
            wb = load_workbook(output, data_only=True, read_only=True)
            ws = wb["拆分结果"]
            expected = {
                85: (0, 0, 0, 0, 0),
                86: (13.9986, 0.196705096194477, 25.0208882359375, 25.0208882359375, 25.0208882359375),
                87: (40, 0.281035383816206, 35.7477008214215, 71.4954016428429, 71.4954016428429),
                88: (17.166816, 0.0603060340432774, 7.67092753030489, 30.6837101212196, 30.6837101212196),
            }
            columns = (27, 28, 29, 30, 34)
            for row, values in expected.items():
                for col, value in zip(columns, values):
                    self.assertAlmostEqual(ws.cell(row, col).value, value, places=11)
            self.assertAlmostEqual(sum(ws.cell(row, 34).value for row in range(85, 89)), 127.2, places=11)
            wb.close()


if __name__ == "__main__":
    unittest.main()
