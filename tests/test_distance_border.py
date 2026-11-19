import sys
import unittest
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Border, Side

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "server" / "engine"))
import build_ora_report as report


class DistanceBorderTest(unittest.TestCase):
    def test_closing_border_moves_to_last_added_store(self):
        book = Workbook()
        ws = book.active
        ws["A2"] = "实付区间订单占比"
        ws["B3"] = "店号"
        ws["A3"] = "门店"
        stores = [report.Store(f"门店{i}", str(i), str(i), f"门店{i}", f"S{i:02}") for i in range(18)]
        for row, store in enumerate(stores[:16], start=4):
            ws.cell(row, 1, store.name)
            ws.cell(row, 2, store.code)
            ws.cell(row, 3).border = Border(left=Side(style="medium"))
        ws.cell(19, 3).border = Border(left=Side(style="medium"), bottom=Side(style="thin"))

        report.ensure_titled_store_rows(ws, ["实付区间订单占比"], stores, 4, 19, 2, 1)

        self.assertIsNone(ws.cell(19, 3).border.bottom)
        self.assertEqual(ws.cell(21, 3).border.bottom.style, "thin")
        self.assertEqual(ws.cell(21, 3).border.left.style, "medium")
        self.assertEqual(ws.cell(21, 2).value, "S17")


if __name__ == "__main__":
    unittest.main()
