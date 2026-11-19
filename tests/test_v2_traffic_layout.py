import sys
import unittest
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "server" / "engine"))
import build_ora_report as report


def platform(exposure, entry, order):
    return {
        "sales_daily": 100.0,
        "at": 20.0,
        "discount_rate": 0.3,
        "exp_people_daily": exposure,
        "entry_rate": entry,
        "order_rate": order,
        "ad_spend": 10.0,
        "ad_roi": 2.0,
        "score": 4.8,
        "good": 10,
        "bad": 1,
    }


def v2_book(header):
    book = Workbook()
    ws = book.active
    ws.title = "V2"
    for title, title_row, header_row, store_row, total_row in (
        ("业绩", 1, 2, 3, 19),
        ("客单折扣", 21, 22, 23, 39),
        ("流量数据", 41, 42, 43, 59),
        ("推广数据", 62, 63, 64, 80),
    ):
        ws.cell(title_row, 1, title)
        ws.cell(header_row, 1, "店号")
        ws.cell(header_row, 2, "店名")
        ws.cell(store_row, 1, "S01")
        ws.cell(store_row, 2, "测试店")
        ws.cell(total_row, 1, "总计")
    ws["C42"] = header
    return book


class V2TrafficLayoutTest(unittest.TestCase):
    def setUp(self):
        self.previous = Workbook()
        sheet = self.previous.active
        sheet.title = "上期"
        sheet["A21"] = "流量数据"
        sheet["B22"] = "店号"
        sheet["C22"] = "店名"
        sheet["B23"] = "S01"
        sheet["C23"] = "测试店"
        sheet["C24"] = "总计"
        for row in (23, 24):
            for col, value in ((14, 80), (17, 0.08), (19, 0.25), (24, 40), (27, 0.12), (29, 0.30)):
                sheet.cell(row, col, value)
        self.store = report.Store("测试店", "1", "2", "测试店", "S01")
        self.data = {
            "total": platform(150, 0.11, 0.29),
            "mt": platform(100, 0.10, 0.30),
            "ele": platform(50, 0.15, 0.35),
            "mini": platform(0, 0, 0),
        }

    def test_new_layout_writes_both_platforms_and_comparisons(self):
        book = v2_book("美团曝光人数_日均")
        report.write_v2_compact(book, self.previous, [self.store], {"S01": self.data}, self.data)
        ws = book["V2"]
        expected = (100, 0.25, 0.10, 0.02, 0.30, 0.05, 50, 0.25, 0.15, 0.03, 0.35, 0.05)
        for row in (43, 59):
            for col, value in enumerate(expected, start=3):
                self.assertAlmostEqual(ws.cell(row, col).value, value, places=8)
        self.assertEqual(ws["D43"].number_format, "0.0%")
        self.assertEqual(ws["J43"].number_format, "0.0%")

    def test_shifted_traffic_section_still_uses_both_platforms(self):
        book = v2_book("美团曝光人数_日均")
        book["V2"].insert_rows(41, 18)
        report.write_v2_compact(book, self.previous, [self.store], {"S01": self.data}, self.data)
        ws = book["V2"]
        self.assertEqual(ws["C60"].value, "美团曝光人数_日均")
        self.assertEqual(ws["C61"].value, 100)
        self.assertEqual(ws["I61"].value, 50)
        self.assertEqual(ws["C77"].value, 100)
        self.assertEqual(ws["I77"].value, 50)

    def test_legacy_v2_keeps_combined_six_columns(self):
        book = v2_book("曝光人数_日均")
        report.write_v2_compact(book, self.previous, [self.store], {"S01": self.data}, self.data)
        ws = book["V2"]
        self.assertEqual(ws["C43"].value, 150)
        self.assertIsNone(ws["I43"].value)


if __name__ == "__main__":
    unittest.main()
