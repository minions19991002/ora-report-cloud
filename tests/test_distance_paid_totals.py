import sys
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "server" / "engine"))
import build_ora_report as report


def sample_sheet():
    book = Workbook()
    ws = book.active
    ws.title = "订单距离及实付区间"
    for title_row, header_row, first_row, total_row, title in (
        (3, 4, 5, 7, "订单距离分布"),
        (9, 10, 11, 13, "实付区间订单占比"),
    ):
        ws.cell(title_row, 2, title)
        ws.cell(header_row, 2, "门店")
        for offset, code in enumerate(("S1", "S2")):
            row = first_row + offset
            ws.cell(row, 1, f"门店{code}")
            ws.cell(row, 2, code)
            ws.cell(row, 13, code)
            ws.cell(row, 24, code)
        for col in (1, 13, 24):
            ws.cell(total_row, col, "总计")
    return book


class DistancePaidTotalsTest(unittest.TestCase):
    def setUp(self):
        self.stores = [report.Store(f"门店{code}", code, code, f"门店{code}", code) for code in ("S1", "S2")]

    def test_distance_uses_order_count_not_store_or_platform_average(self):
        rows = []
        for platform, code, count, close_share in (
            ("美团", "m1", 100, 0.2),
            ("饿了么", "e1", 300, 0.8),
            ("美团", "m2", 100, 0.5),
        ):
            rows.append({
                "平台": platform, "平台门店ID": code, "订单数": count,
                "[0,0.5Km)": close_share, "[0.5,0.8Km)": 0, "[0.8,1.0Km)": 0,
                "[1.0,2.0Km)": 1 - close_share, "[2.0,3.0Km)": 0,
                "[3.0,4.0Km)": 0, "[4.0,5.0Km)": 0, ">5.0Km": 0,
            })
        with patch.object(report, "read_excel", return_value=pd.DataFrame(rows)):
            distance = report.compute_distance({"m1": "S1", "m2": "S2"}, {"e1": "S1"})
        self.assertAlmostEqual(distance["S1"]["total"]["0_1"], 0.65)
        self.assertEqual(distance["S1"]["total"]["total_count"], 400)
        self.assertAlmostEqual(distance["S2"]["total"]["0_1"], 0.5)

        current = sample_sheet()
        previous = sample_sheet()
        prev = previous["订单距离及实付区间"]
        for row in (7, 13):
            prev.cell(row, 3, 0.6)
            prev.cell(row, 14, 0.4)
            prev.cell(row, 25, 0.8)
        paid = {
            "S1": {
                "total": {"total_count": 400, "0_20_count": 260, "20_30_count": 140, "gt30_count": 0, "0_20": 0.65, "20_30": 0.35, "gt30": 0},
                "mt": {"total_count": 100, "0_20_count": 20, "20_30_count": 80, "gt30_count": 0, "0_20": 0.2, "20_30": 0.8, "gt30": 0},
                "ele": {"total_count": 300, "0_20_count": 240, "20_30_count": 60, "gt30_count": 0, "0_20": 0.8, "20_30": 0.2, "gt30": 0},
            },
            "S2": {
                "total": {"total_count": 100, "0_20_count": 50, "20_30_count": 50, "gt30_count": 0, "0_20": 0.5, "20_30": 0.5, "gt30": 0},
                "mt": {"total_count": 100, "0_20_count": 50, "20_30_count": 50, "gt30_count": 0, "0_20": 0.5, "20_30": 0.5, "gt30": 0},
                "ele": {"total_count": 0, "0_20_count": 0, "20_30_count": 0, "gt30_count": 0, "0_20": 0, "20_30": 0, "gt30": 0},
            },
        }
        report.write_distance_and_paid(current, previous, distance, paid, self.stores)
        ws = current["订单距离及实付区间"]
        for row in (7, 13):
            self.assertAlmostEqual(ws.cell(row, 3).value, 0.62)
            self.assertAlmostEqual(ws.cell(row, 5).value, 0.02)
            self.assertAlmostEqual(ws.cell(row, 14).value, 0.35)
            self.assertAlmostEqual(ws.cell(row, 25).value, 0.8)
            self.assertEqual(ws.cell(row, 3).number_format, "0.0%")

        for row in (7, 13):
            prev.cell(row, 1).value = None
        current_without_prior_total = sample_sheet()
        report.write_distance_and_paid(current_without_prior_total, previous, distance, paid, self.stores)
        without_prior = current_without_prior_total["订单距离及实付区间"]
        for row in (7, 13):
            self.assertAlmostEqual(without_prior.cell(row, 3).value, 0.62)
            self.assertIsNone(without_prior.cell(row, 4).value)
            self.assertIsNone(without_prior.cell(row, 5).value)


if __name__ == "__main__":
    unittest.main()
