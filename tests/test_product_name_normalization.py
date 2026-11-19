import sys
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "server" / "engine"))
import build_ora_report as report


class ProductNameNormalizationTest(unittest.TestCase):
    def test_exact_alias_and_whitespace(self):
        self.assertEqual(
            report.canonical_package("【当季新品】 双杯 尝鲜套餐 1"),
            "【当季新品】双杯尝鲜套餐",
        )
        self.assertEqual(report.clean_product_name("  冰\u3000美 式\t"), "冰美式")
        self.assertEqual(
            report.canonical_package("【当季新品】双杯尝鲜套餐10"),
            "【当季新品】双杯尝鲜套餐10",
        )

    def test_current_platforms_and_previous_period_share_clean_name(self):
        period = "2026-06-15"
        sources = {
            "Ora外送商品数据.xlsx": pd.DataFrame({
                "date_id": [period], "sku_name": ["冰 美 式"],
                "quantity": [4], "gross_amount": [100],
            }),
            "美团商品数据.xlsx": pd.DataFrame({
                "日期": [period, period],
                "商品名": ["【当季新品】 双杯尝鲜套餐 1", "【当季新品】双杯尝鲜套餐"],
                "商品销量": [2, 3], "商品销售额": [60, 90],
            }),
            "饿了么商品数据.xlsx": pd.DataFrame({
                "日期": [period], "商品名称": ["【当季新品】双杯 尝鲜套餐1"],
                "销量": [4], "销售额": [120],
            }),
        }
        previous = Workbook()
        single = previous.active
        single.title = "商品销售排行-单品"
        single["B2"] = "饮品"
        single["C2"] = "冰 美 式"
        single["D2"] = 5
        single["B3"] = "总计"
        package = previous.create_sheet("商品销售排行-套餐")
        package["B2"] = "【当季新品】双杯尝鲜套餐1"
        package["C2"] = 7
        package["B3"] = "【当季新品】双杯尝鲜套餐"
        package["C3"] = 3
        package["A4"] = "总计"

        with patch.object(report, "read_excel_columns", side_effect=lambda name, *_args, **_kwargs: sources[name].copy()):
            singles, packages, prev_single, prev_package, _ = report.compute_products(previous, 7)

        self.assertEqual([(item["name"], item["qty"]) for item in singles], [("冰美式", 4)])
        target = "【当季新品】双杯尝鲜套餐"
        self.assertEqual([(item["name"], item["qty"]) for item in packages if item["name"] == target], [(target, 9)])
        self.assertEqual(prev_single["冰美式"], 5)
        self.assertEqual(prev_package[target], 10)
        self.assertTrue(all(not any(char.isspace() for char in item["name"]) for item in singles + packages))

        output = Workbook()
        output_single = output.active
        output_single.title = "商品销售排行-单品"
        output_single["B3"] = "总计"
        output_package = output.create_sheet("商品销售排行-套餐")
        output_package["A9"] = "总计"
        report.write_products(output, previous, singles, packages, prev_single, prev_package, 7, [target])
        self.assertEqual(output_single["C2"].value, "冰美式")
        self.assertEqual(output_single["L2"].value, "冰美式")
        self.assertEqual(output_package["B2"].value, target)
        self.assertEqual(output_package["F2"].value, 10)


if __name__ == "__main__":
    unittest.main()
