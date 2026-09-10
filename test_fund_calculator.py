import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

from fund_calculator import CalculationError, RunOptions, insurance_bucket, resolved_visit_type, run_file


class CalculatorTests(unittest.TestCase):
    def test_insurance_exact_names_and_codes(self):
        for value in ("职工", "城镇职工基本医疗保险", "310", 310, 310.0):
            self.assertEqual(insurance_bucket(value), 0, value)
        for value in ("城乡", "居民", "城乡居民", "城乡居民基本医疗保险", "390", 390, 390.0, " 城乡 "):
            self.assertEqual(insurance_bucket(value), 1, value)
        for value in ("其他", None, "", "391", "非职工", "居民补充保险", "城乡其他", "职 工"):
            self.assertIsNone(insurance_bucket(value), value)

    def test_insurance_names_and_codes_are_calculated(self):
        for header in ("险种类型", "INSUTYPE"):
            with self.subTest(header=header), tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "insurance.xlsx"
                workbook = Workbook()
                workbook.active.append(["定点编码", header, "医疗费总额", "基金支付总额", "符合范围金额", "医疗类别"])
                for insurance in ("城乡", "职工", "390", 390, 390.0, "310", 310):
                    for visit in (11, 21):
                        workbook.active.append(["H1", insurance, 100, 50, 20, visit])
                for insurance in ("其他", "非职工", "居民补充保险"):
                    workbook.active.append(["H1", insurance, 100, 50, 20, 21])
                workbook.save(path)
                workbook.close()
                result = run_file(path, RunOptions("通用", "自动识别", "廊坊市", "三级", None, None, True))
                self.assertEqual((result.successful, result.excluded, result.errors), (14, 3, 0))
                # 相同分组先汇总符合范围金额，再乘比例并四舍五入。
                self.assertEqual(result.total_fund, Decimal("166.22"))

    def test_group_totals_are_calculated_before_rate_and_rounding(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "grouped.xlsx"
            workbook = Workbook()
            source = workbook.active
            source.append(["定点编码", "险种类型", "医疗费总额", "基金支付总额", "符合范围金额", "医疗类别"])
            source.append(["H1", "城乡", 100, 0, 10, 11])
            source.append(["H1", "城乡", 100, 100, 90, 11])
            source.append(["H1", "城乡", 100, 50, 0.01, 21])
            source.append(["H1", "城乡", 100, 50, 0.01, 21])
            workbook.save(path)
            workbook.close()

            result = run_file(path, RunOptions("通用", "自动识别", "廊坊市", "三级", None, None, True))

            # 门诊：(10 + 90) * ((0 + 100) / (100 + 100)) = 50.00。
            # 住院：(0.01 + 0.01) * 56.84% = 0.01，不能逐行取整为 0.02。
            self.assertEqual(result.total_fund, Decimal("50.01"))
            saved = load_workbook(path, data_only=True)
            output = saved["基金测算"]
            self.assertEqual(output["F8"].value, 50)
            self.assertEqual(output["G8"].value, 0.5)
            self.assertEqual(output["F9"].value, 0.01)
            saved.close()

    def test_code_priority_and_name_fallback(self):
        for code in ("11", 14, 51.0):
            self.assertEqual(resolved_visit_type(code, "自动识别", "普通住院"), "门诊")
        for code in ("21", 22, 52.0):
            self.assertEqual(resolved_visit_type(code, "自动识别", "普通门诊"), "住院")
        for raw, name in (("普通门诊", None), (None, "普通门诊"), ("未知", "普通门诊")):
            self.assertEqual(resolved_visit_type(raw, "自动识别", name), "门诊")
        self.assertEqual(resolved_visit_type("11", "住院"), "住院")
        with self.assertRaises(CalculationError):
            resolved_visit_type("99", "自动识别", "未知")

    def test_exclusion_totals_and_original_row_numbers(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "input.xlsx"
            workbook = Workbook()
            source = workbook.active
            source.title = "明细"
            source.append(["定点编码", "定点名称", "险种类型", "医疗费总额", "基金支付总额", "符合范围金额1", "医疗类别", "医疗类别名称"])
            source.append(["H1", "医院", "职工", 100, 50, 20, 11, "普通住院"])
            source.append(["H1", "医院", "城乡居民基本医疗保险", 200, 100, 30, 21, "普通门诊"])
            source.append(["H1", "医院", "其他", None, None, None, 11, "普通门诊"])
            source.append(["H1", "医院", "其他", 999, 999, 999, 21, "普通住院"])
            source.append(["H1", "医院", None, 999, 999, 999, 11, "普通门诊"])
            source.append([None] * 8)
            source.append(["H1", "医院", "职工", 100, 50, 20, 99, "未知"])
            original = list(source.values)
            workbook.save(path)
            workbook.close()
            result = run_file(path, RunOptions("通用", "自动识别", "廊坊市", "三级", None, None, True))
            self.assertEqual((result.processed, result.successful, result.excluded, result.errors), (6, 2, 3, 1))
            self.assertEqual(result.total_fund, Decimal("27.05"))
            self.assertTrue(result.warnings[0].startswith("第 8 行："))
            saved = load_workbook(path)
            self.assertEqual(list(saved["明细"].values), original)
            output = saved["基金测算"]
            self.assertEqual(output["F5"].value, 3)
            self.assertEqual(sum(output.cell(row, 5).value for row in (8, 9)), 300)
            self.assertEqual(output["G8"].value, 0.5)
            self.assertEqual(output["G9"].value, 0.5684)
            saved.close()

    def test_name_only_column(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "name.xlsx"
            workbook = Workbook()
            workbook.active.append(["定点编码", "险种类型", "医疗费总额", "基金支付总额", "符合范围金额", "医疗类别名称"])
            workbook.active.append(["H1", "居民", 100, 50, 20, "普通门诊"])
            workbook.save(path)
            workbook.close()
            result = run_file(path, RunOptions("通用", "自动识别", "廊坊市", "三级", None, None, True))
            self.assertEqual((result.successful, result.errors, result.total_fund), (1, 0, Decimal("10.00")))


if __name__ == "__main__":
    unittest.main()
