import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from reimbursement_rates import load_records, save_records, parse_pdf_rows, validate_record, ocr_rows
from fund_calculator import RunOptions, run_file
import test_fund_calculator
from openpyxl import load_workbook


class RateTests(unittest.TestCase):
    def test_combined_year_and_province_city(self):
        records = parse_pdf_rows("2024-2025年合计 三级 二级 一级", [
            ["三级机构", "二级机构", "一级机构"],
            ["职工", "居民", "合计"] * 3,
            ["广东省", "云浮市", "门诊+住院", "67.76%", "59.87%", "61.57%", "69.35%", "69.68%", "69.62%", "80.53%", "75.40%", "75.73%"],
        ])
        self.assertEqual(len(records), 3)
        self.assertTrue(all(r["year"] == "2024-2025（合计）" and r["area"] == "云浮市" for r in records))
        self.assertEqual(records[2]["rates"], ["0.8053", "0.754", "0.7573"])
        self.assertEqual(validate_record(records[0]), records[0])

    def test_ocr_row_order_and_low_confidence(self):
        def item(x, y, text, confidence=1):
            return [[[x,y],[x+10,y],[x+10,y+10],[x,y+10]], text, confidence]
        self.assertEqual(ocr_rows([item(20,0,"居民"),item(0,0,"职工"),item(0,20,"云浮市"),item(40,0,"错误",0.5)]), [["职工", "居民"], ["云浮市"]])

    def test_split_header_with_intermediate_row(self):
        records = parse_pdf_rows("2026年三级二级一级", [
            ["地市", "三级", None, None, "二级", None, None, "一级", None, None],
            [None] * 10,
            [None, "职工", "居民", "合计", "职工", "居民", "合计", "职工", "居民", "合计"],
            ["石家庄市", "84.07%", "61.58%", "69.46%", "91.69%", "76.57%", "79.58%", "91.57%", "87.30%", "87.96%"],
        ])
        self.assertEqual(len(records), 3)
        self.assertEqual(records[1]["level"], "二级")
        self.assertEqual(Decimal(records[1]["rates"][0]), Decimal("0.9169"))

    def test_parse_rows(self):
        records = parse_pdf_rows("2026年基金支付比例", [
            ["地市", "等级", "职工", "居民", "合计"],
            ["济南市", "三级", "84.23%", "60%", "70%"],
            [None, "二级", "90%", "80%", "85%"],
        ])
        self.assertEqual(len(records), 2)
        self.assertEqual(records[0]["year"], "2026")
        self.assertEqual(records[1]["area"], "济南市")
        self.assertEqual(Decimal(records[0]["rates"][0]), Decimal("0.8423"))

    def test_unknown_columns_and_year_fail(self):
        for text, rows in [("无年份", []), ("2025年 2026年", []), ("2026年", [["济南市", "三级", "90%", "80%", "85%"]])]:
            with self.assertRaises(ValueError):
                parse_pdf_rows(text, rows)

    def test_invalid_rate(self):
        for value in ("NaN", "1.01", "-0.1"):
            with self.assertRaises(ValueError):
                validate_record(dict(year="2026", area="济南市", level="三级", rates=[value, "0.5", "0.6"]))

    def test_persistence_and_calculation(self):
        with tempfile.TemporaryDirectory() as directory:
            config = Path(directory) / "rates.json"
            record = dict(year="2026", area="济南市", level="三级", rates=["0.8", "0.7", "0.75"])
            save_records([record], config)
            self.assertEqual(load_records(config), [record])
            path = Path(directory) / "source.xlsx"
            test_fund_calculator.CalculatorTests.make_source(path)
            options = RunOptions("通用", "住院", "济南市", "三级", None, None, False,
                                 rate_year="2026", custom_rates={"济南市": {"三级": tuple(map(Decimal, record["rates"]))}})
            result = run_file(path, options)
            self.assertEqual(result.total_fund, Decimal("14.00"))
            workbook = load_workbook(path)
            self.assertTrue(any("2026" in str(cell.value) for row in workbook[result.sheet_name] for cell in row))
            workbook.close()


if __name__ == "__main__":
    unittest.main()
