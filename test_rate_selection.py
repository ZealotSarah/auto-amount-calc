import unittest
from unittest.mock import patch

from fund_calculator import CalculatorApp


class RateSelectionTests(unittest.TestCase):
    def setUp(self):
        with patch("fund_calculator.load_records", return_value=[]):
            self.app = CalculatorApp()
        self.app.withdraw()
        self.addCleanup(self.app.destroy)

    def button(self, parent, text):
        for child in parent.winfo_children():
            if child.winfo_class() == "TButton" and child.cget("text") == text:
                return child
            found = self.button(child, text)
            if found is not None:
                return found

    def manager(self):
        self.app.manage_rates()
        return next(w for w in self.app.winfo_children() if w.winfo_class() == "Toplevel")

    def test_import_save_selects_year_city_and_calculation_rates(self):
        records = [dict(year="2024-2025（合计）", area="云浮市", level=level,
                        rates=["0.6776", "0.5987", "0.6157"])
                   for level in ("三级", "二级", "一级")]
        window = self.manager()
        with patch("fund_calculator.filedialog.askopenfilename", return_value="sample.pdf"), \
                patch("fund_calculator.read_pdf", return_value=records), \
                patch("fund_calculator.messagebox.showinfo"):
            self.button(window, "导入 PDF").invoke()
        with patch("fund_calculator.save_records") as save:
            self.button(window, "保存全部比例").invoke()
        save.assert_called_once_with(records)
        self.assertEqual(self.app.rate_year.get(), "2024-2025（合计）")
        self.assertEqual(self.app.pooling_area.get(), "云浮市")
        self.assertIn("云浮市", self.app.area_combo.cget("values"))
        self.assertNotIn("石家庄市", self.app.area_combo.cget("values"))
        self.assertEqual(str(self.app.selected_rates()["云浮市"]["三级"][0]), "0.6776")

    def test_save_without_changes_keeps_selection(self):
        self.app.pooling_area.set("石家庄市")
        window = self.manager()
        with patch("fund_calculator.save_records"):
            self.button(window, "保存全部比例").invoke()
        self.assertEqual(self.app.rate_year.get(), "2024-2025（内置）")
        self.assertEqual(self.app.pooling_area.get(), "石家庄市")

    def test_manual_entry_save_selects_new_city(self):
        window = self.manager()
        form = window.winfo_children()[0]
        values = ("2026", "广州市", "二级", "80", "70", "75")
        for widget in form.winfo_children():
            if widget.winfo_class() in {"TEntry", "TCombobox"}:
                value = values[int(widget.grid_info()["column"])]
                if widget.winfo_class() == "TCombobox":
                    widget.set(value)
                else:
                    widget.insert(0, value)
        self.button(window, "添加 / 更新录入行").invoke()
        with patch("fund_calculator.save_records"):
            self.button(window, "保存全部比例").invoke()
        self.assertEqual(self.app.rate_year.get(), "2026")
        self.assertEqual(self.app.pooling_area.get(), "广州市")
        self.assertIn("广州市", self.app.area_combo.cget("values"))


if __name__ == "__main__":
    unittest.main()
