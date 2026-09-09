"""基金金额自动测算工具。

结果 Sheet 仅输出用户约定的七列汇总字段。住院比例内置自
《24年25年目录内（费用）住院基金支付比例》；门诊不使用该比例表。
"""

from __future__ import annotations

import datetime as dt
import os
import re
import shutil
import sys
import threading
import traceback
from collections import OrderedDict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


RATE_VERSION = "24-25 年住院基金支付比例"
OUTPUT_SHEET = "基金测算"

# 统筹区: {医疗机构级别: (职工, 居民, 合计)}。数值为小数比例。
FUND_RATES: dict[str, dict[str, tuple[Decimal, Decimal, Decimal]]] = {
    "河北省": {"三级": (Decimal("0.8423"), Decimal("0.6112"), Decimal("0.6860")), "二级": (Decimal("0.8882"), Decimal("0.7227"), Decimal("0.7550")), "一级": (Decimal("0.9109"), Decimal("0.8422"), Decimal("0.8511"))},
    "石家庄市": {"三级": (Decimal("0.8407"), Decimal("0.6158"), Decimal("0.6946")), "二级": (Decimal("0.9169"), Decimal("0.7657"), Decimal("0.7958")), "一级": (Decimal("0.9157"), Decimal("0.8730"), Decimal("0.8796"))},
    "辛集市": {"三级": (Decimal("0.8016"), Decimal("0.4953"), Decimal("0.5507")), "二级": (Decimal("0.8278"), Decimal("0.6711"), Decimal("0.6905")), "一级": (Decimal("0.8944"), Decimal("0.7424"), Decimal("0.7672"))},
    "唐山市": {"三级": (Decimal("0.8226"), Decimal("0.5689"), Decimal("0.6874")), "二级": (Decimal("0.8868"), Decimal("0.6471"), Decimal("0.7275")), "一级": (Decimal("0.9126"), Decimal("0.8150"), Decimal("0.8435"))},
    "秦皇岛市": {"三级": (Decimal("0.8368"), Decimal("0.5892"), Decimal("0.6892")), "二级": (Decimal("0.8974"), Decimal("0.7384"), Decimal("0.7970")), "一级": (Decimal("0.9029"), Decimal("0.8865"), Decimal("0.8899"))},
    "邯郸市": {"三级": (Decimal("0.8757"), Decimal("0.6893"), Decimal("0.7425")), "二级": (Decimal("0.8941"), Decimal("0.7667"), Decimal("0.7850")), "一级": (Decimal("0.9039"), Decimal("0.8510"), Decimal("0.8550"))},
    "邢台市": {"三级": (Decimal("0.8960"), Decimal("0.6712"), Decimal("0.7260")), "二级": (Decimal("0.9468"), Decimal("0.7669"), Decimal("0.7931")), "一级": (Decimal("0.9576"), Decimal("0.8614"), Decimal("0.8666"))},
    "保定市": {"三级": (Decimal("0.8238"), Decimal("0.5594"), Decimal("0.6386")), "二级": (Decimal("0.8496"), Decimal("0.7242"), Decimal("0.7406")), "一级": (Decimal("0.8874"), Decimal("0.8683"), Decimal("0.8703"))},
    "定州市": {"三级": (Decimal("0.8033"), Decimal("0.5159"), Decimal("0.5675")), "二级": (Decimal("0.8754"), Decimal("0.6263"), Decimal("0.6586")), "一级": (Decimal("0.9372"), Decimal("0.8123"), Decimal("0.8176"))},
    "张家口市": {"三级": (Decimal("0.8649"), Decimal("0.7032"), Decimal("0.7563")), "二级": (Decimal("0.8767"), Decimal("0.7908"), Decimal("0.8097")), "一级": (Decimal("0.9266"), Decimal("0.9030"), Decimal("0.9070"))},
    "承德市": {"三级": (Decimal("0.8555"), Decimal("0.5894"), Decimal("0.6625")), "二级": (Decimal("0.8593"), Decimal("0.7169"), Decimal("0.7382")), "一级": (Decimal("0.8777"), Decimal("0.8701"), Decimal("0.8705"))},
    "沧州市": {"三级": (Decimal("0.8475"), Decimal("0.5926"), Decimal("0.6521")), "二级": (Decimal("0.8639"), Decimal("0.7189"), Decimal("0.7388")), "一级": (Decimal("0.9258"), Decimal("0.8626"), Decimal("0.8672"))},
    "廊坊市": {"三级": (Decimal("0.8458"), Decimal("0.5684"), Decimal("0.6424")), "二级": (Decimal("0.9131"), Decimal("0.6778"), Decimal("0.7155")), "一级": (Decimal("0.9581"), Decimal("0.8326"), Decimal("0.8484"))},
    "衡水市": {"三级": (Decimal("0.8063"), Decimal("0.5892"), Decimal("0.6346")), "二级": (Decimal("0.8331"), Decimal("0.6632"), Decimal("0.6866")), "一级": (Decimal("0.8830"), Decimal("0.7769"), Decimal("0.7854"))},
    "雄安新区": {"三级": (Decimal("0.8283"), Decimal("0.6516"), Decimal("0.6719")), "二级": (Decimal("0.8589"), Decimal("0.6552"), Decimal("0.6713")), "一级": (Decimal("0.8462"), Decimal("0.6528"), Decimal("0.6537"))},
    "省本级": {"三级": (Decimal("0.8649"), Decimal("0.8649"), Decimal("0.8649")), "二级": (Decimal("0.9150"), Decimal("0.9150"), Decimal("0.9150")), "一级": (Decimal("0.8975"), Decimal("0.8975"), Decimal("0.8975"))},
}


class CalculationError(Exception):
    """A business-data error that should be shown to the operator."""


@dataclass
class RunOptions:
    rule_type: str
    visit_type: str
    pooling_area: str
    institution_level: str
    deduction_price: Decimal | None
    deduction_quantity: Decimal | None
    overwrite_result: bool


@dataclass
class FileResult:
    path: Path
    sheet_name: str
    processed: int
    successful: int
    errors: int
    total_fund: Decimal
    warnings: list[str]


def clean_header(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).replace("　", "")


def as_decimal(value: Any, label: str) -> Decimal:
    if value is None or str(value).strip() == "":
        raise CalculationError(f"{label}为空")
    try:
        return Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, ValueError) as exc:
        raise CalculationError(f"{label}不是数字：{value}") from exc


def optional_decimal(value: str, label: str) -> Decimal | None:
    if not value.strip():
        return None
    return as_decimal(value, label)


def text_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def find_column(headers: dict[str, int], candidates: Iterable[str]) -> int | None:
    for candidate in candidates:
        exact = clean_header(candidate)
        if exact in headers:
            return headers[exact]
    for candidate in candidates:
        prefix = clean_header(candidate)
        matches = [index for header, index in headers.items() if header.startswith(prefix)]
        if matches:
            return min(matches)
    return None


def determine_source_sheet(workbook) -> tuple[Any, int, dict[str, int]]:
    required = ("定点编码", "险种类型", "医疗费总额", "基金支付总额", "符合范围金额")
    best: tuple[int, int, Any, int, dict[str, int]] | None = None
    for sheet in workbook.worksheets:
        if sheet.title.startswith(OUTPUT_SHEET):
            continue
        for header_row in range(1, min(10, sheet.max_row) + 1):
            headers = {
                clean_header(cell.value): cell.column
                for cell in sheet[header_row]
                if clean_header(cell.value)
            }
            score = sum(find_column(headers, (field,)) is not None for field in required)
            candidate = (score, sheet.max_row, sheet, header_row, headers)
            if best is None or candidate[:2] > best[:2]:
                best = candidate
    if best is None or best[0] < 4:
        raise CalculationError("找不到业务数据 Sheet：至少需要定点编码、险种类型、医疗费总额、基金支付总额、符合范围金额等表头。")
    return best[2], best[3], best[4]


def insurance_bucket(insurance: str) -> int | None:
    text = insurance.replace(" ", "")
    if "职工" in text:
        return 0
    if "居民" in text or "城乡" in text:
        return 1
    return None


def resolved_visit_type(raw: Any, selected: str) -> str:
    if selected != "自动识别":
        return selected
    value = text_value(raw)
    if "住院" in value:
        return "住院"
    if "门诊" in value:
        return "门诊"
    raise CalculationError(f"无法从医疗类别识别住院或门诊：{value or '空值'}")


def run_file(path: Path, options: RunOptions) -> FileResult:
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise CalculationError("仅支持 .xlsx 和 .xlsm 文件。")

    keep_vba = path.suffix.lower() == ".xlsm"
    workbook = load_workbook(path, keep_vba=keep_vba)
    source, header_row, headers = determine_source_sheet(workbook)
    columns = {
        "code": find_column(headers, ("定点编码", "医疗机构编码")),
        "name": find_column(headers, ("定点名称", "医疗机构名称")),
        "insurance": find_column(headers, ("险种类型", "险种类别")),
        "medical_total": find_column(headers, ("医疗费总额", "医疗总额")),
        "fund_total": find_column(headers, ("基金支付总额",)),
        "scope_amount": find_column(headers, ("符合范围金额",)),
        "visit_type": find_column(headers, ("医疗类别",)),
        "violation_amount": find_column(headers, ("违规金额",)),
        "price": find_column(headers, ("单价",)),
        "quantity": find_column(headers, ("数量",)),
    }
    missing = [name for name in ("code", "insurance", "medical_total", "fund_total", "scope_amount") if columns[name] is None]
    if missing:
        labels = {"code": "定点编码", "insurance": "险种类型", "medical_total": "医疗费总额", "fund_total": "基金支付总额", "scope_amount": "符合范围金额"}
        raise CalculationError("缺少必填列：" + "、".join(labels[name] for name in missing))
    if options.visit_type == "自动识别" and columns["visit_type"] is None:
        raise CalculationError("选择“自动识别”时，原表必须有“医疗类别”列。")
    if options.rule_type == "串换" and columns["violation_amount"] is None and (options.deduction_price is None or options.deduction_quantity is None):
        raise CalculationError("串换规则且无“违规金额”列时，必须填写扣减单价和扣减数量。")

    output_name = create_output_sheet_name(workbook, options.overwrite_result)
    if output_name == OUTPUT_SHEET and OUTPUT_SHEET in workbook.sheetnames:
        del workbook[OUTPUT_SHEET]
    output = workbook.create_sheet(output_name)

    groups: OrderedDict[tuple[str, str, str], dict[str, Any]] = OrderedDict()
    warnings: list[str] = []
    processed = successful = errors = 0
    total_fund = Decimal("0")
    for row in source.iter_rows(min_row=header_row + 1, values_only=True):
        if not any(cell is not None and str(cell).strip() for cell in row):
            continue
        processed += 1
        get = lambda name: row[columns[name] - 1] if columns[name] else None
        code = text_value(get("code")) or "未填写编码"
        name = text_value(get("name")) or "未填写名称"
        insurance = text_value(get("insurance")) or "未填写险种"
        try:
            medical_total = as_decimal(get("medical_total"), "医疗费总额")
            fund_total = as_decimal(get("fund_total"), "基金支付总额")
            scope_amount = as_decimal(get("scope_amount"), "符合范围金额")
            if medical_total < 0 or fund_total < 0 or scope_amount < 0:
                raise CalculationError("医疗费总额、基金支付总额和符合范围金额不能为负数")
            visit_type = resolved_visit_type(get("visit_type"), options.visit_type)
            key = (code, insurance, visit_type)
            group = groups.setdefault(key, {
                "name": name, "medical_total": Decimal("0"), "fund_total": Decimal("0"),
                "fund_amount": Decimal("0"), "types": {visit_type}, "rates": set(),
            })
            group["medical_total"] += medical_total
            group["fund_total"] += fund_total

            if visit_type == "门诊":
                if medical_total == 0:
                    raise CalculationError("门诊医疗费总额为 0，无法计算")
                rate = fund_total / medical_total
                amount = scope_amount * rate
            else:
                bucket = insurance_bucket(insurance)
                if bucket is None:
                    raise CalculationError(f"住院险种无法识别为职工或居民：{insurance}")
                rate = FUND_RATES[options.pooling_area][options.institution_level][bucket]
                if options.rule_type == "通用":
                    base = scope_amount
                elif columns["violation_amount"] is not None:
                    base = as_decimal(get("violation_amount"), "违规金额")
                else:
                    price = as_decimal(get("price"), "原单价")
                    quantity = as_decimal(get("quantity"), "原数量")
                    assert options.deduction_price is not None and options.deduction_quantity is not None
                    base = (price - options.deduction_price) * (quantity - options.deduction_quantity)
                    if base < 0:
                        raise CalculationError("串换扣减后的计算基数为负数")
                if base < 0:
                    raise CalculationError("计算基数为负数")
                amount = base * rate

            amount = amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            group["fund_amount"] += amount
            group["rates"].add(rate)
            successful += 1
            total_fund += amount
        except CalculationError as exc:
            errors += 1
            if len(warnings) < 30:
                warnings.append(f"第 {header_row + processed} 行：{exc}")

    write_output_sheet(output, path, source.title, options, groups, processed, successful, errors, total_fund, warnings)
    try:
        workbook.save(path)
    except PermissionError as exc:
        raise CalculationError(f"无法保存文件。请关闭 Excel 中已打开的工作簿后重试：{path.name}") from exc
    return FileResult(path, output_name, processed, successful, errors, total_fund, warnings)


def create_output_sheet_name(workbook, overwrite: bool) -> str:
    if OUTPUT_SHEET not in workbook.sheetnames or overwrite:
        return OUTPUT_SHEET
    suffix = dt.datetime.now().strftime("%m%d_%H%M%S")
    return f"{OUTPUT_SHEET}_{suffix}"


def write_output_sheet(output, path: Path, source_name: str, options: RunOptions, groups: OrderedDict, processed: int, successful: int, errors: int, total_fund: Decimal, warnings: list[str]) -> None:
    navy = PatternFill("solid", fgColor="1F4E78")
    blue = PatternFill("solid", fgColor="D9EAF7")
    output.merge_cells("A1:G1")
    output["A1"] = "基金金额测算结果"
    output["A1"].font = Font(size=14, bold=True, color="FFFFFF")
    output["A1"].fill = navy
    output["A1"].alignment = Alignment(horizontal="center")
    metadata = [
        ("源文件", path.name), ("源数据 Sheet", source_name), ("规则大类", options.rule_type),
        ("业务类型", options.visit_type), ("参保地（住院比例）", options.pooling_area),
        ("医疗机构级别", options.institution_level), ("比例版本", RATE_VERSION),
        ("处理记录数", processed), ("计算成功数", successful), ("异常数", errors),
        ("基金金额合计", total_fund),
    ]
    for index, (label, value) in enumerate(metadata):
        row = 2 + index // 3
        column = 1 + (index % 3) * 2
        output.cell(row, column, label).font = Font(bold=True)
        output.cell(row, column).fill = blue
        output.cell(row, column + 1, value)
    output["D5"].number_format = "#,##0.00"

    header_row = 7
    headers = ["医疗机构编码", "医疗机构名称", "险种类别", "医疗类别", "医疗总额", "基金金额", "报销比例"]
    for column, header in enumerate(headers, 1):
        cell = output.cell(header_row, column, header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = navy
        cell.alignment = Alignment(horizontal="center")
    for row_number, ((code, insurance, visit_type), group) in enumerate(groups.items(), header_row + 1):
        rate = group_rate(group)
        output.cell(row_number, 1, code)
        output.cell(row_number, 2, group["name"])
        output.cell(row_number, 3, insurance)
        output.cell(row_number, 4, visit_type)
        output.cell(row_number, 5, float(group["medical_total"]))
        output.cell(row_number, 6, float(group["fund_amount"]))
        output.cell(row_number, 7, float(rate) if isinstance(rate, Decimal) else rate)
        output.cell(row_number, 5).number_format = "#,##0.00"
        output.cell(row_number, 6).number_format = "#,##0.00"
        if isinstance(rate, Decimal):
            output.cell(row_number, 7).number_format = "0.00%"
    if groups:
        last_row = header_row + len(groups)
        table = Table(displayName=unique_table_name(output), ref=f"A{header_row}:G{last_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        output.add_table(table)
    else:
        output.cell(header_row + 1, 1, "没有可汇总的数据")

    output.column_dimensions["A"].width = 20
    output.column_dimensions["B"].width = 46
    output.column_dimensions["C"].width = 28
    output.column_dimensions["D"].width = 14
    output.column_dimensions["E"].width = 16
    output.column_dimensions["F"].width = 16
    output.column_dimensions["G"].width = 14
    output.freeze_panes = "A8"
    if warnings:
        output["A" + str(header_row + len(groups) + 3)] = "异常示例（最多显示 30 条）"
        output["A" + str(header_row + len(groups) + 3)].font = Font(bold=True, color="C00000")
        for offset, warning in enumerate(warnings, 4):
            output["A" + str(header_row + len(groups) + offset)] = warning
            output.merge_cells(start_row=header_row + len(groups) + offset, start_column=1, end_row=header_row + len(groups) + offset, end_column=7)


def group_rate(group: dict[str, Any]) -> Decimal | str:
    types = group["types"]
    if types == {"住院"} and len(group["rates"]) == 1:
        return next(iter(group["rates"]))
    if types == {"门诊"}:
        return Decimal("0") if group["medical_total"] == 0 else group["fund_total"] / group["medical_total"]
    return "混合类型"


def unique_table_name(sheet) -> str:
    """Excel table names are workbook-wide, even on different sheets."""
    existing = {name for worksheet in sheet.parent.worksheets for name in worksheet.tables}
    base = "FundCalculationSummary"
    if base not in existing:
        return base
    number = 2
    while f"{base}{number}" in existing:
        number += 1
    return f"{base}{number}"


class CalculatorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("基金金额自动测算工具")
        self.geometry("760x610")
        self.minsize(680, 560)
        self.files: list[Path] = []
        self.rule_type = tk.StringVar(value="通用")
        self.visit_type = tk.StringVar(value="自动识别")
        self.pooling_area = tk.StringVar(value="张家口市")
        self.institution_level = tk.StringVar(value="三级")
        self.deduction_price = tk.StringVar()
        self.deduction_quantity = tk.StringVar()
        self.overwrite_result = tk.BooleanVar(value=False)
        self._build()

    def _build(self) -> None:
        container = ttk.Frame(self, padding=16)
        container.pack(fill="both", expand=True)
        ttk.Label(container, text="基金金额自动测算工具", font=("Microsoft YaHei UI", 16, "bold")).pack(anchor="w")
        ttk.Label(container, text="住院按 PDF 比例计算；门诊按（基金支付总额 ÷ 医疗费总额）× 符合范围金额计算。", foreground="#555555").pack(anchor="w", pady=(4, 12))
        file_bar = ttk.Frame(container)
        file_bar.pack(fill="x")
        ttk.Button(file_bar, text="选择 Excel 文件", command=self.pick_files).pack(side="left")
        ttk.Button(file_bar, text="清空", command=self.clear_files).pack(side="left", padx=8)
        self.file_label = ttk.Label(file_bar, text="尚未选择文件")
        self.file_label.pack(side="left", padx=8)

        params = ttk.LabelFrame(container, text="参数", padding=12)
        params.pack(fill="x", pady=12)
        self.add_combo(params, "规则大类", self.rule_type, ["通用", "串换"], 0, 0)
        self.add_combo(params, "业务类型", self.visit_type, ["自动识别", "住院", "门诊"], 0, 1)
        self.add_combo(params, "参保地（住院报销比例）", self.pooling_area, list(FUND_RATES), 1, 0)
        self.add_combo(params, "医疗机构级别", self.institution_level, ["三级", "二级", "一级"], 1, 1)
        ttk.Label(params, text="扣减单价（仅串换无违规金额列时）").grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(params, textvariable=self.deduction_price, width=26).grid(row=2, column=1, sticky="ew", pady=(10, 0))
        ttk.Label(params, text="扣减数量（仅串换无违规金额列时）").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(params, textvariable=self.deduction_quantity, width=26).grid(row=3, column=1, sticky="ew", pady=(8, 0))
        ttk.Checkbutton(params, text="覆盖已有“基金测算”Sheet（否则自动新建带时间的 Sheet）", variable=self.overwrite_result).grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))
        params.columnconfigure(1, weight=1)

        controls = ttk.Frame(container)
        controls.pack(fill="x")
        self.run_button = ttk.Button(controls, text="开始测算", command=self.start)
        self.run_button.pack(side="left")
        ttk.Label(controls, text="  输出写回所选原 Excel 文件，请先关闭已打开的文件。", foreground="#A33").pack(side="left")
        self.log = tk.Text(container, height=13, wrap="word", state="disabled", font=("Consolas", 10))
        self.log.pack(fill="both", expand=True, pady=(12, 0))

    def add_combo(self, parent, label: str, variable: tk.StringVar, values: list[str], row: int, column: int) -> None:
        base = column * 2
        ttk.Label(parent, text=label).grid(row=row, column=base, sticky="w", padx=(0, 10), pady=4)
        combo = ttk.Combobox(parent, textvariable=variable, values=values, state="readonly", width=26)
        combo.grid(row=row, column=base + 1, sticky="ew", padx=(0, 20), pady=4)
        parent.columnconfigure(base + 1, weight=1)

    def pick_files(self) -> None:
        selections = filedialog.askopenfilenames(title="选择需要测算的 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xlsm")])
        self.files = [Path(item) for item in selections]
        self.file_label.configure(text=f"已选择 {len(self.files)} 个文件" if self.files else "尚未选择文件")

    def clear_files(self) -> None:
        self.files = []
        self.file_label.configure(text="尚未选择文件")

    def start(self) -> None:
        if not self.files:
            messagebox.showwarning("未选择文件", "请先选择至少一个 Excel 文件。")
            return
        try:
            options = RunOptions(
                self.rule_type.get(), self.visit_type.get(), self.pooling_area.get(), self.institution_level.get(),
                optional_decimal(self.deduction_price.get(), "扣减单价"),
                optional_decimal(self.deduction_quantity.get(), "扣减数量"), self.overwrite_result.get(),
            )
        except CalculationError as exc:
            messagebox.showerror("参数错误", str(exc))
            return
        self.run_button.configure(state="disabled")
        self.write_log("开始处理...\n")
        threading.Thread(target=self.worker, args=(options,), daemon=True).start()

    def worker(self, options: RunOptions) -> None:
        results: list[FileResult] = []
        failures: list[str] = []
        for path in self.files:
            try:
                result = run_file(path, options)
                results.append(result)
                self.after(0, self.write_log, f"完成：{path.name} → {result.sheet_name}；成功 {result.successful}，异常 {result.errors}，基金金额 {result.total_fund:,.2f}\n")
            except Exception as exc:  # Keep batch processing independent per file.
                failures.append(f"{path.name}：{exc}")
                self.after(0, self.write_log, f"失败：{path.name} → {exc}\n")
        self.after(0, self.done, results, failures)

    def done(self, results: list[FileResult], failures: list[str]) -> None:
        self.run_button.configure(state="normal")
        message = f"处理完成：成功 {len(results)} 个文件"
        if failures:
            message += f"，失败 {len(failures)} 个文件"
        messagebox.showinfo("完成", message)

    def write_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", message)
        self.log.see("end")
        self.log.configure(state="disabled")


def main() -> None:
    app = CalculatorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
