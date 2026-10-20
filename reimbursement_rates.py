"""报销比例的本地保存与 PDF 表格识别。"""
import json
import os
import re
from decimal import Decimal
from pathlib import Path


LEVELS = ("三级", "二级", "一级")


def config_path():
    return Path(os.environ.get("APPDATA", str(Path.home()))) / "基金金额自动测算工具" / "rates.json"


def validate_record(record):
    year = str(record["year"]).strip()
    area = str(record["area"]).strip()
    if not re.fullmatch(r"20\d{2}(?:-20\d{2}（合计）)?", year) or not area:
        raise ValueError("请填写四位年份或合并年度（例如 2024-2025（合计））和地市。")
    if record["level"] not in LEVELS:
        raise ValueError("请选择医院等级。")
    rates = [Decimal(str(value)) for value in record["rates"]]
    if len(rates) != 3 or any(not v.is_finite() or not 0 <= v <= 1 for v in rates):
        raise ValueError("职工、居民、合计比例必须在 0%—100% 之间。")
    return dict(year=year, area=area, level=record["level"], rates=[str(v) for v in rates])


def load_records(path=None):
    path = path or config_path()
    if not path.exists():
        return []
    return [validate_record(record) for record in json.loads(path.read_text(encoding="utf-8"))]


def save_records(records, path=None):
    records = [validate_record(record) for record in records]
    path = path or config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def identify_years(text):
    combined = re.search(r"(20\d{2})\s*[-—–~～至]\s*(20\d{2})\s*年?\s*合计", text)
    if combined:
        return [f"{combined[1]}-{combined[2]}（合计）"]
    years = sorted(set(re.findall(r"(?<!\d)(20\d{2})\s*年", text)))
    years += ["20" + y for y in re.findall(r"(?<!\d)(\d{2})\s*年", text) if "20" + y not in years]
    if not years:
        raise ValueError("未识别到年份，请手动填写。")
    if len(years) > 1:
        raise ValueError("PDF 中有多个年份，无法确定各列对应年份，请手动核对录入。")
    return years


def parse_pdf_rows(text, rows):
    """仅识别年份明确、列含义明确的表格；不从文件名推测。"""
    years = identify_years(text)
    records = []
    area = ""
    header = ""
    for row in rows:
        cells = [re.sub(r"\s+", "", str(cell or "")) for cell in row]
        line = " ".join(cells)
        if "职工" in line and "居民" in line and "合计" in line:
            header += " " + line
            continue
        if all(level in line for level in LEVELS):
            header = line
            continue
        if not all(label in header for label in ("职工", "居民", "合计")):
            continue
        if not (header.index("职工") < header.index("居民") < header.index("合计")):
            continue
        found = re.findall(r"[\u4e00-\u9fff]+?(?:市|自治州|地区|新区|省)|省本级", line)
        if found:
            area = found[-1]
        levels = [level for level in LEVELS if level in line]
        values = []
        for cell in cells:
            if re.fullmatch(r"\d+(?:\.\d+)?[%％]?", cell):
                number = Decimal(cell.rstrip("%％"))
                values.append(number / 100 if cell.endswith(("%", "％")) or number > 1 else number)
        if not area:
            continue
        if len(levels) == 1 and len(values) == 3:
            groups = [(levels[0], values)]
        elif len(values) == 9 and all(level in text for level in LEVELS):
            # 仅接受表头明确按三级、二级、一级排列的宽表。
            if not all(level in header for level in LEVELS) or not (header.index("三级") < header.index("二级") < header.index("一级")):
                continue
            groups = [(level, values[i * 3:i * 3 + 3]) for i, level in enumerate(LEVELS)]
        else:
            continue
        for year in years:
            for level, rates in groups:
                records.append(validate_record(dict(year=year, area=area, level=level, rates=rates)))
    if not records:
        raise ValueError("未识别到列含义明确的比例表（扫描件或复杂表格请手动录入）。")
    return records


def ocr_rows(result):
    """按识别框的位置重建行，保留横向列顺序。"""
    lines = []
    items = sorted(result or [], key=lambda item: sum(p[1] for p in item[0]) / 4)
    for box, text, confidence in items:
        if confidence < 0.9:
            continue
        y = sum(p[1] for p in box) / 4
        height = max(p[1] for p in box) - min(p[1] for p in box)
        if not lines or abs(y - lines[-1][0]) > height * 0.45:
            lines.append((y, []))
        lines[-1][1].append((min(p[0] for p in box), text))
    return [[text for _, text in sorted(cells)] for _, cells in lines]


def scan_page(page, engine):
    import numpy as np
    image = page.to_image(resolution=180).original
    for angle in (0, 90, 180, 270):
        result, _ = engine(np.array(image.rotate(angle, expand=True)))
        rows = ocr_rows(result)
        text = "\n".join(" ".join(row) for row in rows)
        try:
            return parse_pdf_rows(text, rows)
        except ValueError:
            continue
    raise ValueError("扫描件文字识别后仍无法确定年份、地市或比例，请手动补录。")


def read_pdf(path):
    import pdfplumber
    records = []
    diagnostics = []
    engine = None
    with pdfplumber.open(path) as document:
        for page_number, page in enumerate(document.pages, 1):
            text = page.extract_text() or ""
            tables = page.extract_tables()
            if not text.strip():
                from rapidocr_onnxruntime import RapidOCR
                if engine is None:
                    engine = RapidOCR(intra_op_num_threads=2, inter_op_num_threads=2)
                try:
                    records.extend(scan_page(page, engine))
                except ValueError as exc:
                    diagnostics.append(f"第 {page_number} 页：{exc}")
                continue
            if not tables:
                diagnostics.append(f"第 {page_number} 页：未找到表格" if text.strip() else f"第 {page_number} 页：未提取到文字，可能是扫描件")
            for table in tables:
                try:
                    records.extend(parse_pdf_rows(text, table))
                except ValueError as exc:
                    diagnostics.append(f"第 {page_number} 页：{exc}")
    if not records:
        detail = "\n".join(dict.fromkeys(diagnostics))
        raise ValueError("无法识别 PDF 中的报销比例。\n" + detail + "\n可手动录入年份、地市和比例。")
    keys = {}
    for record in records:
        key = (record["year"], record["area"], record["level"])
        if key in keys and keys[key] != record:
            raise ValueError("PDF 存在同年份、地市和等级的不同报销比例，请手动核对录入。")
        keys[key] = record
    return list(keys.values())
