from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

APP_NAME = "项目库>立项阶段 溯源"
BRANCH_RULES = json.loads(
    r'''
[
  {
    "name": "建议书",
    "match": "contains",
    "value": "建议书",
    "source_map": {
      "项目单位": "单位基本情况",
      "项目名称": "项目名称",
      "密级": "封面",
      "项目状态": "固定为：建议书",
      "项目建设地址（省）": "建设地点及自然条件",
      "项目建设地址（市）": "建设地点及自然条件",
      "项目建设地址（区/县）": "建设地点及自然条件",
      "项目建设具体地址": "建设地点及自然条件",
      "建设具体地点": "建设地点及自然条件",
      "建设周期（月）": "建设地点及自然条件",
      "主要建设内容": "建设内容及规模等",
      "项目建设目标": "建设目标",
      "新增土地面积（平方米）": "建设内容及规模等",
      "新增建筑面积（平方米）": "建设内容及规模等",
      "改造建筑面积（平方米）": "建设内容及规模等",
      "新增工艺设备台套数": "建设内容及规模等",
      "改造设备台套数": "建设内容及规模等",
      "招标设备台套数": "附表：工艺设备明细表",
      "进口设备台套数": "附表：工艺设备明细表",
      "工艺设备购置费（万元）": "投资估算和资金来源",
      "建筑工程费（万元）": "投资估算和资金来源",
      "安装工程费（万元）": "投资估算和资金来源",
      "其他费用（万元）": "投资估算和资金来源",
      "项目总投资（万元）": "投资估算和资金来源",
      "时间": "从文档描述中提取，如果没有明确项目时间，就不取值；如果可以获取到的具体时间，将提取的时间转换为对应的国家“五年规划”时期，例如：十三五，十四五"
    }
  },
  {
    "name": "批复",
    "match": "contains",
    "value": "批复",
    "source_map": {
      "项目单位": "项目建设单位",
      "项目名称": "投资估算表",
      "密级": "封面",
      "项目状态": "固定为：建议书",
      "项目建设地址（省）": "建设地址",
      "项目建设地址（市）": "建设地址",
      "项目建设地址（区/县）": "建设地址",
      "项目建设具体地址": "建设地址",
      "建设周期（月）": "",
      "首批投资计划下达时间": "",
      "首批投资计划下达文号": "",
      "立项批复日期": "立项批文的其他事项的签章处提取",
      "立项批复文号": "立项批文的副标题",
      "建设地点": "建设地址",
      "主要建设内容": "主要建设内容",
      "项目建设目标": "建设目标",
      "项目批复资金中央内预算": "投资规模及资金来源",
      "项目批复资金银行贷款": "投资规模及资金来源",
      "项目批复资金自筹资金": "投资规模及资金来源",
      "项目批复总投资": "投资规模及资金来源",
      "新增土地面积（平方米）": "主要建设内容",
      "新增建筑面积（平方米）": "主要建设内容",
      "改造建筑面积（平方米）": "主要建设内容",
      "新增工艺设备台套数": "主要建设内容",
      "改造设备台套数": "主要建设内容",
      "招标设备台套数": "附表：工艺设备明细表",
      "进口设备台套数": "附表：工艺设备明细表",
      "工艺设备购置费（万元）": "投资估算表",
      "建筑工程费（万元）": "投资估算表",
      "安装工程费（万元）": "投资估算表",
      "其他费用（万元）": "投资估算表",
      "项目总投资（万元）": "投资估算表",
      "时间": "从文档描述中提取，如果没有明确项目时间，就不取值；如果可以获取到的具体时间，将提取的时间转换为对应的国家“五年规划”时期，例如：十三五，十四五"
    }
  }
]
'''
)

WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
SHEET_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
OFFICE_REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
TAG_RE = re.compile(r"<[^>]+>")


def configure_stdio() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")


def to_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)
    return json.dumps(value, ensure_ascii=False)


def read_plain_text(path: Path) -> str:
    for encoding in ("utf-8", "gb18030", "utf-16"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="ignore")


def clean_heading_text(text: str) -> str:
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"^\d+(?:\.\d+)*\s*", "", text)
    return text.strip()


def markdown_table(rows: list[list[str]]) -> str:
    if not rows:
        return ""
    width = max(len(row) for row in rows)
    normalized = [row + [""] * (width - len(row)) for row in rows]
    lines = [
        "| " + " | ".join(normalized[0]) + " |",
        "| " + " | ".join(["---"] * width) + " |",
    ]
    for row in normalized[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def format_docx_paragraph(text: str, style_val: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return ""
    style_val = unicodedata.normalize("NFKC", style_val or "")
    if "Heading" in style_val:
        level_match = re.search(r"(\d+)", style_val)
        level = min(int(level_match.group(1)) if level_match else 2, 6)
        return f'{"#" * level} {clean_heading_text(text)}'
    if re.match(r"^\d+(?:\.\d+)*\s*[\u4e00-\u9fffA-Za-z]", text) and len(text) <= 80:
        return f"## {clean_heading_text(text)}"
    if len(text) <= 36 and re.match(r"^[\u4e00-\u9fffA-Za-z0-9（）()、\-—\s]+$", text):
        return f"## {clean_heading_text(text)}"
    return text


def read_docx_text(path: Path) -> str:
    try:
        with zipfile.ZipFile(path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        body = root.find("w:body", WORD_NS)
        blocks: list[str] = []
        if body is not None:
            for child in body:
                local = child.tag.rsplit("}", 1)[-1]
                if local == "p":
                    style_node = child.find("w:pPr/w:pStyle", WORD_NS)
                    style_val = ""
                    if style_node is not None:
                        style_val = style_node.attrib.get(
                            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", ""
                        )
                    text = "".join(node.text or "" for node in child.findall(".//w:t", WORD_NS)).strip()
                    line = format_docx_paragraph(text, style_val)
                    if line:
                        blocks.append(line)
                elif local == "tbl":
                    rows: list[list[str]] = []
                    for row in child.findall("w:tr", WORD_NS):
                        values: list[str] = []
                        for cell in row.findall("w:tc", WORD_NS):
                            text = "".join(node.text or "" for node in cell.findall(".//w:t", WORD_NS))
                            text = re.sub(r"\s+", " ", text).strip()
                            values.append(text)
                        if any(values):
                            rows.append(values)
                    table_text = markdown_table(rows)
                    if table_text:
                        blocks.append(table_text)
        text = "\n\n".join(blocks).strip()
        if text:
            return text
    except Exception:
        pass

    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    xml = xml.replace("</w:p>", "\n").replace("</w:tr>", "\n")
    text = TAG_RE.sub("", xml)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return "\n".join(lines)


def col_to_index(ref: str) -> int:
    letters = "".join(ch for ch in ref if ch.isalpha()).upper()
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch) - 64)
    return value or 1


def parse_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    result: list[str] = []
    for item in root.findall("a:si", SHEET_NS):
        text = "".join(node.text or "" for node in item.findall(".//a:t", SHEET_NS))
        result.append(unicodedata.normalize("NFKC", text))
    return result


def workbook_sheet_targets(zf: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib.get("Id", ""): rel.attrib.get("Target", "") for rel in rels.findall("r:Relationship", REL_NS)}
    result: list[tuple[str, str]] = []
    for sheet in workbook.findall(".//a:sheets/a:sheet", SHEET_NS):
        name = unicodedata.normalize("NFKC", sheet.attrib.get("name", "Sheet"))
        rid = sheet.attrib.get(OFFICE_REL_ID, "")
        target = rel_map.get(rid, "")
        if not target:
            continue
        if target.startswith("/"):
            target = target.lstrip("/")
        elif not target.startswith("xl/"):
            target = "xl/" + target.lstrip("./")
        result.append((name, target))
    return result


def parse_sheet_rows(xml_bytes: bytes, shared_strings: list[str]) -> list[list[str]]:
    root = ET.fromstring(xml_bytes)
    rows: list[list[str]] = []
    for row in root.findall(".//a:sheetData/a:row", SHEET_NS):
        values: dict[int, str] = {}
        max_col = 0
        for cell in row.findall("a:c", SHEET_NS):
            ref = cell.attrib.get("r", "")
            col_idx = col_to_index(ref)
            max_col = max(max_col, col_idx)
            cell_type = cell.attrib.get("t", "")
            if cell_type == "inlineStr":
                value = "".join(node.text or "" for node in cell.findall(".//a:t", SHEET_NS))
            else:
                value_node = cell.find("a:v", SHEET_NS)
                value = value_node.text if value_node is not None and value_node.text is not None else ""
                if cell_type == "s" and value.isdigit():
                    index = int(value)
                    value = shared_strings[index] if 0 <= index < len(shared_strings) else value
                elif cell_type == "b":
                    value = "TRUE" if value == "1" else "FALSE"
            values[col_idx] = unicodedata.normalize("NFKC", value).strip()
        if max_col:
            row_values = [values.get(idx, "") for idx in range(1, max_col + 1)]
            if any(item for item in row_values):
                rows.append(row_values)
    return rows


def read_xlsx_sections(path: Path) -> list[str]:
    sections: list[str] = []
    with zipfile.ZipFile(path) as zf:
        shared_strings = parse_shared_strings(zf)
        for sheet_name, target in workbook_sheet_targets(zf):
            if target not in zf.namelist():
                continue
            rows = parse_sheet_rows(zf.read(target), shared_strings)
            if not rows:
                continue
            sections.append(f"## {sheet_name}\n{markdown_table(rows)}")
    return sections


def read_csv_sections(path: Path) -> list[str]:
    text = read_plain_text(path)
    rows: list[list[str]] = []
    for row in csv.reader(text.splitlines()):
        values = [unicodedata.normalize("NFKC", item).strip() for item in row]
        if any(values):
            rows.append(values)
    if not rows:
        return []
    return [f"## {path.stem}\n{markdown_table(rows)}"]


def build_document_text(file_paths: list[Path]) -> str:
    blocks: list[str] = []
    for path in file_paths:
        suffix = path.suffix.lower()
        if suffix == ".docx":
            text = read_docx_text(path)
            if text:
                blocks.append(text)
        elif suffix == ".xlsx":
            blocks.extend(read_xlsx_sections(path))
        elif suffix == ".csv":
            blocks.extend(read_csv_sections(path))
        elif suffix in {".txt", ".md", ".json", ".html"}:
            text = read_plain_text(path)
            if text.strip():
                blocks.append(text)
    return "\n\n".join(block.strip() for block in blocks if block.strip())


def split_sections(text: str) -> dict[str, str]:
    text = to_text(text)
    result: dict[str, str] = {}
    if not text.strip():
        return result

    current_title: str | None = None
    current_lines: list[str] = []

    def flush() -> None:
        nonlocal current_title, current_lines
        if not current_title:
            current_lines = []
            return
        content = "\n".join(line for line in current_lines if line.strip()).strip()
        if not content:
            current_lines = []
            return
        aliases = {current_title, clean_heading_text(current_title)}
        for alias in list(aliases):
            for part in re.split(r"\s*[-—]\s*", alias):
                part = clean_heading_text(part)
                if part:
                    aliases.add(part)
        for alias in aliases:
            if alias and alias not in result:
                result[alias] = content
        current_lines = []

    for key, value in re.findall(r"([^\n：:]{1,40})[：:]+\s*([^\n]{1,200})", text):
        key = clean_heading_text(key)
        value = value.strip()
        if key and value and key not in result:
            result[key] = value

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        heading_match = re.match(r"^(#+)\s*(.+)$", line)
        numbered = re.match(r"^\d+(?:\.\d+)*\s*[\u4e00-\u9fffA-Za-z].*$", line) and len(line) <= 80
        if heading_match or numbered:
            flush()
            current_title = clean_heading_text(heading_match.group(2) if heading_match else line)
            current_lines = []
            continue
        if current_title is None:
            current_title = "全文"
        current_lines.append(line)
    flush()
    result.setdefault("全文", text[:20000])
    return result


def normalize_text(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "")
    text = text.replace("（", "(").replace("）", ")")
    return re.sub(r"\s+", "", text)


def guidance_section_tokens(guidance: str) -> list[str]:
    guidance = unicodedata.normalize("NFKC", guidance or "")
    normalized = guidance.replace("章节处", "章节").replace("章节中", "章节").replace("章节", "|").replace("处", "|").replace("中", "|")
    normalized = normalized.replace("从该文档的", "").replace("从该文档", "").replace("从文档描述", "")
    parts = re.split(r"[|,，。；;：:]+", normalized)
    tokens: list[str] = []
    for part in parts:
        part = part.strip().strip("-").strip()
        if not part:
            continue
        for item in re.split(r"\s*[-—]\s*", part):
            item = clean_heading_text(item)
            if item and len(item) >= 2 and item not in tokens:
                tokens.append(item)
    return tokens


def find_relevant_section_text(guidance: str, section_map: dict[str, str], doc_text: str) -> str:
    tokens = guidance_section_tokens(guidance)
    if not tokens:
        return doc_text
    normalized_titles = {title: normalize_text(title) for title in section_map}
    matches: list[str] = []
    for token in tokens:
        normalized_token = normalize_text(token)
        for title, content in section_map.items():
            normalized_title = normalized_titles[title]
            if normalized_token and (normalized_token in normalized_title or normalized_title in normalized_token):
                if content not in matches:
                    matches.append(content)
    if matches:
        return "\n\n".join(matches)
    normalized_tokens = [normalize_text(token) for token in tokens]
    lines = [line.strip() for line in doc_text.splitlines() if line.strip()]
    matched_lines = [line for line in lines if any(token and token in normalize_text(line) for token in normalized_tokens)]
    if matched_lines:
        return "\n".join(matched_lines[:12])
    return doc_text


def search_value_by_keywords(text: str, keywords: list[str], numeric: bool = False) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for keyword in keywords:
        direct = re.search(rf"{re.escape(keyword)}\s*[：:]\s*([^\n]{{1,200}})", text)
        if direct:
            value = direct.group(1).strip().strip("|").strip()
            if numeric:
                number = re.search(r"-?\d+(?:\.\d+)?", value)
                if number:
                    return number.group(0)
            elif value:
                return value
        for line in lines:
            if keyword not in line:
                continue
            if numeric:
                numbers = re.findall(r"-?\d+(?:\.\d+)?", line)
                if numbers:
                    return numbers[-1]
            else:
                if "|" in line:
                    cells = [cell.strip() for cell in line.strip("|").split("|") if cell.strip()]
                    for index, cell in enumerate(cells):
                        if keyword in cell and index + 1 < len(cells):
                            return cells[index + 1]
                    if cells:
                        return cells[-1]
                cleaned = re.sub(rf".*?{re.escape(keyword)}\s*", "", line)
                cleaned = cleaned.strip("：:|- ").strip()
                if cleaned:
                    return cleaned
    return ""


def extract_fixed_value(guidance: str) -> str:
    guidance = unicodedata.normalize("NFKC", guidance or "")
    patterns = [
        r"固定为[:：]\s*([^，。,；;\n]+)",
        r"固定值[:：]\s*([^，。,；;\n]+)",
        r"固定取值为[“\"]([^”\"]+)[”\"]",
        r"直接输出[“\"]([^”\"]+)[”\"]",
    ]
    for pattern in patterns:
        match = re.search(pattern, guidance)
        if match:
            return match.group(1).strip()
    return ""


def extract_date(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "")
    match = re.search(r"(20\d{2})[年/\-.](\d{1,2})(?:[月/\-.](\d{1,2}))?", text)
    if not match or match.group(3) is None:
        return ""
    year = int(match.group(1))
    month = int(match.group(2))
    day = int(match.group(3))
    return f"{year:04d}/{month:02d}/{day:02d}"


def extract_document_number(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "")
    patterns = [
        r"([^\s，。；;()（）]{0,20}〔20\d{2}〕\d+号)",
        r"([^\s，。；;()（）]{0,20}\[20\d{2}\]\d+号)",
        r"([^\s，。；;()（）]{0,20}20\d{2}[^\s，。；;()（）]{0,8}\d+号)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1).strip()
    return ""


def extract_address_components(address: str) -> dict[str, str]:
    address = unicodedata.normalize("NFKC", address or "").strip()
    result = {"省": "", "市": "", "区县": "", "详细": address}
    province = re.search(r"((?:[^省市区县]{2,20}省)|(?:北京|天津|上海|重庆)市|(?:内蒙古|广西|西藏|宁夏|新疆)自治区)", address)
    city = re.search(r"([^省]{2,20}市)", address)
    district = re.search(r"([^市]{1,20}(?:区|县|旗))", address)
    if province:
        result["省"] = province.group(1)
    if city:
        result["市"] = city.group(1)
    if district:
        result["区县"] = district.group(1)
    return result


def infer_five_year_plan(text: str) -> str:
    years = [int(item) for item in re.findall(r"20\d{2}", unicodedata.normalize("NFKC", text or ""))]
    if not years:
        return ""
    year = max(years)
    if 2011 <= year <= 2015:
        return "十二五"
    if 2016 <= year <= 2020:
        return "十三五"
    if 2021 <= year <= 2025:
        return "十四五"
    if 2026 <= year <= 2030:
        return "十五五"
    return ""


def first_meaningful_line(text: str) -> str:
    for line in text.splitlines():
        line = line.strip().strip("|").strip()
        if not line or line.startswith("---"):
            continue
        if line.startswith("#"):
            continue
        return line
    return ""


def format_numeric(value: str, label: str) -> str:
    if not value:
        return ""
    number = float(value)
    if label.endswith("套数") or label.endswith("次数") or label.endswith("周期") or label.endswith("周期（月）"):
        return str(int(round(number)))
    if number.is_integer():
        return str(int(number))
    return f"{number:.1f}"


def infer_value(label: str, guidance: str, doc_text: str, section_map: dict[str, str]) -> str:
    label = unicodedata.normalize("NFKC", label or "").strip()
    guidance = unicodedata.normalize("NFKC", guidance or "").strip()
    if not label:
        return ""

    fixed = extract_fixed_value(guidance)
    if fixed:
        return fixed

    relevant_text = find_relevant_section_text(guidance, section_map, doc_text)
    combined_text = relevant_text or doc_text

    if label == "密级":
        return search_value_by_keywords(doc_text, ["密级"]) or "非密"

    if label == "产品性质":
        return "民品" if infer_value("密级", guidance, doc_text, section_map) == "非密" else "军品"

    if label == "时间":
        return infer_five_year_plan(doc_text)

    if label == "项目单位":
        return search_value_by_keywords(combined_text, ["单位名称", "项目单位", "项目建设单位", "建设单位"]) or search_value_by_keywords(
            doc_text, ["单位名称", "项目单位", "项目建设单位", "建设单位"]
        )

    if label == "项目名称":
        return search_value_by_keywords(combined_text, ["项目名称", "工程名称", "建设项目名称"]) or search_value_by_keywords(
            doc_text, ["项目名称", "工程名称"]
        )

    if label in {"建设地点", "项目建设具体地址", "建设具体地点"}:
        return search_value_by_keywords(combined_text, ["建设地点", "建设地址", "厂址选择", "场址位置", "建设具体地点"]) or search_value_by_keywords(
            doc_text, ["建设地点", "建设地址"]
        )

    if label == "项目建设地址（省）":
        address = infer_value("建设地点", guidance, doc_text, section_map) or combined_text
        return extract_address_components(address).get("省", "")

    if label == "项目建设地址（市）":
        address = infer_value("建设地点", guidance, doc_text, section_map) or combined_text
        return extract_address_components(address).get("市", "")

    if label == "项目建设地址（区/县）":
        address = infer_value("建设地点", guidance, doc_text, section_map) or combined_text
        return extract_address_components(address).get("区县", "")

    if "日期" in label:
        return extract_date(combined_text) or extract_date(doc_text)

    if "文号" in label:
        return extract_document_number(combined_text) or extract_document_number(doc_text)

    if "周期" in label:
        value = search_value_by_keywords(combined_text, ["建设周期", "周期"], numeric=True) or search_value_by_keywords(
            doc_text, ["建设周期", "周期"], numeric=True
        )
        return format_numeric(value, label) if value else ""

    if label in {"主要建设内容", "项目建设目标", "项目建设成效", "竣工验收建设内容"}:
        value = search_value_by_keywords(combined_text, [label]) or search_value_by_keywords(combined_text, ["建设目标", "建设成效"])
        return value or first_meaningful_line(combined_text)

    numeric_keywords = {
        "新增土地面积（平方米）": ["新增土地面积"],
        "新增建筑面积（平方米）": ["新增建筑面积", "新建建筑面积"],
        "改造建筑面积（平方米）": ["改造建筑面积"],
        "新增工艺设备台套数": ["新增工艺设备", "新增设备"],
        "改造设备台套数": ["改造设备台套数", "改造设备"],
        "招标设备台套数": ["招标设备台套数", "招标设备"],
        "进口设备台套数": ["进口设备台套数", "进口设备"],
        "工艺设备购置费（万元）": ["工艺设备购置费", "设备购置费"],
        "建筑工程费（万元）": ["建筑工程费"],
        "安装工程费（万元）": ["安装工程费"],
        "其他费用（万元）": ["其他费用"],
        "项目总投资（万元）": ["项目总投资", "固定资产投资", "总投资"],
        "项目批复资金中央内预算": ["中央内预算", "中央预算内", "中央预算内专项投资"],
        "项目批复资金银行贷款": ["银行贷款"],
        "项目批复资金自筹资金": ["自筹资金"],
        "项目批复总投资": ["批复总投资", "总投资"],
        "土地购置费（元）": ["土地购置费", "土地征用费"],
        "投资累计下达中央内预算": ["累计下达中央内预算"],
        "投资累计下达中央预算内专项投资": ["累计下达中央预算内专项投资", "中央预算内专项投资"],
        "投资累计下达国有资本金预算": ["累计下达国有资本金预算", "国有资本金预算"],
        "投资累计下达银行贷款": ["累计下达银行贷款", "银行贷款"],
        "投资累计下达集团贷款": ["累计下达集团贷款", "集团贷款"],
        "投资累计下达自筹资金": ["累计下达自筹资金", "自筹资金"],
        "投资累计下达上市募集资金": ["累计下达上市募集资金", "上市募集资金"],
        "投资累计下达债券": ["累计下达债券", "债券"],
        "投资累计下达铺底流动资金": ["累计下达铺底流动资金", "铺底流动资金"],
        "其他投资累计下达": ["其他投资累计下达"],
        "投机累计下达总投资": ["累计下达总投资", "总投资"],
    }
    if label in numeric_keywords:
        value = search_value_by_keywords(combined_text, numeric_keywords[label], numeric=True) or search_value_by_keywords(
            doc_text, numeric_keywords[label], numeric=True
        )
        return format_numeric(value, label) if value else ""

    if label == "超期情况":
        return search_value_by_keywords(combined_text, ["超期情况"]) or first_meaningful_line(combined_text)

    value = search_value_by_keywords(combined_text, [label]) or search_value_by_keywords(doc_text, [label])
    return value or first_meaningful_line(combined_text)


def trim_source_text(text: str, value: str = "") -> str:
    text = unicodedata.normalize("NFKC", text or "").strip()
    if not text:
        return ""
    if value:
        idx = text.find(value)
        if idx >= 0 and len(text) > 240:
            start = max(0, idx - 120)
            end = min(len(text), idx + len(value) + 120)
            text = text[start:end].strip()
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if len(lines) > 10:
        lines = lines[:10]
    return "\n".join(lines)[:600].strip()


def find_source_snippet(guidance: str, section_map: dict[str, str], doc_text: str, value: str) -> str:
    guidance = unicodedata.normalize("NFKC", guidance or "").strip()
    fixed = extract_fixed_value(guidance)
    if fixed:
        return guidance
    if not guidance:
        return trim_source_text(value)
    relevant = find_relevant_section_text(guidance, section_map, doc_text)
    return trim_source_text(relevant, value)


def select_branch(question: str) -> dict[str, Any] | None:
    question = unicodedata.normalize("NFKC", question or "").strip()
    for branch in BRANCH_RULES:
        mode = branch.get("match")
        value = unicodedata.normalize("NFKC", branch.get("value", ""))
        if mode == "contains" and value in question:
            return branch
        if mode == "eq" and question == value:
            return branch
    return None


def run(file_paths: list[str], question: str) -> str:
    if not isinstance(file_paths, list) or not all(isinstance(item, str) for item in file_paths):
        raise TypeError("file_paths must be list[str]")
    paths = [Path(item).expanduser().resolve() for item in file_paths]
    for path in paths:
        if not path.is_file():
            raise FileNotFoundError(f"File not found: {path}")

    branch = select_branch(question)
    if branch is None:
        return ""

    doc_text = build_document_text(paths)
    section_map = split_sections(doc_text)
    parts: list[str] = []
    for label, guidance in branch.get("source_map", {}).items():
        value = to_text(infer_value(label, guidance, doc_text, section_map)).strip()
        if not value:
            continue
        parts.append(f"### {label}：")
        parts.append(value)
        source = find_source_snippet(guidance, section_map, doc_text, value)
        if source:
            parts.append("### 数据来源段落：")
            parts.append(source)
        parts.append("")
    return "\n".join(parts).strip()


def main() -> int:
    configure_stdio()
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--question", required=True, help="Question used to select the branch")
    parser.add_argument("files", nargs="+", help="Input document paths")
    args = parser.parse_args()
    print(run(args.files, args.question))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
