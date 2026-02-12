#!/usr/bin/env python3
"""Build a governance audit workbook from LS TR and Data Index source files."""

from __future__ import annotations

import importlib.util
import re
import zipfile
from pathlib import Path
from typing import Iterable
import xml.etree.ElementTree as ET


LS_TR_FILE = Path("LS_TR_ContactCandidate_Fields.xlsx")
LS_TR_SHEET = "HC TR ContactCandidate Fields"
DATA_INDEX_FILE = Path("DATA_INDEX_Attributes.xlsx")
OUTPUT_FILE = Path("LS_TR_Full_Governance_Audit.xlsx")
OUTPUT_SHEET = "Governance Audit"

FILE_FALLBACKS: dict[Path, list[Path]] = {
    LS_TR_FILE: [
        Path("LS TR ContactCandidate Fields.xlsx"),
        Path("LS_TR ContactCandidate Fields.xlsx"),
        Path("John - Customer.io Use Cases - Copy.xlsx"),
    ],
    DATA_INDEX_FILE: [Path("DATA INDEX - Attributes.xlsx")],
}

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", NS_MAIN)


def resolve_file(path: Path) -> Path:
    candidates = [path, *FILE_FALLBACKS.get(path, [])]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise FileNotFoundError(f"None of these files were found: {', '.join(map(str, candidates))}")


def classify_data_category(field_name: str) -> str:
    value = field_name.lower()
    if any(token in value for token in ("ssn", "bank", "routing", "tax")):
        return "Sensitive"
    if any(token in value for token in ("resume", "employment", "cover")):
        return "Operational"
    if any(token in value for token in ("status", "unit", "segment", "specialty", "vertical")):
        return "Marketing"
    if any(token in value for token in ("id", "index")):
        return "System"
    return "Evaluate"


def classify_pii_level(field_name: str) -> str:
    value = field_name.lower()
    if "ssn" in value:
        return "Restricted"
    if any(token in value for token in ("email", "phone", "address", "birth")):
        return "High"
    if "name" in value:
        return "Moderate"
    return "None"


def recommended_action(exists_in_cio: str, data_category: str) -> str:
    if exists_in_cio == "N":
        return "Remove"
    if data_category == "Sensitive":
        return "Remove"
    if data_category == "Marketing":
        return "Keep"
    return "Evaluate"


def _col_letters(index: int) -> str:
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out: list[str] = []
    for si in root.findall(f"{{{NS_MAIN}}}si"):
        texts = [t.text or "" for t in si.findall(f".//{{{NS_MAIN}}}t")]
        out.append("".join(texts))
    return out


def _sheet_path_by_name(zf: zipfile.ZipFile, sheet_name: str) -> str:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {r.attrib["Id"]: r.attrib["Target"] for r in rels.findall(f"{{{NS_PKG_REL}}}Relationship")}
    sheets = wb.findall(f".//{{{NS_MAIN}}}sheet")

    def norm(s: str) -> str:
        return s.strip().lower()

    target_sheet = None
    for s in sheets:
        if s.attrib.get("name") == sheet_name or norm(s.attrib.get("name", "")) == norm(sheet_name):
            target_sheet = s
            break
    if target_sheet is None:
        names = [s.attrib.get("name", "") for s in sheets]
        raise KeyError(f"Sheet '{sheet_name}' not found. Available sheets: {names}")

    rel_id = target_sheet.attrib.get(f"{{{NS_REL}}}id")
    if not rel_id or rel_id not in rel_map:
        raise KeyError(f"Relationship not found for sheet '{sheet_name}'")
    return "xl/" + rel_map[rel_id].lstrip("/")


def _cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "s":
        v = cell.find(f"{{{NS_MAIN}}}v")
        if v is None or v.text is None:
            return ""
        idx = int(v.text)
        return shared_strings[idx] if 0 <= idx < len(shared_strings) else ""
    if cell_type == "inlineStr":
        t = cell.find(f".//{{{NS_MAIN}}}t")
        return t.text if t is not None and t.text else ""
    v = cell.find(f"{{{NS_MAIN}}}v")
    return v.text if v is not None and v.text is not None else ""


def _read_sheet_rows(path: Path, sheet_name: str) -> list[dict[str, str]]:
    with zipfile.ZipFile(path) as zf:
        shared = _read_shared_strings(zf)
        sheet_path = _sheet_path_by_name(zf, sheet_name)
        ws = ET.fromstring(zf.read(sheet_path))

    rows = ws.findall(f".//{{{NS_MAIN}}}sheetData/{{{NS_MAIN}}}row")
    if not rows:
        return []

    parsed_rows: list[dict[int, str]] = []
    for row in rows:
        row_data: dict[int, str] = {}
        for cell in row.findall(f"{{{NS_MAIN}}}c"):
            ref = cell.attrib.get("r", "")
            m = re.match(r"([A-Z]+)(\d+)", ref)
            if not m:
                continue
            col_letters = m.group(1)
            col_index = 0
            for ch in col_letters:
                col_index = col_index * 26 + (ord(ch) - 64)
            row_data[col_index] = _cell_text(cell, shared)
        parsed_rows.append(row_data)

    header_map = parsed_rows[0]
    headers = {idx: val.strip() for idx, val in header_map.items() if val.strip()}
    out: list[dict[str, str]] = []
    for row_data in parsed_rows[1:]:
        row_dict = {hdr: row_data.get(idx, "") for idx, hdr in headers.items()}
        out.append(row_dict)
    return out


def _write_xlsx(path: Path, sheet_name: str, rows: list[dict[str, str]], headers: list[str]) -> None:
    widths = []
    for h in headers:
        max_len = len(h)
        for row in rows:
            max_len = max(max_len, len(str(row.get(h, ""))))
        widths.append(min(max(max_len + 2, 16), 60))

    wb_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}"><sheets><sheet name="{sheet_name}" sheetId="1" r:id="rId1"/></sheets></workbook>'''
    wb_rels = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{NS_PKG_REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'''
    root_rels = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{NS_PKG_REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'''
    styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="{NS_MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>'''

    cols_xml = "".join(
        f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
        for i, w in enumerate(widths, start=1)
    )

    def inline_cell(ref: str, value: str) -> str:
        v = (value or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return f'<c r="{ref}" t="inlineStr"><is><t>{v}</t></is></c>'

    xml_rows = []
    all_rows = [dict(zip(headers, headers)), *rows]
    for r_idx, row in enumerate(all_rows, start=1):
        cells = []
        for c_idx, h in enumerate(headers, start=1):
            ref = f"{_col_letters(c_idx)}{r_idx}"
            cells.append(inline_cell(ref, str(row.get(h, ""))))
        xml_rows.append(f'<row r="{r_idx}">' + "".join(cells) + "</row>")

    pane = '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
    ws_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{NS_MAIN}">{pane}<cols>{cols_xml}</cols><sheetData>{''.join(xml_rows)}</sheetData></worksheet>'''

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", wb_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/worksheets/sheet1.xml", ws_xml)
        zf.writestr("xl/styles.xml", styles)


def build_governance_audit_fallback() -> None:
    ls_tr_path = resolve_file(LS_TR_FILE)
    data_index_path = resolve_file(DATA_INDEX_FILE)

    ls_rows = _read_sheet_rows(ls_tr_path, LS_TR_SHEET)
    data_rows = _read_sheet_rows(data_index_path, "Sheet1")

    if not ls_rows:
        raise ValueError("No data rows found in LS TR source sheet.")

    if "Field API Name" not in ls_rows[0]:
        raise KeyError("Missing required column in LS TR file: 'Field API Name'")
    if not data_rows or "Name" not in data_rows[0]:
        raise KeyError("Missing required column in Data Index file: 'Name'")

    normalized_index_names = {
        (row.get("Name", "") or "").strip().lower()
        for row in data_rows
    }

    output_rows = []
    original_headers = list(ls_rows[0].keys())
    for row in ls_rows:
        field_api_name = (row.get("Field API Name", "") or "")
        norm = field_api_name.strip().lower()
        exists = "Y" if norm in normalized_index_names else "N"
        category = classify_data_category(field_api_name)
        pii = classify_pii_level(field_api_name)
        action = recommended_action(exists, category)

        new_row = dict(row)
        new_row["Exists in C.IO Data Index (Y/N)"] = exists
        new_row["Data Category"] = category
        new_row["PII Level"] = pii
        new_row["Recommended Action"] = action
        output_rows.append(new_row)

    headers = [
        *original_headers,
        "Exists in C.IO Data Index (Y/N)",
        "Data Category",
        "PII Level",
        "Recommended Action",
    ]
    _write_xlsx(OUTPUT_FILE, OUTPUT_SHEET, output_rows, headers)


def build_governance_audit_pandas() -> None:
    import pandas as pd

    ls_tr_path = resolve_file(LS_TR_FILE)
    data_index_path = resolve_file(DATA_INDEX_FILE)

    workbook = pd.ExcelFile(ls_tr_path)
    target_sheet = next(
        (s for s in workbook.sheet_names if s == LS_TR_SHEET or s.strip().lower() == LS_TR_SHEET.strip().lower()),
        None,
    )
    if target_sheet is None:
        raise KeyError(f"Sheet '{LS_TR_SHEET}' not found in {ls_tr_path}")

    ls_tr_df = pd.read_excel(ls_tr_path, sheet_name=target_sheet)
    data_index_df = pd.read_excel(data_index_path)

    if "Field API Name" not in ls_tr_df.columns:
        raise KeyError("Missing required column in LS TR file: 'Field API Name'")
    if "Name" not in data_index_df.columns:
        raise KeyError("Missing required column in Data Index file: 'Name'")

    normalized_index_names = set(data_index_df["Name"].fillna("").astype(str).str.strip().str.lower())
    result_df = ls_tr_df.copy()
    normalized_field_name = result_df["Field API Name"].fillna("").astype(str).str.strip().str.lower()

    result_df["Exists in C.IO Data Index (Y/N)"] = normalized_field_name.map(lambda n: "Y" if n in normalized_index_names else "N")
    result_df["Data Category"] = result_df["Field API Name"].fillna("").astype(str).map(classify_data_category)
    result_df["PII Level"] = result_df["Field API Name"].fillna("").astype(str).map(classify_pii_level)
    result_df["Recommended Action"] = result_df.apply(
        lambda r: recommended_action(r["Exists in C.IO Data Index (Y/N)"], r["Data Category"]), axis=1
    )

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        result_df.to_excel(writer, index=False, sheet_name=OUTPUT_SHEET)
        worksheet = writer.sheets[OUTPUT_SHEET]
        worksheet.freeze_panes(1, 0)
        for idx, col in enumerate(result_df.columns):
            series = result_df[col].fillna("").astype(str)
            max_len = max(len(str(col)), series.map(len).max() if not series.empty else 0)
            worksheet.set_column(idx, idx, min(max(max_len + 2, 16), 60))


def build_governance_audit() -> None:
    has_pandas_stack = all(importlib.util.find_spec(pkg) for pkg in ("pandas", "xlsxwriter", "openpyxl"))
    if has_pandas_stack:
        build_governance_audit_pandas()
    else:
        build_governance_audit_fallback()
    print(f"Governance audit created: {OUTPUT_FILE}")


if __name__ == "__main__":
    build_governance_audit()
