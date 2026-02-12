#!/usr/bin/env python3
"""Build a governance audit workbook from HC TR and Data Index source files."""

from __future__ import annotations

import importlib.util
import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


# ===== CONFIG =====
LS_TR_FILE = Path("AUDITED REPORT.xlsx")
LS_TR_SHEET = "HC TR ContactCandidate Fields"
DATA_INDEX_FILE = Path("DATA INDEX - Attributes.xlsx")
OUTPUT_FILE = Path("HC_TR_Full_Governance_Audit.xlsx")
OUTPUT_SHEET = "Governance Audit"
# ==================


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", NS_MAIN)


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


# =======================
# Fallback XLSX Reader
# =======================

def _read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out = []
    for si in root.findall(f"{{{NS_MAIN}}}si"):
        texts = [t.text or "" for t in si.findall(f".//{{{NS_MAIN}}}t")]
        out.append("".join(texts))
    return out


def _sheet_path_by_name(zf: zipfile.ZipFile, sheet_name: str) -> str:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {r.attrib["Id"]: r.attrib["Target"] for r in rels.findall(f"{{{NS_PKG_REL}}}Relationship")}
    sheets = wb.findall(f".//{{{NS_MAIN}}}sheet")

    for s in sheets:
        if s.attrib.get("name") == sheet_name:
            rel_id = s.attrib.get(f"{{{NS_REL}}}id")
            return "xl/" + rel_map[rel_id].lstrip("/")

    names = [s.attrib.get("name") for s in sheets]
    raise KeyError(f"Sheet '{sheet_name}' not found. Available: {names}")


def _cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "s":
        v = cell.find(f"{{{NS_MAIN}}}v")
        if v is None:
            return ""
        return shared_strings[int(v.text)]
    v = cell.find(f"{{{NS_MAIN}}}v")
    return v.text if v is not None else ""


def _read_sheet_rows(path: Path, sheet_name: str):
    with zipfile.ZipFile(path) as zf:
        shared = _read_shared_strings(zf)
        sheet_path = _sheet_path_by_name(zf, sheet_name)
        ws = ET.fromstring(zf.read(sheet_path))

    rows = ws.findall(f".//{{{NS_MAIN}}}sheetData/{{{NS_MAIN}}}row")
    parsed = []

    for row in rows:
        row_data = {}
        for cell in row.findall(f"{{{NS_MAIN}}}c"):
            ref = cell.attrib.get("r")
            match = re.match(r"([A-Z]+)(\d+)", ref)
            col_letters = match.group(1)
            col_index = 0
            for ch in col_letters:
                col_index = col_index * 26 + (ord(ch) - 64)
            row_data[col_index] = _cell_text(cell, shared)
        parsed.append(row_data)

    headers = {k: v.strip() for k, v in parsed[0].items() if v.strip()}
    output = []

    for row in parsed[1:]:
        row_dict = {headers[k]: row.get(k, "") for k in headers}
        output.append(row_dict)

    return output


# =======================
# Main Logic
# =======================

def build_governance_audit():
    if not LS_TR_FILE.exists():
        raise FileNotFoundError(f"Missing file: {LS_TR_FILE}")
    if not DATA_INDEX_FILE.exists():
        raise FileNotFoundError(f"Missing file: {DATA_INDEX_FILE}")

    ls_rows = _read_sheet_rows(LS_TR_FILE, LS_TR_SHEET)
    data_rows = _read_sheet_rows(DATA_INDEX_FILE, "Sheet1")

    if "Field API Name" not in ls_rows[0]:
        raise KeyError("Missing column: Field API Name")

    normalized_index_names = {
        (row.get("Name", "") or "").strip().lower()
        for row in data_rows
    }

    output_rows = []
    original_headers = list(ls_rows[0].keys())

    for row in ls_rows:
        field = row.get("Field API Name", "")
        norm = field.strip().lower()
        exists = "Y" if norm in normalized_index_names else "N"
        category = classify_data_category(field)
        pii = classify_pii_level(field)
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

    # Write output
    import pandas as pd
    df = pd.DataFrame(output_rows)
    df = df[headers]

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=OUTPUT_SHEET)
        worksheet = writer.sheets[OUTPUT_SHEET]
        worksheet.freeze_panes(1, 0)

        for idx, col in enumerate(df.columns):
            max_len = max(len(str(col)), df[col].astype(str).map(len).max())
            worksheet.set_column(idx, idx, min(max(max_len + 2, 16), 60))

    print(f"Governance audit created: {OUTPUT_FILE}")


if __name__ == "__main__":
    build_governance_audit()
