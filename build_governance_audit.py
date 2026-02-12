#!/usr/bin/env python3
"""Build a governance audit workbook from LS TR and Data Index source files."""

from __future__ import annotations

from pathlib import Path

import pandas as pd


LS_TR_FILE = Path("LS_TR_ContactCandidate_Fields.xlsx")
LS_TR_SHEET = "LS TR ContactCandidate Fields"
DATA_INDEX_FILE = Path("DATA_INDEX_Attributes.xlsx")
OUTPUT_FILE = Path("LS_TR_Full_Governance_Audit.xlsx")
OUTPUT_SHEET = "Governance Audit"


# Common local filename variants seen in shared folders.
FILE_FALLBACKS: dict[Path, list[Path]] = {
    LS_TR_FILE: [Path("LS TR ContactCandidate Fields.xlsx")],
    DATA_INDEX_FILE: [Path("DATA INDEX - Attributes.xlsx")],
}


def resolve_file(path: Path) -> Path:
    """Return the first existing file path among the configured path and fallbacks."""
    candidates = [path, *FILE_FALLBACKS.get(path, [])]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    options = ", ".join(str(candidate) for candidate in candidates)
    raise FileNotFoundError(f"None of these files were found: {options}")


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


def build_governance_audit() -> None:
    ls_tr_path = resolve_file(LS_TR_FILE)
    data_index_path = resolve_file(DATA_INDEX_FILE)

    ls_tr_df = pd.read_excel(ls_tr_path, sheet_name=LS_TR_SHEET)
    data_index_df = pd.read_excel(data_index_path)

    required_ls_tr_col = "Field API Name"
    required_index_col = "Name"

    if required_ls_tr_col not in ls_tr_df.columns:
        raise KeyError(f"Missing required column in LS TR file: '{required_ls_tr_col}'")
    if required_index_col not in data_index_df.columns:
        raise KeyError(f"Missing required column in Data Index file: '{required_index_col}'")

    # Normalize for matching while preserving original source values.
    normalized_index_names = set(
        data_index_df[required_index_col]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    result_df = ls_tr_df.copy()
    normalized_field_name = (
        result_df[required_ls_tr_col].fillna("").astype(str).str.strip().str.lower()
    )
    result_df["Exists in C.IO Data Index (Y/N)"] = normalized_field_name.map(
        lambda name: "Y" if name in normalized_index_names else "N"
    )
    result_df["Data Category"] = result_df[required_ls_tr_col].fillna("").astype(str).map(
        classify_data_category
    )
    result_df["PII Level"] = result_df[required_ls_tr_col].fillna("").astype(str).map(
        classify_pii_level
    )
    result_df["Recommended Action"] = result_df.apply(
        lambda row: recommended_action(
            row["Exists in C.IO Data Index (Y/N)"], row["Data Category"]
        ),
        axis=1,
    )

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        result_df.to_excel(writer, index=False, sheet_name=OUTPUT_SHEET)
        worksheet = writer.sheets[OUTPUT_SHEET]
        worksheet.freeze_panes(1, 0)

        for col_index, column_name in enumerate(result_df.columns):
            column_series = result_df[column_name].fillna("").astype(str)
            max_len = max(
                len(str(column_name)),
                column_series.map(len).max() if not column_series.empty else 0,
            )
            worksheet.set_column(col_index, col_index, min(max(max_len + 2, 16), 60))

    print(f"Governance audit created: {OUTPUT_FILE}")


if __name__ == "__main__":
    build_governance_audit()
