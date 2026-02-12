#!/usr/bin/env python3
"""Build HC TR Governance Audit"""

from pathlib import Path
import pandas as pd


# ================= CONFIG =================
SOURCE_FILE = Path("John - Customer.io Use Cases - Copy.xlsx")
SOURCE_SHEET = "HC TR ContactCandidate Fields"

DATA_INDEX_FILE = Path("DATA INDEX - Attributes.xlsx")

OUTPUT_FILE = Path("HC_TR_Full_Governance_Audit.xlsx")
OUTPUT_SHEET = "Governance Audit"

API_COLUMN = "Field Analysis: Field Name"
DATA_INDEX_COLUMN = "Name"
# ==========================================


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


def recommended_action(exists: str, category: str) -> str:
    if exists == "N":
        return "Remove"
    if category == "Sensitive":
        return "Remove"
    if category == "Marketing":
        return "Keep"
    return "Evaluate"


def build_governance_audit():

    if not SOURCE_FILE.exists():
        raise FileNotFoundError(f"Missing source file: {SOURCE_FILE}")

    if not DATA_INDEX_FILE.exists():
        raise FileNotFoundError(f"Missing Data Index file: {DATA_INDEX_FILE}")

    # Load data
    source_df = pd.read_excel(SOURCE_FILE, sheet_name=SOURCE_SHEET)
    data_index_df = pd.read_excel(DATA_INDEX_FILE)

    if API_COLUMN not in source_df.columns:
        raise KeyError(f"Missing required column in HC sheet: '{API_COLUMN}'")

    if DATA_INDEX_COLUMN not in data_index_df.columns:
        raise KeyError(f"Missing required column in Data Index: '{DATA_INDEX_COLUMN}'")

    # Normalize names
    normalized_index_names = (
        data_index_df[DATA_INDEX_COLUMN]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    normalized_index_set = set(normalized_index_names)

    normalized_field_names = (
        source_df[API_COLUMN]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    # Exists in C.IO
    source_df["Exists in C.IO Data Index (Y/N)"] = normalized_field_names.map(
        lambda x: "Y" if x in normalized_index_set else "N"
    )

    # Data Category
    source_df["Data Category"] = source_df[API_COLUMN].fillna("").astype(str).map(classify_data_category)

    # PII Level
    source_df["PII Level"] = source_df[API_COLUMN].fillna("").astype(str).map(classify_pii_level)

    # Recommended Action
    source_df["Recommended Action"] = source_df.apply(
        lambda row: recommended_action(
            row["Exists in C.IO Data Index (Y/N)"],
            row["Data Category"],
        ),
        axis=1,
    )

    # Write output
    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        source_df.to_excel(writer, index=False, sheet_name=OUTPUT_SHEET)

        worksheet = writer.sheets[OUTPUT_SHEET]
        worksheet.freeze_panes(1, 0)

        for idx, col in enumerate(source_df.columns):
            max_len = max(
                len(str(col)),
                source_df[col].astype(str).map(len).max(),
            )
            worksheet.set_column(idx, idx, min(max(max_len + 2, 16), 60))

    print(f"HC Governance audit created: {OUTPUT_FILE}")


if __name__ == "__main__":
    build_governance_audit()
