import pandas as pd
import numpy as np
from pathlib import Path

# =========================================================
# FILE PATHS
# =========================================================

BASE_FILE = "Base_Abstract_05182026.xlsb"
SOW_FILE = "SOW_Suppliers___Work_Orders_2 (40).xlsx"

OUTPUT_FILE = "Updated_SOW_Output.xlsx"

# =========================================================
# LOAD EXCEL FILES
# =========================================================

base_df = pd.read_excel(BASE_FILE, engine="pyxlsb")
sow_df = pd.read_excel(SOW_FILE)

# Ensure Result column exists and is object type for string assignment
if "Result" not in sow_df.columns:
    sow_df["Result"] = ""
sow_df["Result"] = sow_df["Result"].astype(object)

# Ensure PRISM Project Id column exists and is object type
if "PRISM Project Id" not in sow_df.columns:
    sow_df["PRISM Project Id"] = ""
sow_df["PRISM Project Id"] = sow_df["PRISM Project Id"].astype(object)

# Ensure Mphasis Emp Id is object type for mixed content
sow_df["Mphasis Emp Id"] = sow_df["Mphasis Emp Id"].astype(object)

# =========================================================
# COLUMN NAMES
# =========================================================

BASE_NAME_COL = "EMPLOYEE NAME"
BASE_EMP_ID_COL = "EMPLOYEE NUMBER"
BASE_PROJECT_COL = "PROJECT ID"

SOW_NAME_COL = "Employee Name"
SOW_EMP_ID_COL = "Mphasis Emp Id"
SOW_PROJECT_COL = "Project Id"
SOW_PRISM_PROJECT_COL = "PRISM Project Id"
SOW_RESULT_COL = "Result"

# =========================================================
# HELPER FUNCTION
# =========================================================


def normalize_name(name):
    """
    Normalize names for comparison.
    Handles:
    - lowercase
    - extra spaces
    - comma-separated format (Last, First)
    - reverse names
    Example:
        John Deer
        Deer John
        Deer, John
    All become comparable format
    """

    if pd.isna(name):
        return ""

    name = str(name).strip().lower()

    # Handle comma-separated format (Last, First) by removing comma and space
    name = name.replace(",", "")

    # Split words
    parts = name.split()

    # Remove empty strings
    parts = [p for p in parts if p]

    # Sort alphabetically to normalize regardless of order
    parts.sort()

    # Join back
    return " ".join(parts)


# =========================================================
# CREATE LOOKUP DICTIONARY
# =========================================================

base_lookup = {}

for _, row in base_df.iterrows():
    normalized_name = normalize_name(row.get(BASE_NAME_COL))

    base_lookup[normalized_name] = {
        "employee_number": row.get(BASE_EMP_ID_COL),
        "project_id": row.get(BASE_PROJECT_COL),
        "employee_name": row.get(BASE_NAME_COL),
    }

# =========================================================
# SUMMARY COUNTERS
# =========================================================

total_records = 0
perfect_matches = 0
project_mismatches = 0
name_mismatches = 0
updated_emp_ids = 0

# =========================================================
# PROCESS SOW FILE
# =========================================================

for index, row in sow_df.iterrows():
    total_records += 1

    sow_name = row.get(SOW_NAME_COL)
    sow_project = row.get(SOW_PROJECT_COL)
    current_emp_id = row.get(SOW_EMP_ID_COL)

    normalized_sow_name = normalize_name(sow_name)

    # -----------------------------------------------------
    # NAME NOT FOUND
    # -----------------------------------------------------

    if normalized_sow_name not in base_lookup:
        sow_df.at[index, SOW_RESULT_COL] = "Name mismatched"
        name_mismatches += 1
        continue

    # -----------------------------------------------------
    # NAME FOUND
    # -----------------------------------------------------

    matched_data = base_lookup[normalized_sow_name]

    base_emp_id = matched_data["employee_number"]
    base_project = matched_data["project_id"]

    # -----------------------------------------------------
    # UPDATE EMPLOYEE ID ONLY IF BLANK
    # -----------------------------------------------------

    if (
        pd.isna(current_emp_id)
        or str(current_emp_id).strip() == ""
        or str(current_emp_id).strip().lower() == "blank"
    ):
        sow_df.at[index, SOW_EMP_ID_COL] = base_emp_id
        updated_emp_ids += 1

    # -----------------------------------------------------
    # UPDATE PRISM PROJECT ID
    # -----------------------------------------------------

    sow_df.at[index, SOW_PRISM_PROJECT_COL] = base_project

    # -----------------------------------------------------
    # PROJECT MATCH CHECK
    # -----------------------------------------------------

    sow_project_clean = str(sow_project).strip().lower()
    base_project_clean = str(base_project).strip().lower()

    if sow_project_clean == base_project_clean:
        sow_df.at[index, SOW_RESULT_COL] = "Perfect"
        perfect_matches += 1

    else:
        sow_df.at[index, SOW_RESULT_COL] = "Project MisMatched"
        project_mismatches += 1

# =========================================================
# SAVE OUTPUT
# =========================================================

sow_df.to_excel(OUTPUT_FILE, index=False)

# =========================================================
# SUMMARY
# =========================================================

print("\n======================================")
print("Excel Comparison Completed")
print("======================================")

print(f"Total Records              : {total_records}")
print(f"Perfect Matches            : {perfect_matches}")
print(f"Project Mismatches         : {project_mismatches}")
print(f"Name Mismatches            : {name_mismatches}")
print(f"Employee IDs Updated       : {updated_emp_ids}")

print("\nUpdated File Saved As:")
print(OUTPUT_FILE)

print("======================================")
