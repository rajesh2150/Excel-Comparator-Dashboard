"""
Excel Comparator Module
Handles the core logic for comparing Base and SOW Excel files
"""

import pandas as pd
import numpy as np
from typing import Dict, Tuple


class ExcelComparator:
    """
    Compares two Excel files and updates SOW with data from Base file.
    Handles intelligent employee name matching with reverse name formats.
    """

    def __init__(self):
        """Initialize the comparator with default column names"""
        self.base_df = None
        self.sow_df = None
        self.base_lookup = {}
        self.stats = {
            "total_records": 0,
            "perfect_matches": 0,
            "project_mismatches": 0,
            "name_mismatches": 0,
            "updated_emp_ids": 0,
            "blank_emp_ids_found": 0,
        }

        # Column name mappings
        self.BASE_NAME_COL = "EMPLOYEE NAME"
        self.BASE_EMP_ID_COL = "EMPLOYEE NUMBER"
        self.BASE_PROJECT_COL = "PROJECT ID"

        self.SOW_NAME_COL = "Employee Name"
        self.SOW_EMP_ID_COL = "Mphasis Emp Id"
        self.SOW_PROJECT_COL = "Project Id"
        self.SOW_PRISM_PROJECT_COL = "PRISM Project Id"
        self.SOW_RESULT_COL = "Result"

    @staticmethod
    def normalize_name(name: str) -> str:
        """
        Normalize names for comparison.
        Handles:
        - lowercase
        - extra spaces
        - comma-separated format (Last, First)
        - reverse names

        Example:
            'John Deer', 'Deer John', 'Deer, John' all become comparable format

        Args:
            name: Input name string

        Returns:
            Normalized name string
        """
        if pd.isna(name):
            return ""

        name = str(name).strip().lower()

        # Handle comma-separated format (Last, First)
        name = name.replace(",", "")

        # Split words and filter empty strings
        parts = [p for p in name.split() if p]

        # Sort alphabetically for order-independent matching
        parts.sort()

        return " ".join(parts)

    def load_files(self, base_file_path: str, sow_file_path: str) -> bool:
        """
        Load Excel files from given paths.

        Args:
            base_file_path: Path to Base Abstract Excel file (.xlsb or .xlsx)
            sow_file_path: Path to SOW Work Orders Excel file (.xlsx)

        Returns:
            True if successful, False otherwise
        """
        try:
            # Determine engine based on file extension
            if base_file_path.endswith(".xlsb"):
                self.base_df = pd.read_excel(base_file_path, engine="pyxlsb")
            else:
                self.base_df = pd.read_excel(base_file_path)

            self.sow_df = pd.read_excel(sow_file_path)

            # Validate required columns
            self._validate_columns()

            # Prepare SOW dataframe columns
            self._prepare_sow_columns()

            return True

        except Exception as e:
            raise Exception(f"Error loading files: {str(e)}")

    def _validate_columns(self) -> None:
        """Validate that required columns exist in both dataframes"""
        base_required = [
            self.BASE_NAME_COL,
            self.BASE_EMP_ID_COL,
            self.BASE_PROJECT_COL,
        ]
        sow_required = [self.SOW_NAME_COL, self.SOW_EMP_ID_COL, self.SOW_PROJECT_COL]

        missing_base = [col for col in base_required if col not in self.base_df.columns]
        missing_sow = [col for col in sow_required if col not in self.sow_df.columns]

        if missing_base:
            raise ValueError(f"Missing columns in Base file: {missing_base}")
        if missing_sow:
            raise ValueError(f"Missing columns in SOW file: {missing_sow}")

    def _prepare_sow_columns(self) -> None:
        """Prepare SOW dataframe by ensuring result columns are object type"""
        if self.SOW_RESULT_COL not in self.sow_df.columns:
            self.sow_df[self.SOW_RESULT_COL] = ""
        self.sow_df[self.SOW_RESULT_COL] = self.sow_df[self.SOW_RESULT_COL].astype(
            object
        )

        if self.SOW_PRISM_PROJECT_COL not in self.sow_df.columns:
            self.sow_df[self.SOW_PRISM_PROJECT_COL] = ""
        self.sow_df[self.SOW_PRISM_PROJECT_COL] = self.sow_df[
            self.SOW_PRISM_PROJECT_COL
        ].astype(object)

        self.sow_df[self.SOW_EMP_ID_COL] = self.sow_df[self.SOW_EMP_ID_COL].astype(
            object
        )

    def _build_lookup(self) -> None:
        """Build lookup dictionary from Base dataframe"""
        self.base_lookup = {}

        for _, row in self.base_df.iterrows():
            normalized_name = self.normalize_name(row.get(self.BASE_NAME_COL))

            if normalized_name:  # Only add non-empty names
                self.base_lookup[normalized_name] = {
                    "employee_number": row.get(self.BASE_EMP_ID_COL),
                    "project_id": row.get(self.BASE_PROJECT_COL),
                    "employee_name": row.get(self.BASE_NAME_COL),
                }

    def compare(self) -> pd.DataFrame:
        """
        Compare Base and SOW files and update SOW with matching data.

        Returns:
            Updated SOW dataframe

        Raises:
            Exception if dataframes are not loaded
        """
        if self.base_df is None or self.sow_df is None:
            raise Exception("Files not loaded. Call load_files() first.")

        # Build lookup from base file
        self._build_lookup()

        # Reset statistics
        self._reset_stats()

        # Process each row in SOW file
        for index, row in self.sow_df.iterrows():
            self._process_row(index, row)

        return self.sow_df

    def _process_row(self, index: int, row: pd.Series) -> None:
        """
        Process a single row from SOW file.

        Args:
            index: Row index
            row: Row data
        """
        self.stats["total_records"] += 1

        sow_name = row.get(self.SOW_NAME_COL)
        sow_project = row.get(self.SOW_PROJECT_COL)
        current_emp_id = row.get(self.SOW_EMP_ID_COL)

        normalized_sow_name = self.normalize_name(sow_name)

        # Check if name found in base lookup
        if normalized_sow_name not in self.base_lookup:
            self.sow_df.at[index, self.SOW_RESULT_COL] = "Name mismatched"
            self.stats["name_mismatches"] += 1
            return

        # Name found - extract matched data
        matched_data = self.base_lookup[normalized_sow_name]
        base_emp_id = matched_data["employee_number"]
        base_project = matched_data["project_id"]

        # Update employee ID only if blank
        if self._is_blank(current_emp_id):
            self.stats["blank_emp_ids_found"] += 1
            self.sow_df.at[index, self.SOW_EMP_ID_COL] = base_emp_id
            self.stats["updated_emp_ids"] += 1
        else:
            self.stats["blank_emp_ids_found"] += 1

        # Always update PRISM Project ID
        self.sow_df.at[index, self.SOW_PRISM_PROJECT_COL] = base_project

        # Check project match
        sow_project_clean = str(sow_project).strip().lower()
        base_project_clean = str(base_project).strip().lower()

        if sow_project_clean == base_project_clean:
            self.sow_df.at[index, self.SOW_RESULT_COL] = "Perfect"
            self.stats["perfect_matches"] += 1
        else:
            self.sow_df.at[index, self.SOW_RESULT_COL] = "Project MisMatched"
            self.stats["project_mismatches"] += 1

    @staticmethod
    def _is_blank(value) -> bool:
        """
        Check if a value is blank/empty.

        Args:
            value: Value to check

        Returns:
            True if blank, False otherwise
        """
        return (
            pd.isna(value)
            or str(value).strip() == ""
            or str(value).strip().lower() == "blank"
        )

    def _reset_stats(self) -> None:
        """Reset statistics counters"""
        self.stats = {
            "total_records": 0,
            "perfect_matches": 0,
            "project_mismatches": 0,
            "name_mismatches": 0,
            "updated_emp_ids": 0,
            "blank_emp_ids_found": 0,
        }

    def get_stats(self) -> Dict:
        """
        Get comparison statistics.

        Returns:
            Dictionary containing statistics
        """
        return self.stats.copy()

    def get_result_dataframe(self) -> pd.DataFrame:
        """
        Get the updated SOW dataframe with key columns.

        Returns:
            Dataframe with selected columns for display
        """
        if self.sow_df is None:
            raise Exception("No data available")

        columns_to_show = [
            self.SOW_NAME_COL,
            self.SOW_EMP_ID_COL,
            self.SOW_PROJECT_COL,
            self.SOW_PRISM_PROJECT_COL,
            self.SOW_RESULT_COL,
        ]

        # Only include columns that exist
        available_columns = [
            col for col in columns_to_show if col in self.sow_df.columns
        ]

        return self.sow_df[available_columns].copy()

    def save_output(self, output_path: str) -> None:
        """
        Save updated SOW dataframe to Excel file.

        Args:
            output_path: Path to save the output file
        """
        if self.sow_df is None:
            raise Exception("No data to save")

        self.sow_df.to_excel(output_path, index=False)
