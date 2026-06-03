"""
Excel Comparator Module
Handles the core logic for comparing Base and SOW Excel files
with filtering by Work Order Status and Employee ID blanks
"""

import pandas as pd
from typing import Dict


class ExcelComparator:
    """
    Compares two Excel files and updates SOW with data from Base file.
    Handles intelligent employee name matching with reverse name formats.
    """

    def __init__(self):
        """Initialize the comparator with default column names"""
        self.base_df = None
        self.sow_df = None
        self.sow_df_filtered = None  # Filtered SOW data
        self.base_lookup = {}
        self.stats = {
            "total_records_in_sow": 0,
            "work_order_status_filtered": 0,
            "blank_emp_ids_count": 0,
            "perfect_matches": 0,
            "name_mismatches": 0,
            "project_mismatches": 0,
            "updated_emp_ids": 0,
            # Status-wise breakdown
            "accepted_total": 0,
            "accepted_name_mismatch": 0,
            "accepted_project_mismatch": 0,
            "accepted_perfect": 0,
            "activated_total": 0,
            "activated_name_mismatch": 0,
            "activated_project_mismatch": 0,
            "activated_perfect": 0,
            "confirmed_total": 0,
            "confirmed_name_mismatch": 0,
            "confirmed_project_mismatch": 0,
            "confirmed_perfect": 0,
            "pending_approval_total": 0,
            "pending_approval_name_mismatch": 0,
            "pending_approval_project_mismatch": 0,
            "pending_approval_perfect": 0,
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
        self.SOW_STATUS_COL = "Work Order Status"

        # Valid Work Order Status values to process
        self.VALID_STATUS = ["Accepted", "Activated", "Confirmed", "Pending Approval"]

    @staticmethod
    def normalize_name(name: str) -> str:
        """
        Normalize names for comparison.
        Extracts ONLY first and last names, ignoring middle names.

        Handles:
        - lastname, firstname format (SOW)
        - firstname lastname format (Base)
        - Case insensitive
        - Extra spaces

        Example:
            'Deer, John Michael' → 'deer john'
            'John Michael Deer' → 'john deer'
            Both return comparable format with first & last names only

        Args:
            name: Input name string

        Returns:
            Normalized name string with first and last names only
        """
        if pd.isna(name):
            return ""

        name = str(name).strip().lower()

        # Handle comma-separated format (Last, First Middle)
        if "," in name:
            # Split by comma: "Lastname, Firstname Middle" → ["Lastname", "Firstname Middle"]
            parts = name.split(",")
            lastname = parts[0].strip()

            # Get first name from the remaining part
            firstname_part = parts[1].strip() if len(parts) > 1 else ""
            firstname = firstname_part.split()[0] if firstname_part else ""

            # Return: firstname lastname (sorted for consistent matching)
            names = [firstname, lastname]
        else:
            # Split words: "Firstname Middle Lastname"
            words = name.split()

            if len(words) == 1:
                # Only one name
                names = words
            elif len(words) == 2:
                # Two names: first and last
                names = words
            else:
                # Multiple names: take first and last only
                names = [words[0], words[-1]]

        # Filter empty strings and sort
        names = [n for n in names if n]
        names.sort()

        return " ".join(names)

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

    def _filter_sow_data(self) -> None:
        """
        Filter SOW data:
        1. Remove rows with invalid Work Order Status (keep only: Accepted, Activated, Confirmed, Pending Approval)
        2. Filter to only blank Employee IDs

        PERMANENTLY removes invalid status rows from self.sow_df
        """
        # Filter 1: Work Order Status - REMOVE invalid status rows permanently
        if self.SOW_STATUS_COL in self.sow_df.columns:
            mask_status = self.sow_df[self.SOW_STATUS_COL].isin(self.VALID_STATUS)
            removed_invalid_status = len(self.sow_df) - len(self.sow_df[mask_status])

            # Permanently remove rows with invalid status
            self.sow_df = self.sow_df[mask_status].copy()
            self.stats["work_order_status_filtered"] = removed_invalid_status
        else:
            self.stats["work_order_status_filtered"] = 0

        # Update total records after removing invalid status rows
        self.stats["total_records_in_sow"] = len(self.sow_df)

        # Filter 2: Blank Employee IDs only (for processing)
        mask_blank_emp_id = self.sow_df[self.SOW_EMP_ID_COL].apply(
            lambda x: self._is_blank(x)
        )
        self.sow_df_filtered = self.sow_df[mask_blank_emp_id].copy()
        self.stats["blank_emp_ids_count"] = len(self.sow_df_filtered)

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

        Only processes SOW records that:
        1. Have valid Work Order Status (Accepted, Activated, Confirmed, Pending Approval)
        2. Have blank Employee IDs

        Returns:
            Updated SOW dataframe

        Raises:
            Exception if dataframes are not loaded
        """
        if self.base_df is None or self.sow_df is None:
            raise Exception("Files not loaded. Call load_files() first.")

        # Filter SOW data first
        self._filter_sow_data()

        # Build lookup from base file
        self._build_lookup()

        # Reset statistics
        self._reset_stats()

        # Process each row in filtered SOW file
        for index, row in self.sow_df_filtered.iterrows():
            self._process_row(index, row)

        return self.sow_df

    def _process_row(self, index: int, row: pd.Series) -> None:
        """
        Process a single row from filtered SOW file.

        Logic:
        1. Extract name and normalize (first + last name only)
        2. Check if name exists in base file
        3. If name mismatch: fill result with "Name MisMatch"
        4. If name matched:
           - Fill Employee ID from base
           - Check project ID match
           - If project mismatch: fill result with "Project Mismatch"
           - If everything matches: fill result with "Perfect"
           - Always fill PRISM Project ID from base

        Args:
            index: Row index
            row: Row data
        """
        sow_name = row.get(self.SOW_NAME_COL)
        sow_project = row.get(self.SOW_PROJECT_COL)
        sow_status = row.get(self.SOW_STATUS_COL)

        normalized_sow_name = self.normalize_name(sow_name)

        # Track by status
        if sow_status == "Accepted":
            self.stats["accepted_total"] += 1
        elif sow_status == "Activated":
            self.stats["activated_total"] += 1
        elif sow_status == "Confirmed":
            self.stats["confirmed_total"] += 1
        elif sow_status == "Pending Approval":
            self.stats["pending_approval_total"] += 1

        # Check if name found in base lookup
        if normalized_sow_name not in self.base_lookup:
            self.sow_df.at[index, self.SOW_RESULT_COL] = "Name MisMatch"
            self.stats["name_mismatches"] += 1

            # Track by status
            if sow_status == "Accepted":
                self.stats["accepted_name_mismatch"] += 1
            elif sow_status == "Activated":
                self.stats["activated_name_mismatch"] += 1
            elif sow_status == "Confirmed":
                self.stats["confirmed_name_mismatch"] += 1
            elif sow_status == "Pending Approval":
                self.stats["pending_approval_name_mismatch"] += 1
            return

        # Name found - extract matched data
        matched_data = self.base_lookup[normalized_sow_name]
        base_emp_id = matched_data["employee_number"]
        base_project = matched_data["project_id"]

        # Fill Employee ID from Base (since we filtered for blank IDs only)
        self.sow_df.at[index, self.SOW_EMP_ID_COL] = base_emp_id
        self.stats["updated_emp_ids"] += 1

        # Check if SOW Project ID is empty/blank
        sow_project_is_blank = self._is_blank(sow_project)

        if sow_project_is_blank:
            # SOW Project ID is empty - NOT a perfect match
            # Don't fill PRISM ID or Result
            self.stats["project_mismatches"] += 1

            # Track by status
            if sow_status == "Accepted":
                self.stats["accepted_project_mismatch"] += 1
            elif sow_status == "Activated":
                self.stats["activated_project_mismatch"] += 1
            elif sow_status == "Confirmed":
                self.stats["confirmed_project_mismatch"] += 1
            elif sow_status == "Pending Approval":
                self.stats["pending_approval_project_mismatch"] += 1
        else:
            # SOW Project ID is not empty - compare with Base Project ID
            # Normalize both project IDs: strip spaces, convert to string, and convert to int if numeric
            sow_project_str = str(sow_project).strip()
            base_project_str = str(base_project).strip()

            # Try to convert to int for numeric comparison (handles "117315" vs 117315)
            try:
                sow_project_int = int(float(sow_project_str))
                base_project_int = int(float(base_project_str))
                sow_project_normalized = str(sow_project_int)
                base_project_normalized = str(base_project_int)
            except (ValueError, TypeError):
                # If conversion fails, use string comparison
                sow_project_normalized = sow_project_str
                base_project_normalized = base_project_str

            if sow_project_normalized == base_project_normalized:
                # Project IDs match exactly - PERFECT
                # Fill Result with "Emp id updated"
                self.sow_df.at[index, self.SOW_RESULT_COL] = "Emp ID updated"
                self.stats["perfect_matches"] += 1

                # Track by status
                if sow_status == "Accepted":
                    self.stats["accepted_perfect"] += 1
                elif sow_status == "Activated":
                    self.stats["activated_perfect"] += 1
                elif sow_status == "Confirmed":
                    self.stats["confirmed_perfect"] += 1
                elif sow_status == "Pending Approval":
                    self.stats["pending_approval_perfect"] += 1
            else:
                # Project IDs don't match - PROJECT MISMATCH
                # Fill PRISM ID with base project and mark as "Emp id updated and project mismatch"
                self.sow_df.at[index, self.SOW_RESULT_COL] = (
                    "Emp ID updated and project mismatch"
                )
                self.sow_df.at[index, self.SOW_PRISM_PROJECT_COL] = base_project
                self.stats["project_mismatches"] += 1

                # Track by status
                if sow_status == "Accepted":
                    self.stats["accepted_project_mismatch"] += 1
                elif sow_status == "Activated":
                    self.stats["activated_project_mismatch"] += 1
                elif sow_status == "Confirmed":
                    self.stats["confirmed_project_mismatch"] += 1
                elif sow_status == "Pending Approval":
                    self.stats["pending_approval_project_mismatch"] += 1

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
        """Reset statistics counters for comparison run"""
        # Reset only comparison-related stats, keep filter stats
        self.stats["perfect_matches"] = 0
        self.stats["project_mismatches"] = 0
        self.stats["name_mismatches"] = 0
        self.stats["updated_emp_ids"] = 0
        # Reset status-wise breakdown
        self.stats["accepted_total"] = 0
        self.stats["accepted_name_mismatch"] = 0
        self.stats["accepted_project_mismatch"] = 0
        self.stats["accepted_perfect"] = 0
        self.stats["activated_total"] = 0
        self.stats["activated_name_mismatch"] = 0
        self.stats["activated_project_mismatch"] = 0
        self.stats["activated_perfect"] = 0
        self.stats["confirmed_total"] = 0
        self.stats["confirmed_name_mismatch"] = 0
        self.stats["confirmed_project_mismatch"] = 0
        self.stats["confirmed_perfect"] = 0
        self.stats["pending_approval_total"] = 0
        self.stats["pending_approval_name_mismatch"] = 0
        self.stats["pending_approval_project_mismatch"] = 0
        self.stats["pending_approval_perfect"] = 0

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
