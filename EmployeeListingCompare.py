import pandas as pd
import numpy as np
import sys
from pathlib import Path
from datetime import datetime
import re



# ============================================================
# SETTINGS
# ============================================================

SHEET_NAME = "Export"

KEY_COLUMN = "EmployeeId"
FIRST_NAME_COLUMN = "FirstName"
SURNAME_COLUMN = "Surname"
PAY_SCHEDULE_COLUMN = "PaySchedule"


# ============================================================
# FUNCTIONS
# ============================================================

def load_employee_file(file_path):
    """
    Read the Export sheet from an employee Excel export
    and prepare the dataframe.
    """

    df = pd.read_excel(
        file_path,
        sheet_name=SHEET_NAME,
        header=0
    )

    # Check required columns exist
    required_columns = [
        KEY_COLUMN,
        FIRST_NAME_COLUMN,
        SURNAME_COLUMN,
        PAY_SCHEDULE_COLUMN
    ]

    missing_columns = [
        col for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            f"The following required columns are missing from "
            f"{file_path.name}: {missing_columns}"
        )

    # Create EmployeeName
    df["EmployeeName"] = (
        df[FIRST_NAME_COLUMN]
        .fillna("")
        .astype(str)
        .str.strip()
        + " "
        + df[SURNAME_COLUMN]
        .fillna("")
        .astype(str)
        .str.strip()
    ).str.strip()

    # Check EmployeeId
    if df[KEY_COLUMN].isna().any():
        raise ValueError(
            f"{file_path.name} contains employees with no EmployeeId."
        )

    if df[KEY_COLUMN].duplicated().any():
        duplicates = (
            df.loc[
                df[KEY_COLUMN].duplicated(keep=False),
                KEY_COLUMN
            ]
            .unique()
            .tolist()
        )

        raise ValueError(
            f"{file_path.name} contains duplicate EmployeeId values: "
            f"{duplicates}"
        )

    return df


def compare_employee_files(df_old, df_new):
    """
    Compare two employee dataframes.

    Returns:
        df_changed
        df_added
        df_removed
    """

    # --------------------------------------------------------
    # Employee IDs
    # --------------------------------------------------------

    old_ids = set(df_old[KEY_COLUMN])
    new_ids = set(df_new[KEY_COLUMN])

    added_ids = new_ids - old_ids
    removed_ids = old_ids - new_ids
    common_ids = old_ids & new_ids

    # --------------------------------------------------------
    # Compare common employees
    # --------------------------------------------------------

    df_old_common = (
        df_old[
            df_old[KEY_COLUMN].isin(common_ids)
        ]
        .copy()
        .set_index(KEY_COLUMN)
    )

    df_new_common = (
        df_new[
            df_new[KEY_COLUMN].isin(common_ids)
        ]
        .copy()
        .set_index(KEY_COLUMN)
    )

    # Only compare columns that exist in both files
    common_columns = df_old_common.columns.intersection(
        df_new_common.columns
    )

    df_old_compare = df_old_common[common_columns]
    df_new_compare = df_new_common[common_columns]

    # --------------------------------------------------------
    # Find changed cells
    # --------------------------------------------------------

    differences = []

    for employee_id in common_ids:

        old_row = df_old_compare.loc[employee_id]
        new_row = df_new_compare.loc[employee_id]

        # PaySchedule is used as an identifying/filtering field,
        # but it can also be a changed field.
        pay_schedule = new_row[PAY_SCHEDULE_COLUMN]

        for column in common_columns:

            old_value = old_row[column]
            new_value = new_row[column]

            # Both blank = no difference
            if pd.isna(old_value) and pd.isna(new_value):
                continue

            # One blank, one populated = difference
            if pd.isna(old_value) != pd.isna(new_value):

                differences.append({
                    "Change Type": "Changed",
                    "EmployeeId": employee_id,
                    "EmployeeName": new_row["EmployeeName"],
                    "PaySchedule": pay_schedule,
                    "Column": column,
                    "Old Value": old_value,
                    "New Value": new_value
                })

                continue

            # Both populated and different
            if old_value != new_value:

                differences.append({
                    "Change Type": "Changed",
                    "EmployeeId": employee_id,
                    "EmployeeName": new_row["EmployeeName"],
                    "PaySchedule": pay_schedule,
                    "Column": column,
                    "Old Value": old_value,
                    "New Value": new_value
                })

    df_changed = pd.DataFrame(
        differences,
        columns=[
            "Change Type",
            "EmployeeId",
            "EmployeeName",
            "PaySchedule",
            "Column",
            "Old Value",
            "New Value"
        ]
    )

    # --------------------------------------------------------
    # Added employees
    # --------------------------------------------------------

    df_added_source = df_new[
        df_new[KEY_COLUMN].isin(added_ids)
    ].copy()

    added_records = []

    for _, row in df_added_source.iterrows():

        added_records.append({
            "Change Type": "Added",
            "EmployeeId": row[KEY_COLUMN],
            "EmployeeName": row["EmployeeName"],
            "PaySchedule": row[PAY_SCHEDULE_COLUMN],
            "Column": "",
            "Old Value": np.nan,
            "New Value": "New Employee"
        })

    df_added = pd.DataFrame(
        added_records,
        columns=[
            "Change Type",
            "EmployeeId",
            "EmployeeName",
            "PaySchedule",
            "Column",
            "Old Value",
            "New Value"
        ]
    )

    # --------------------------------------------------------
    # Removed employees
    # --------------------------------------------------------

    df_removed_source = df_old[
        df_old[KEY_COLUMN].isin(removed_ids)
    ].copy()

    removed_records = []

    for _, row in df_removed_source.iterrows():

        removed_records.append({
            "Change Type": "Removed",
            "EmployeeId": row[KEY_COLUMN],
            "EmployeeName": row["EmployeeName"],
            "PaySchedule": row[PAY_SCHEDULE_COLUMN],
            "Column": "",
            "Old Value": "Existing Employee",
            "New Value": np.nan
        })

    df_removed = pd.DataFrame(
        removed_records,
        columns=[
            "Change Type",
            "EmployeeId",
            "EmployeeName",
            "PaySchedule",
            "Column",
            "Old Value",
            "New Value"
        ]
    )

    return df_changed, df_added, df_removed


def format_output(writer, sheet_name, dataframe):
    """
    Write dataframe to Excel and apply basic formatting.
    """

    dataframe.to_excel(
        writer,
        sheet_name=sheet_name,
        index=False
    )

    worksheet = writer.book[sheet_name]

    # Freeze header row
    worksheet.freeze_panes = "A2"

    # Add Excel autofilter
    if len(dataframe.columns) > 0:
        worksheet.auto_filter.ref = worksheet.dimensions

    # Automatically adjust column widths
    for column_cells in worksheet.columns:

        max_length = 0
        column_letter = column_cells[0].column_letter

        for cell in column_cells:

            try:
                cell_length = len(str(cell.value))
                max_length = max(max_length, cell_length)
            except Exception:
                pass

        # Keep widths sensible
        adjusted_width = min(max(max_length + 2, 10), 40)

        worksheet.column_dimensions[
            column_letter
        ].width = adjusted_width



def extract_timestamp(file_path):
    """
    Extract the YYYYMMDDHHMMSS timestamp from the filename.

    Example:
    PremosoPtyLtd_EmployeeData_20260824101458.xlsx

    returns:
    2026-08-24 10:14:58
    """

    filename = Path(file_path).name

    match = re.search(
        r"(\d{14})",
        filename
    )

    if not match:
        raise ValueError(
            f"Could not find a 14-digit timestamp in filename:\n"
            f"{filename}\n\n"
            f"Expected something like:\n"
            f"PremosoPtyLtd_EmployeeData_20260824101458.xlsx"
        )

    timestamp_string = match.group(1)

    try:
        return datetime.strptime(
            timestamp_string,
            "%Y%m%d%H%M%S"
        )

    except ValueError:
        raise ValueError(
            f"Invalid timestamp in filename:\n{filename}"
        )

# ============================================================
# MAIN PROGRAM
# ============================================================

def main():

    # --------------------------------------------------------
    # Check command-line arguments
    # --------------------------------------------------------

    if len(sys.argv) != 3:

        print()
        print("Usage:")
        print(
            "py employeelistingcompare.py "
            "\"FILE1.xlsx\" \"FILE2.xlsx\""
        )
        print()
        print(
            "The program will automatically determine "
            "which file is older and which is newer "
            "from the timestamp in the filename."
        )
        print()

        sys.exit(1)

    # --------------------------------------------------------
    # Get file paths
    # --------------------------------------------------------

    
    file1 = Path(sys.argv[1])
    file2 = Path(sys.argv[2])
   

    # --------------------------------------------------------
    # Check files exist
    # --------------------------------------------------------

    if not file1.exists():

        print()
        print(f"ERROR: File not found:")
        print(file1)
        print()

        sys.exit(1)


    if not file2.exists():

        print()
        print(f"ERROR: File not found:")
        print(file2)
        print()

        sys.exit(1)


    # --------------------------------------------------------
    # Determine which file is OLD and which is NEW
    # --------------------------------------------------------

    timestamp1 = extract_timestamp(file1)
    timestamp2 = extract_timestamp(file2)


    if timestamp1 < timestamp2:

        old_file = file1
        new_file = file2

    elif timestamp2 < timestamp1:

        old_file = file2
        new_file = file1

    else:

        print()
        print("ERROR: Both files have the same timestamp.")
        print()
        print(f"File 1: {file1.name}")
        print(f"File 2: {file2.name}")
        print()

        sys.exit(1)     

    # --------------------------------------------------------
    # Display information
    # --------------------------------------------------------

    print()
    print("=" * 60)
    print("EMPLOYEE FILE COMPARISON")
    print("=" * 60)

    print(f"OLD file: {old_file.name}")
    print(f"     Timestamp: {timestamp1 if old_file == file1 else timestamp2}")

    print()

    print(f"NEW file: {new_file.name}")
    print(f"     Timestamp: {timestamp2 if new_file == file2 else timestamp1}")

    print()

    # --------------------------------------------------------
    # Load files
    # --------------------------------------------------------

    print("Reading old employee file...")
    df_old = load_employee_file(old_file)

    print("Reading new employee file...")
    df_new = load_employee_file(new_file)

    # --------------------------------------------------------
    # Compare
    # --------------------------------------------------------

    print("Comparing employee data...")

    df_changed, df_added, df_removed = compare_employee_files(
        df_old,
        df_new
    )

    # --------------------------------------------------------
    # Combine all differences
    # --------------------------------------------------------

    df_all_differences = pd.concat(
        [
            df_added,
            df_removed,
            df_changed
        ],
        ignore_index=True
    )

    # Sort
    if not df_all_differences.empty:

        df_all_differences = (
            df_all_differences
            .sort_values(
                ["EmployeeId", "Column"],
                na_position="first"
            )
            .reset_index(drop=True)
        )

    # --------------------------------------------------------
    # Output filename
    # --------------------------------------------------------

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    output_file = new_file.parent / (
        f"EmployeeComparison_{timestamp}.xlsx"
    )

    # --------------------------------------------------------
    # Write Excel file
    # --------------------------------------------------------

    print("Creating Excel report...")

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        format_output(
            writer,
            "All Differences",
            df_all_differences
        )

        format_output(
            writer,
            "Changed Cells",
            df_changed
        )

        format_output(
            writer,
            "Added Employees",
            df_added
        )

        format_output(
            writer,
            "Removed Employees",
            df_removed
        )

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    print()
    print("=" * 60)
    print("COMPARISON COMPLETE")
    print("=" * 60)

    print(f"Old employees : {df_old[KEY_COLUMN].nunique():,}")
    print(f"New employees : {df_new[KEY_COLUMN].nunique():,}")
    print(f"Added         : {len(df_added):,}")
    print(f"Removed       : {len(df_removed):,}")
    print(f"Changed cells : {len(df_changed):,}")
    print(
        f"Employees affected: "
        f"{df_all_differences[KEY_COLUMN].nunique():,}"
    )

    print()
    print("Output file:")
    print(output_file)
    print()

    print("=" * 60)


if __name__ == "__main__":
    main()





