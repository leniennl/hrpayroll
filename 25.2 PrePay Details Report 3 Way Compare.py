import pandas as pd
import zipfile
import sys, os
from tkinter import Tk

# get path from clipboard
Filepath = Tk().selection_get(selection="CLIPBOARD")
if not os.path.isdir(Filepath):
    print("copy a valid path to the clipboard")
    sys.exit()
counter = 0
counter2 = 0
counter3 = 0
for filefolders, subfolders, filenames in os.walk(Filepath):
    for filename in filenames:
        if str(filename).startswith("01APayPreparation_CSV_") and str(
            filename
        ).endswith(".zip"):
            prepay_zip_filepath = str(filefolders + "/" + filename)
            counter += 1
        if str(filename).startswith("Kronos Timesheets Pay Element Check"):
            kronos_filepath = str(filefolders + "/" + filename)
            counter2 += 1
        if str(filename).startswith("MXNS Salaried Employee Earnings Check"):
            salaried_employee_filepath = str(filefolders + "/" + filename)
            counter3 += 1
if counter == 0 or counter2 == 0 or counter3 == 0:
    print("Missing file(s) in this folder.")
    sys.exit()
elif counter >= 2 or counter2 >= 2 or counter3 >= 2:
    print(str(counter) + " files found! Delete all but one.")
    sys.exit()


# Function to read the CSV file from the zip archive
def read_csv_from_zip(zip_file_path, csv_file_name):
    try:
        # Open the zip file in read mode
        with zipfile.ZipFile(zip_file_path, "r") as zip_file:
            # Find the CSV file inside the zip archive
            for file_name in zip_file.namelist():
                if csv_file_name in file_name:
                    df = pd.read_csv(zip_file.open(file_name))
                    return df
    except Exception as e:
        print(f"Error: {e}")
        return None


# Call the function to read the CSV file into a DataFrame
dataframe = read_csv_from_zip(prepay_zip_filepath, "PAYROLL_DETAILS_PE_")
if dataframe is not None:
    dataframe.fillna(0, inplace=True)
    dataframe["Pay Elem Total from PrePay Reports"] = (
        dataframe["Normal"]
        + dataframe["Casual Normal"]
        + dataframe["Saturday 150%"]
        + dataframe["Sunday 200%"]
        + dataframe["O/Time 1.5"]
        + dataframe["O/Time 2.0"]
        + dataframe["Annual Lve"]
        + dataframe["Sick/Personal"]
        + dataframe["LS Leave"]
        + dataframe["Carers Lve"]
        + dataframe["LWOP"]
    )
    pay_ele_prepay_df = dataframe[["Emp Code", "Pay Elem Total from PrePay Reports"]]
    pay_ele_prepay_df = pay_ele_prepay_df.copy()
    pay_ele_prepay_df.rename(columns={"Emp Code": "Employee code"}, inplace=True)

    pay_ele_kronos_df = pd.read_excel(
        kronos_filepath, sheet_name="PayElementTotal"
    )  # read in Kronos reports and compare
    compare_df = pay_ele_prepay_df.merge(
        pay_ele_kronos_df, on="Employee code", how="left"
    )
    salaried_df = pd.read_excel(
        salaried_employee_filepath,
        sheet_name="Employee Listing _Salaried",
        usecols=["Employee code", "Fortnightly Salary "],
    )  # read in from salaried employee files and compare
    three_way_compare_df = compare_df.merge(salaried_df, on="Employee code", how="left")
    three_way_compare_df["Pay Elements Total from Kronos"].fillna(0, inplace=True)
    three_way_compare_df["Fortnightly Salary "].fillna(0, inplace=True)
    three_way_compare_df["Diff"] = (
        three_way_compare_df["Pay Elem Total from PrePay Reports"]
        - three_way_compare_df["Pay Elements Total from Kronos"]
        - three_way_compare_df["Fortnightly Salary "]
    )
    three_way_compare_df = three_way_compare_df.round(2)

    # Write the new DataFrame to the existing Excel file using openpyxl
    with pd.ExcelWriter(Filepath + "/PrePay_Details_3_Way_Compare.xlsx") as writer:
        three_way_compare_df.to_excel(
            writer, sheet_name="three_way_compare", index=False
        )
    print("file generated in clipboard file path: PrePay_Details_3_Way_Compare.xlsx")

else:
    print("Failed to read the CSV file from the zip archive.")
