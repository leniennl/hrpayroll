import pandas as pd
import warnings
import sys, os
from tkinter import Tk  # py3k

# get path from clipboard
Filepath = Tk().selection_get(selection="CLIPBOARD")
if not os.path.isdir(Filepath):
    print("copy a valid path to the clipboard")
    sys.exit()
counter = 0
counter2 = 0
for filefolders, subfolders, filenames in os.walk(Filepath):
    for filename in filenames:
        if str(filename).startswith("DetailedCalculatedTimeCounters-System"):
            timesheetfilepath = str(filefolders + "/" + filename)
            counter += 1
        if str(filename).startswith("PAY RATE"):
            ratefilepath = str(filefolders + "/" + filename)
            counter2 += 1
if counter == 0 or counter2 == 0:
    print("Missing file(s) in this folder.")
    sys.exit()
elif counter >= 2 or counter2 >= 2:
    print(str(counter) + " files found! Delete all but one.")
    sys.exit()

# Rea in constants
MULTIPLIER = pd.read_excel(
    r"J:\Secure\ADP Payroll Report\2024\FN 230704\MSNX Counter Multiplier.xlsx"
)
HOURLY_RATE = pd.read_csv(
    ratefilepath, usecols=["Employee code", "Trans Rate"]
)  # use latest rate export
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
timesheets = pd.read_excel(timesheetfilepath, skiprows=7)

# Assuming your DataFrame is named 'df'
grouped_df = (
    timesheets.groupby(["Employee Id", "Counter Name"])["Counter Hours"]
    .sum()
    .reset_index()
)

# add hours multiplier to grouped data
merged_df = grouped_df.merge(MULTIPLIER, on="Counter Name", how="left")
# modify Employee Id for matching with rate
merged_df["Employee code"] = (
    merged_df["Employee Id"].str.replace("SILLK", "").astype("int64")
)

# add in rate as per EMP ID
merged_rate = merged_df.merge(HOURLY_RATE, on="Employee code", how="left")
# add gross. Calculate
merged_rate["Pay Elements Total from Kronos"] = (
    merged_rate["Counter Hours"] * merged_rate["Multiplier"] * merged_rate["Trans Rate"]
)

# Calculate Pay Element Total
gross = merged_rate.groupby("Employee code")["Pay Elements Total from Kronos"].sum().reset_index().round(2)


# Create an Excel writer object
with pd.ExcelWriter(
    Filepath + "/Kronos Timesheets Pay Element Check.xlsx"
) as writer:
    # Write each DataFrame to a different sheet in the Excel file
    MULTIPLIER.to_excel(writer, sheet_name="Multiplier",index=False)
    timesheets.to_excel(writer, sheet_name="SourceData",index=False)
    grouped_df.to_excel(writer, sheet_name="EMPCounterTotal",index=False)
    merged_df.to_excel(writer, sheet_name="AddMultipler",index=False)
    merged_rate.to_excel(writer, sheet_name="AddRate",index=False)
    gross.to_excel(writer, sheet_name="PayElementTotal",index=False)
    
print("File generated.")
