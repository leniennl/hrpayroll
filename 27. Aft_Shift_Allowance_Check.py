# 20.09.2023

"""Calculate Afternoon Shift Allowance as per EBA using Kronos's DetailedHours report"""

import sys
import os
from tkinter import Tk  # py3k
import pandas as pd


def extract_float_from_schedule(schedule_str):
    """turn 7.6 hours into 7.6"""
    try:
        return float(schedule_str.split()[0])
    except ValueError:
        return schedule_str


def entitled_duration_calc(row):
    """calculate entitled duration, max 7.6 hours"""
    if isinstance(row["Schedule"], float):
        if row["Hours"] - 0.5 < row["Schedule"]:
            return row["Hours"] - 0.5
        else:
            return row["Schedule"]
    else:
        # Schedule value is 'Not Scheduled', indicating an exra shift
        return 0


# get path from clipboard
FilepathCB = Tk().selection_get(selection="CLIPBOARD")
if not os.path.isdir(FilepathCB):
    print("copy a valid path to the clipboard")
    sys.exit()
counter = 0
counter2 = 0
for filefolders, subfolders, filenames in os.walk(FilepathCB):
    for filename in filenames:
        if str(filename).startswith("DetailedHours"):
            timesheetfilepath = str(filefolders + "/" + filename)
            counter += 1
        if str(filename).upper().startswith("PAY RATE"):
            ratefilepath = str(filefolders + "/" + filename)
            counter2 += 1
if counter == 0:
    print("Missing DetailedHours_ in this folder.")
    sys.exit()
elif counter2 == 0:
    print("Missing pay rate file in this folder.")
    sys.exit()
elif counter >= 2 or counter2 >= 2:
    print("More than 1 file found: DetailedHours AND Pay Rate")
    sys.exit()


DETAILED_REPORT_DF = pd.read_excel(
    timesheetfilepath, skiprows=7, sheet_name="report", skipfooter=3
)
DETAILED_REPORT_DF["Employee Id"] = (
    DETAILED_REPORT_DF["Employee Id"].str.extract(r"(\d+)").astype(int)
)  # turn SILLKER19801  to 19801

PAYRATE_DF = pd.read_csv(ratefilepath, usecols=["Employee code", "Trans Rate"])
PAYRATE_DF.rename(
    columns={"Employee code": "Employee Id", "Trans Rate": "Hourly Rate"}, inplace=True
)

DESIGNATED_EMP_LIST = pd.read_excel(
    r"J:\Secure\ADP NEW\Afternoon shift.xlsx", usecols=["Employee Id"]
)
DESIGNATED_EMP_LIST = DESIGNATED_EMP_LIST.merge(
    PAYRATE_DF, on="Employee Id", how="left"
)

shift_worker_actual_working_df = DETAILED_REPORT_DF.merge(
    DESIGNATED_EMP_LIST, on="Employee Id", how="inner"
)
shift_worker_actual_working_df["Schedule"] = shift_worker_actual_working_df[
    "Schedule"
].apply(extract_float_from_schedule)
shift_worker_actual_working_df[
    "Aft Shift Duration"
] = shift_worker_actual_working_df.apply(entitled_duration_calc, axis=1)

eligible_worker_eligible_shift_df = shift_worker_actual_working_df[
    (
        shift_worker_actual_working_df["End"] > "18:06"
    )  # not a shift finishing before 18:06
    & (shift_worker_actual_working_df["Time Off\nName"].isna())  # not a Timeoff shift
    & (
        ~shift_worker_actual_working_df["Date"].apply(
            lambda x: x.weekday() in [5, 6]
        )  # not a weekend shift
    )
]
eligible_worker_eligible_shift_df = eligible_worker_eligible_shift_df.copy()
eligible_worker_eligible_shift_df["Allowance"] = (
    eligible_worker_eligible_shift_df["Hourly Rate"]
    * eligible_worker_eligible_shift_df["Aft Shift Duration"]
    * 0.15
).round(2)

pivot_table = pd.pivot_table(
    eligible_worker_eligible_shift_df,
    index=["Employee Id", "Last Name", "First Name"],
    values=["Allowance"],
    aggfunc="sum",
)
pivot_table = pivot_table.reset_index()
pivot_table.insert(
    pivot_table.columns.get_loc("Allowance"),
    "Allowance / Deduction / Element Code",
    "A044",
)
pivot_table.insert(pivot_table.columns.get_loc("Allowance"), "Hours", "")
pivot_table.insert(pivot_table.columns.get_loc("Allowance") + 1, "From", "")
pivot_table.insert(pivot_table.columns.get_loc("Allowance") + 2, "To", "")
pivot_table.insert(
    pivot_table.columns.get_loc("Allowance") + 3,
    "Comment",
    "change allowance as per actual",
)


with pd.ExcelWriter(FilepathCB + "/Aft Allowance Check.xlsx") as writer:
    try:
        shift_worker_actual_working_df.to_excel(
            writer, sheet_name="Eligible Worker", index=False
        )
        eligible_worker_eligible_shift_df.to_excel(
            writer, sheet_name="Eligible Worker Shift", index=False
        )
        pivot_table.to_excel(writer, sheet_name="Pivot_Table", index=False)
        print("Aft Allowance Check.xlsx saved at " + FilepathCB)
    except PermissionError:
        print("File in use. Close the file and try again.")
        sys.exit()
