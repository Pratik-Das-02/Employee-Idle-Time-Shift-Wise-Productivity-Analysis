# =====================================================================
# EMPLOYEE IDLE TIME & SHIFT-WISE PRODUCTIVITY ANALYSIS
# Reads an Excel submission history file and creates an employee-wise
# idle-time report arranged by Shift A to Shift G.
#
# Designed to run in Google Colab.
# =====================================================================

import pandas as pd
import numpy as np
from google.colab import files


# =====================================================================
# 1. SETTINGS
# =====================================================================

SHEET_NAME = "Submissions History"

SHIFT_DURATION = pd.Timedelta(hours=9)
BREAK_TIME = pd.Timedelta(hours=1)

SHIFT_ORDER = ["A", "B", "C", "D", "E", "F", "G"]


# =====================================================================
# 2. UPLOAD EXCEL FILE
# =====================================================================

print("=" * 80)
print("UPLOAD YOUR EXCEL FILE")
print("=" * 80)

uploaded = files.upload()

if not uploaded:
    raise ValueError("Please upload an Excel file.")

INPUT_FILE = list(uploaded.keys())[0]

print("\nUploaded file:", INPUT_FILE)


# =====================================================================
# 3. READ EXCEL
# =====================================================================

df = pd.read_excel(
    INPUT_FILE,
    sheet_name=SHEET_NAME
)

# Standardize column names
df.columns = (
    df.columns.astype(str)
    .str.replace("\xa0", " ", regex=False)
    .str.replace("\n", " ", regex=False)
    .str.strip()
    .str.replace(r"\s+", " ", regex=True)
)

print("\nTotal Excel rows:", len(df))


# =====================================================================
# 4. IDENTIFY REQUIRED COLUMNS
# =====================================================================

def normalize_column_name(name):
    return (
        str(name)
        .lower()
        .replace(" ", "")
        .replace("_", "")
        .strip()
    )


column_lookup = {
    normalize_column_name(col): col
    for col in df.columns
}

required_aliases = {
    "Assigned To": ["assignedto"],
    "Actual SLA": ["actualsla"]
}

rename_map = {}

for standard_name, aliases in required_aliases.items():
    found_column = None

    for alias in aliases:
        if alias in column_lookup:
            found_column = column_lookup[alias]
            break

    if found_column is None:
        raise ValueError(
            f"Missing column: {standard_name}\n"
            f"Available columns: {df.columns.tolist()}"
        )

    rename_map[found_column] = standard_name

df = df.rename(columns=rename_map)

print("Required columns found successfully.")


# =====================================================================
# 5. CLEAN EMPLOYEE NAMES
# =====================================================================

df["Assigned To"] = (
    df["Assigned To"]
    .fillna("")
    .astype(str)
    .str.replace("\xa0", " ", regex=False)
    .str.strip()
    .str.replace(r"\s+", " ", regex=True)
)


# =====================================================================
# 6. CONVERT ACTUAL SLA
# =====================================================================

def parse_sla(value):
    """Convert Excel time values or duration strings to Timedelta."""
    if pd.isna(value):
        return pd.NaT

    # Excel stores time/duration values as a fraction of one day.
    if isinstance(value, (int, float, np.integer, np.floating)):
        return pd.to_timedelta(float(value), unit="D")

    value = str(value).strip()

    if value.lower() in ["", "nan", "nat", "none"]:
        return pd.NaT

    try:
        return pd.to_timedelta(value)
    except (ValueError, TypeError, OverflowError):
        return pd.NaT


df["Actual SLA Duration"] = df["Actual SLA"].apply(parse_sla)


# =====================================================================
# 7. SEPARATE INVALID ROWS
# =====================================================================

valid_mask = (
    df["Assigned To"].ne("")
    & df["Actual SLA Duration"].notna()
    & df["Actual SLA Duration"].ge(pd.Timedelta(0))
)

valid_data = df.loc[valid_mask].copy()
invalid_rows = df.loc[~valid_mask].copy()

if valid_data.empty:
    raise ValueError("No valid employee records found.")

print("Valid employee records:", len(valid_data))
print("Invalid or blank records:", len(invalid_rows))


# =====================================================================
# 8. PERSON-WISE CALCULATION
# =====================================================================

summary = (
    valid_data
    .groupby("Assigned To", as_index=False)
    .agg(
        COUNT=("Actual SLA Duration", "count"),
        TOTAL_ACTUAL_SLA=("Actual SLA Duration", "sum")
    )
)


# =====================================================================
# 9. CALCULATE IDLE TIME
# =====================================================================

summary["Shift Duration"] = SHIFT_DURATION
summary["Break Time"] = BREAK_TIME

summary["Idle Time Duration"] = (
    summary["Shift Duration"]
    - summary["Break Time"]
    - summary["TOTAL_ACTUAL_SLA"]
)

# Idle time cannot be negative.
summary["Idle Time Duration"] = (
    summary["Idle Time Duration"]
    .clip(lower=pd.Timedelta(0))
)


# =====================================================================
# 10. FORMAT TIME AS HH:MM:SS
# =====================================================================

def format_time(value):
    if pd.isna(value):
        return "00:00:00"

    total_seconds = max(0, int(value.total_seconds()))

    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60

    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


summary["Total Actual SLA"] = (
    summary["TOTAL_ACTUAL_SLA"].apply(format_time)
)

summary["Shift Duration"] = (
    summary["Shift Duration"].apply(format_time)
)

summary["Break Time"] = (
    summary["Break Time"].apply(format_time)
)

summary["Idle Time"] = (
    summary["Idle Time Duration"].apply(format_time)
)


# =====================================================================
# 11. FINAL REPORT - SHIFT-WISE ARRANGEMENT
# =====================================================================

final_report = summary[
    [
        "Assigned To",
        "COUNT",
        "Idle Time"
    ]
].copy()

final_report = final_report.rename(
    columns={"Assigned To": "NAME"}
)

# Extract shift letter from employee name.
final_report["Shift"] = (
    final_report["NAME"]
    .str.extract(r"Shift\s+([A-G])\s*$", expand=False)
)

# Arrange employees by Shift A to G.
final_report["Shift Sort"] = pd.Categorical(
    final_report["Shift"],
    categories=SHIFT_ORDER,
    ordered=True
)

final_report = (
    final_report
    .sort_values(
        ["Shift Sort", "NAME"],
        na_position="last"
    )
    .drop(columns=["Shift Sort"])
    .reset_index(drop=True)
)

# Put Shift column first.
final_report = final_report[
    ["Shift", "NAME", "COUNT", "Idle Time"]
]


# =====================================================================
# 12. DETAILED CALCULATION SHEET
# =====================================================================

detailed_report = summary[
    [
        "Assigned To",
        "COUNT",
        "Shift Duration",
        "Total Actual SLA",
        "Break Time",
        "Idle Time"
    ]
].copy()

detailed_report = detailed_report.rename(
    columns={"Assigned To": "NAME"}
)

# Extract shift.
detailed_report["Shift"] = (
    detailed_report["NAME"]
    .str.extract(r"Shift\s+([A-G])\s*$", expand=False)
)

detailed_report["Shift Sort"] = pd.Categorical(
    detailed_report["Shift"],
    categories=SHIFT_ORDER,
    ordered=True
)

detailed_report = (
    detailed_report
    .sort_values(
        ["Shift Sort", "NAME"],
        na_position="last"
    )
    .drop(columns=["Shift Sort"])
    .reset_index(drop=True)
)

detailed_report = detailed_report[
    [
        "Shift",
        "NAME",
        "COUNT",
        "Shift Duration",
        "Total Actual SLA",
        "Break Time",
        "Idle Time"
    ]
]


# =====================================================================
# 13. VALIDATION
# =====================================================================

validation = pd.DataFrame([
    {
        "Description": "Total Excel Rows",
        "Value": len(df)
    },
    {
        "Description": "Valid Employee Records",
        "Value": len(valid_data)
    },
    {
        "Description": "Invalid / Blank Records",
        "Value": len(invalid_rows)
    },
    {
        "Description": "Total Employee Groups",
        "Value": len(final_report)
    },
    {
        "Description": "Shift Duration",
        "Value": format_time(SHIFT_DURATION)
    },
    {
        "Description": "Break Per Employee",
        "Value": format_time(BREAK_TIME)
    },
    {
        "Description": "Idle Time Formula",
        "Value": "9 Hours - Total Actual SLA - 1 Hour Break"
    }
])


# =====================================================================
# 14. SAVE EXCEL REPORT
# =====================================================================

OUTPUT_FILE = "Employee_Idle_Time_Final_Report.xlsx"

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    final_report.to_excel(
        writer,
        sheet_name="Person Wise Idle Time",
        index=False
    )

    detailed_report.to_excel(
        writer,
        sheet_name="Detailed Calculation",
        index=False
    )

    validation.to_excel(
        writer,
        sheet_name="Validation",
        index=False
    )

    invalid_rows.to_excel(
        writer,
        sheet_name="Invalid Rows",
        index=False
    )


# =====================================================================
# 15. DISPLAY FINAL RESULT
# =====================================================================

print("\n")
print("=" * 80)
print("PERSON WISE IDLE TIME REPORT - SHIFT WISE")
print("=" * 80)

print(final_report.to_string(index=False))

print("\n")
print("=" * 80)
print("DETAILED CALCULATION")
print("=" * 80)

print(detailed_report.to_string(index=False))

print("\nTotal files counted:", final_report["COUNT"].sum())

print("\nShift-wise file count:")

shift_counts = (
    final_report
    .groupby("Shift", observed=False)["COUNT"]
    .sum()
    .reindex(SHIFT_ORDER, fill_value=0)
)

print(shift_counts.to_string())

print("\nExcel report created successfully:")
print(OUTPUT_FILE)


# =====================================================================
# 16. DOWNLOAD EXCEL
# =====================================================================

files.download(OUTPUT_FILE)
