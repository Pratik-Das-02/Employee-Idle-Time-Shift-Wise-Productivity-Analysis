# Employee Idle Time & Shift-Wise Productivity Analysis

---

## 1. Project Summary

This project automates employee-wise submission counting and estimated idle-time reporting from an Excel submission history workbook using Python.

The script reads the `Submissions History` worksheet, validates the required columns, cleans employee names, converts Actual SLA values into time durations, and separates valid records from invalid or incomplete records.

It then groups the valid records by employee, calculates each employee's submission count and total Actual SLA, estimates idle time using a fixed shift duration and break allowance, and arranges employees by Shift A through Shift G.

The final output is an Excel workbook containing an employee-wise summary, detailed calculations, validation information, and invalid rows.

---

## 2. Business Purpose & Calculation Logic

The purpose of this project is to reduce manual effort in preparing employee-wise productivity and idle-time reports from submission history data.

* **Employee-wise Submission Count:** Counts the valid Actual SLA records for each employee.
* **Total Actual SLA:** Adds the Actual SLA durations for each employee.
* **Shift Duration:** Uses a configured shift duration of 9 hours.
* **Break Allowance:** Deducts a configured 1-hour break from the shift duration.
* **Estimated Idle Time:** Calculates the remaining time after deducting the break allowance and total Actual SLA.

**Idle Time Formula:**

`Idle Time = Shift Duration - Break Time - Total Actual SLA`

With the current settings:

`Idle Time = 9 Hours - 1 Hour Break - Total Actual SLA`

Calculated idle time is restricted to a minimum of `00:00:00`; negative results are clipped to zero.

**Important interpretation:** This is an estimated time balance based on total Actual SLA. It does not measure the actual gaps between submissions or verify attendance, login/logout, or real break timings.

---

## 3. Data Processing & Validation

The script performs the following processing steps:

1. **Excel Upload:** Prompts the user to upload an Excel workbook in Google Colab.
2. **Worksheet Reading:** Reads the worksheet named `Submissions History`.
3. **Column Standardisation:** Cleans column headers by removing extra spaces, line breaks, and non-breaking spaces.
4. **Required Column Validation:** Checks for `Assigned To` and `Actual SLA`.
5. **Employee Name Cleaning:** Removes leading/trailing spaces and standardises repeated whitespace in employee names.
6. **Actual SLA Conversion:** Converts supported Excel time values and duration strings into Pandas time durations.
7. **Valid Record Identification:** Keeps records with a nonblank employee name and a valid, nonnegative Actual SLA duration.
8. **Invalid Record Separation:** Sends blank employee names, missing or unreadable SLA values, and negative SLA values to the invalid-record output.
9. **Employee-Level Aggregation:** Groups valid records by the exact cleaned employee name.
10. **Shift Extraction & Sorting:** Extracts shift letters A–G from names ending in formats such as `Employee Name Shift A` and sorts recognised shifts in A-to-G order.

---

## 4. Report Output & Workbook Structure

The script generates the Excel report:

`Employee_Idle_Time_Final_Report.xlsx`

The workbook contains four worksheets:

* **Person Wise Idle Time:** Displays Shift, employee name, submission count, and estimated idle time.
* **Detailed Calculation:** Displays Shift, employee name, count, shift duration, total Actual SLA, break time, and idle time.
* **Validation:** Displays total Excel rows, valid employee records, invalid or blank records, employee groups, shift duration, break allowance, and the idle-time formula.
* **Invalid Rows:** Contains the source records excluded from employee-wise calculations so that the data can be reviewed.

The script also prints the employee-wise report and shift-wise submission counts in the Google Colab output.

---

## 5. Configuration & Input Requirements

### Current Settings

| Setting | Configured value |
|---|---:|
| Input worksheet | `Submissions History` |
| Shift duration | 9 hours |
| Break time | 1 hour |
| Shift order | A, B, C, D, E, F, G |

These values can be changed in the `SETTINGS` section of the Python script.

### Required Excel Columns

| Column | Requirement |
|---|---|
| `Assigned To` | Employee name or employee/shift label |
| `Actual SLA` | Actual processing duration |

Example employee labels:

```text
Aman Shift A
Rohit Shift B
Priya Shift C
```

The script extracts a shift only when the employee name ends with `Shift A` through `Shift G`. Employees without a recognised shift suffix are still included in the report, with a blank Shift value.

---

## 6. Technologies Used

* **Python:** Main programming language.
* **Pandas:** Excel reading, data cleaning, grouping, aggregation, and report preparation.
* **NumPy:** Numeric type handling during Actual SLA conversion.
* **OpenPyXL:** Excel workbook writing through Pandas.
* **Google Colab:** File upload and generated report download.

---

## 7. How to Run the Project

The current script is designed for Google Colab.

1. Open [Google Colab](https://colab.research.google.com/).
2. Create a new notebook.
3. Upload or paste `employee_idle_time_analysis.py` into a code cell. Alternatively, open the `.py` file and copy its contents into a cell.
4. Run the code.
5. Upload an Excel workbook when prompted.
6. Ensure the workbook contains a worksheet named `Submissions History` with the required columns.
7. Wait for the processing to complete.
8. Download `Employee_Idle_Time_Final_Report.xlsx` when prompted.

### Python Dependencies

The main third-party packages are:

```bash
pip install pandas numpy openpyxl
```

The script imports `google.colab.files` for upload and download. That module is available in Google Colab; running the script locally requires adapting those file-handling steps.

---

## 8. Methodological & Operational Limitations

1. **Single Workbook:** The script processes the first uploaded workbook.
2. **Fixed Worksheet Name:** The input worksheet must be named `Submissions History`.
3. **Required Columns:** Both `Assigned To` and `Actual SLA` must be present.
4. **Exact Employee Grouping:** Employee records are grouped by the cleaned employee name. Different spellings or name variations are not automatically merged.
5. **Shift Naming Format:** Shift extraction recognises A–G only when the name ends with the expected `Shift A` through `Shift G` pattern.
6. **Fixed Shift and Break Settings:** The current calculation uses a 9-hour shift and a 1-hour break for every employee. Individual shift schedules are not read from the workbook.
7. **Estimated Idle Time:** The calculation uses total Actual SLA duration. It does not calculate real inactivity gaps between submission timestamps.
8. **Negative Results:** Any negative calculated idle time is reported as zero.
9. **Invalid Records:** Invalid records are excluded from employee-level calculations and retained in the `Invalid Rows` worksheet.
10. **No Dashboard:** The current script creates an Excel report and console summaries. It does not create a Power BI dashboard.

---

## 9. Submission Manifest

- `employee_idle_time_analysis.py`: Main Python script for Excel upload, data cleaning, employee-wise calculations, shift sorting, and report generation.
- `README.md`: Project overview, calculation methodology, input requirements, execution steps, output description, and limitations.
- `Employee_Idle_Time_Final_Report.xlsx`: Generated output workbook containing the four report worksheets. This file is created when the script runs; it is not required to run the project.

---

## 10. Possible Future Enhancements

* Configurable shift and break timings for each shift.
* Support for multiple uploaded Excel workbooks and worksheets.
* Employee-name mapping to combine known name variations.
* Date-wise and time-range-wise employee productivity analysis.
* Actual inactivity-gap calculation using reliable submission timestamps.
* Additional shift-level summary worksheets and charts.
* Power BI integration for interactive reporting.

---

**Project Focus:** Excel reporting automation, employee-wise productivity analysis, time-duration calculations, data validation, and shift-wise reporting using Python.
