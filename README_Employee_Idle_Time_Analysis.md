

**Repository Link:** `https://github.com/Pratik-Das-02/Employee-Idle-Time-Shift-Wise-Productivity-Analysis.git`

# Employee Idle Time & Shift-Wise Productivity Analysis

A Python-based Excel automation project that calculates employee-wise file counts and estimated idle time from submission history data. It also organises employees by shift and creates a multi-sheet Excel report.

This project is designed to run in **Google Colab**. Upload an Excel workbook, let the script process the data, and download the generated report.

## 1. Project objective

In a manual reporting process, an analyst may need to:
- Count how many submissions were handled by each employee.
- Calculate the total Actual SLA duration for each employee.
- Estimate idle time using a fixed shift duration and break allowance.
- Arrange employees by Shift A, Shift B, Shift C and so on.
- Separate valid records from blank or invalid records.
- Prepare a structured Excel report for review.

This script automates those steps and reduces repetitive Excel work.

## 2. How the analysis works

The script reads the **Submissions History** worksheet and uses these fields:

| Input column | Purpose |
|---|---|
| `Assigned To` | Employee name or assigned employee/shift label |
| `Actual SLA` | Time taken to process a submission |

Column names are cleaned before validation. The script also accepts common formatting differences such as extra spaces, line breaks and non-breaking spaces.

### Calculation settings

The current script uses the following fixed settings:

| Setting | Value |
|---|---:|
| Shift duration | 9 hours |
| Break allowance | 1 hour |
| Shift productive-time allowance | 8 hours |
| Shift sorting order | A, B, C, D, E, F, G |

### Idle time formula

For each employee:

```text
Idle Time = Shift Duration - Break Time - Total Actual SLA
```

With the current settings:

```text
Idle Time = 9 hours - 1 hour - Total Actual SLA
```

The result is not allowed to go below zero. If the calculated idle time is negative, the script reports `00:00:00`.

**Important:** This is an estimated idle-time calculation based on the total Actual SLA duration. It does not measure real gaps between individual submissions or verify the employee's actual login, logout or break timings. The fixed shift and break settings should be reviewed before using the output for operational decisions.

## 3. Main features

- **Excel upload:** Select and upload an Excel workbook in Google Colab.
- **Worksheet selection:** Reads the worksheet named `Submissions History`.
- **Column validation:** Checks that `Assigned To` and `Actual SLA` are available.
- **Data cleaning:** Trims employee names and standardises whitespace.
- **SLA conversion:** Converts Excel time values and supported duration strings into time durations.
- **Invalid-row handling:** Separates rows with a blank employee name, missing/unreadable SLA or negative SLA.
- **Employee-wise summary:** Calculates the submission count and total Actual SLA for each employee.
- **Idle-time calculation:** Applies the configured shift duration and break allowance.
- **Shift extraction:** Extracts shift letters A to G from employee labels formatted like `Employee Name Shift A`.
- **Shift sorting:** Arranges recognised shifts in A-to-G order.
- **Excel report generation:** Creates a workbook with summary, calculation, validation and invalid-record worksheets.
- **Automatic download:** Downloads the completed Excel report in Google Colab.

## 4. Generated Excel report

The output file is named:

`Employee_Idle_Time_Final_Report.xlsx`

It contains four worksheets:

| Worksheet | What it contains |
|---|---|
| `Person Wise Idle Time` | Shift, employee name, submission count and idle time |
| `Detailed Calculation` | Employee-wise count, shift duration, total Actual SLA, break time and idle time |
| `Validation` | Input row count, valid records, invalid records, employee groups and calculation settings |
| `Invalid Rows` | Records excluded from the employee-wise calculation, for checking and correction |

The script also prints the employee-wise summary and shift-wise submission counts in the Colab output.

## 5. Input Excel requirements

Prepare an Excel workbook with a worksheet named exactly:

`Submissions History`

The worksheet must contain these columns:

| Column | Required? | Example |
|---|---|---|
| `Assigned To` | Yes | `Aman Shift A` |
| `Actual SLA` | Yes | `00:12:35` |

The Actual SLA field can contain Excel time/duration values or duration strings that Pandas can parse.

### Employee naming format

For shift extraction, use labels such as:

```text
Aman Shift A
Rohit Shift B
Priya Shift C
```

The script recognises shift letters A to G when the name ends with the pattern `Shift A` through `Shift G`. Names without a recognised shift suffix are still included in the employee summary, but their shift is left blank and they are sorted after recognised shifts.

## 6. How to run the project

### Option A — Run in Google Colab (recommended for the current script)

1. Open [Google Colab](https://colab.research.google.com/).
2. Create a new notebook.
3. Copy the Python script into a code cell.
4. Run the cell.
5. Upload your Excel workbook when prompted.
6. Wait for the processing to finish.
7. Download `Employee_Idle_Time_Final_Report.xlsx` when the download prompt appears.

The current script imports `google.colab.files`, so it is intended for Google Colab rather than direct execution as a normal local Python script.

### Option B — Install the Python dependencies

The main third-party packages used are:

```bash
pip install pandas numpy openpyxl
```

`google.colab` is provided by the Google Colab environment. It is not normally installed as part of a standard local Python setup. To run the script locally, the upload and download steps need to be adapted to use local file paths and normal file selection.

## 7. Repository files

A suggested GitHub repository layout:

```text
Employee-Idle-Time-Analysis/
│
├── README.md
├── employee_idle_time_analysis.py
├── requirements.txt
├── sample_data/
│   └── sample_input.xlsx
└── screenshots/
    └── report_preview.png
```

Suggested `requirements.txt`:

```text
pandas
numpy
openpyxl
```

Only add sample data and screenshots if they are safe to share publicly. Never upload real employee information, client data, confidential reports, passwords or company files.

## 8. Upload the project to GitHub

1. Sign in to [GitHub](https://github.com/).
2. Select **New repository**.
3. Enter the repository name: `Employee-Idle-Time-Analysis`.
4. Add a short description, for example: `Python automation for employee-wise submission counts and estimated idle-time reporting.`
5. Select **Public** if you want recruiters to view it, or **Private** if you want to keep it restricted.
6. Select **Add a README file** only if you are creating the README directly on GitHub. If you upload this README, you can leave that option unchecked.
7. Create the repository.
8. Select **Add file → Upload files**.
9. Upload `README.md`, your Python script and `requirements.txt`.
10. Select **Commit changes**.

You can also upload the project using Git from your computer:

```bash
git init
git add README.md employee_idle_time_analysis.py requirements.txt
git commit -m "Add employee idle time analysis project"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/Employee-Idle-Time-Analysis.git
git push -u origin main
```

Replace `YOUR-USERNAME` with your GitHub username and use the repository URL shown by GitHub.

## 9. Current assumptions and limitations

Please understand these points before presenting the results:

1. The script reads only the first workbook uploaded.
2. It reads only the worksheet named `Submissions History`.
3. It requires `Assigned To` and `Actual SLA`.
4. It groups records by the exact cleaned employee name. Different spellings or name variations are not automatically merged.
5. Shift extraction recognises only A to G in the expected `Shift A` format.
6. Shift duration and break time are fixed in the settings section. They are not calculated from actual employee schedules.
7. Idle time is estimated from total Actual SLA duration. It is not a real-time attendance or inactivity measurement.
8. Negative calculated idle time is clipped to zero.
9. Invalid records are excluded from employee calculations and included in the `Invalid Rows` worksheet.
10. The output report is an Excel workbook; the script does not currently create a Power BI dashboard.

## 10. Possible future improvements

The project can be extended with:
- Configurable shift duration and break time.
- Support for multiple uploaded workbooks and worksheets.
- Employee-name standardisation to merge known name variations.
- Shift-wise summary worksheets in the output Excel file.
- Date-wise and time-range-wise analysis.
- Actual activity-gap calculations when reliable submission timestamps are available.
- Charts for employee and shift comparisons.
- A local desktop interface or a Streamlit application.

These are possible enhancements and are not features of the current script.

## 11. Project summary

This project demonstrates how Python can automate a recurring operational reporting task using Excel data. It cleans input records, validates required fields, calculates employee-wise submission counts and estimated idle time, sorts employees by shift, and exports a structured Excel report.

**Technologies:** Python, Pandas, NumPy, OpenPyXL and Google Colab.

---

*Created as a practical data analytics and reporting automation portfolio project.*
