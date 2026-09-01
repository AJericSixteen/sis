# SIS Grade Automation

A Python automation tool for transferring grades from Excel workbooks to the PHINMA / Saint Jude College Manila Student Information System (SIS).

## Overview

This project provides an automated solution to streamline the grade entry process by reading student names and scores from Excel files and automatically inputting them into the SIS system. The automation handles the tedious task of grade entry while keeping you in control of authentication and verification.

## Project Structure

The project contains two versions of the grade automation:

### `sis_global_grade_book/` (Current/Recommended)
The latest version of the SIS Grade Automation with fully dynamic Excel and SIS configuration capabilities.

### `sis_old_grade_book/` (Legacy)
The previous version of the automation. Maintained for reference.

## Features

The automation tool provides the following capabilities:

- **File Selection**: Choose any Excel `.xlsx` or `.xlsm` file using a Windows file picker at runtime
- **Flexible Worksheet Selection**: Select which worksheet to use from your workbook
- **Dynamic Column Mapping**: Choose the student-name column and score column at runtime
- **Row Range Selection**: Specify the starting and ending rows to process
- **Manual Authentication**: Log into SIS manually (password not stored)
- **Automated Grade Entry**: Reads student data from Excel and enters scores into SIS
- **Intelligent Rounding**: Uses proper whole-number rounding to match Excel's displayed grades
- **Browser Control**: Opens a visible browser for monitoring and manual verification
- **Test Mode**: Optional pause after each student for verification

## Example

```
Excel: Dela Cruz, Juan → 85.25
SIS:   Dela Cruz, Juan → Obtained (%) = 85
```

The score `85.25` displays as `85` in Excel (formatted to 0 decimal places) and is entered as `85` in SIS using proper rounding rules.

## Requirements

- **Python 3.10+**
- **openpyxl** (≥3.1) - For Excel file reading
- **Playwright** (≥1.45) - For browser automation and Chromium browser

## Installation

### 1. Install Python
Ensure Python 3.10 or higher is installed on your system.

### 2. Install Dependencies

Navigate to the `sis_global_grade_book/` folder (or `sis_old_grade_book/` if using the legacy version) and run:

```powershell
pip install -r requirements.txt
playwright install chromium
```

**Alternative** (if `python` is not recognized but `py` works):

```powershell
py -m pip install -r requirements.txt
py -m playwright install chromium
```

## Usage

### 1. Prepare Your Excel File

- The workbook can have any name (no need to rename to `grades.xlsx`)
- It can contain multiple worksheets
- Organize your data with student names and scores in columns

### 2. Run the Automation

Navigate to the `sis_global_grade_book/` folder and execute:

```powershell
py sis_global_automation.py
```

### 3. Configure at Runtime

When the program starts, you'll be prompted to provide:

| Setting | Example |
|---------|---------|
| Excel file | Select via file picker |
| Sheet name | `P1` |
| Student name column | `C` |
| Score column | `R` |
| Starting row | `5` |
| Ending row | `52` |
| SIS URL | Paste the SIS URL |

**Example Configuration:**
- Read student names from column C
- Read scores from column R
- Process rows 5 through 52
- On the P1 worksheet

### 4. Manual Login & Verification

1. The browser opens and displays the SIS login page
2. Log in manually (your credentials are not stored)
3. The automation will find matching students and enter scores
4. The browser remains open after completion for you to verify and save/submit grades

## Score Rounding

The automation uses Excel's standard rounding rules:

| Excel Value | Displayed | Entered in SIS |
|-------------|-----------|----------------|
| 85.25       | 85        | 85             |
| 85.49       | 85        | 85             |
| 85.50       | 86        | 86             |
| 85.75       | 86        | 86             |
| 88.25       | 88        | 88             |

The program does **not** simply truncate decimals—it uses proper rounding to match what Excel displays when formatted to zero decimal places.

## Test Mode

In `sis_global_automation.py`, you can enable test mode for safer testing:

```python
TEST_MODE = True  # Pause after each student for verification
```

Set to `False` for automated continuous processing.

## Important Notes

- ⚠️ **Your password is never stored or transmitted** by this script
- 🔒 Manual SIS login ensures complete control over authentication
- ✅ Browser remains open for manual verification before final submission
- 📊 The script handles rounding to match Excel's display format

## Troubleshooting

### Python not recognized on Windows
If `python` doesn't work, try:
```powershell
py -m pip install -r requirements.txt
py -m playwright install chromium
py sis_global_automation.py
```

### Missing dependencies
Ensure `requirements.txt` is in the same folder and run:
```powershell
pip install -r requirements.txt
```

### Playwright/Chromium issues
Reinstall Playwright:
```powershell
pip install --upgrade playwright
playwright install chromium
```

## Support & Questions

For questions or issues, refer to the individual README files in each folder:
- [sis_global_grade_book/README.txt](./sis_global_grade_book/README.txt)
- [sis_old_grade_book/README.txt](./sis_old_grade_book/README.txt)

---

**Last Updated:** 2024
