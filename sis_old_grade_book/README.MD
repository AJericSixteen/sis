# SIS Grade Automation - Old Grade Book

A Python automation tool for transferring grades from Excel workbooks to the PHINMA / Saint Jude College Manila Student Information System (SIS).

## Overview

This is the **legacy version** of the SIS Grade Automation tool. It provides the same core functionality as the global grade book but represents an earlier implementation.

For the latest version with additional features, see the `sis_global_grade_book` folder.

## What It Does

This automation script streamlines the grade entry process:

1. Select an Excel `.xlsx` or `.xlsm` file using a Windows file picker
2. Choose the worksheet to use from your workbook
3. Specify the student-name column
4. Specify the score column
5. Set the starting and ending Excel rows to process
6. Provide the SIS URL
7. Opens a visible browser for manual login
8. You log into SIS manually (script does not store your password)
9. Reads each student's name and score from the selected Excel range
10. Finds the matching student in SIS and enters the score into the **Obtained (%)** field
11. Keeps the browser open after automation finishes for manual verification and submission

## Example

```
Excel: Dela Cruz, Juan → 85.25
↓
SIS: Dela Cruz, Juan → Obtained (%) = 85
```

## Prerequisites

- **Python 3.10+**
- Windows operating system (for file picker)

## Installation

### 1. Install Python

Install Python 3.10 or higher if not already installed.

### 2. Install Dependencies

Open PowerShell or Command Prompt in this folder and run:

```powershell
pip install -r requirements.txt
playwright install chromium
```

**Alternative** (if `python` is not recognized but `py` works):

```powershell
py -m pip install -r requirements.txt
py -m playwright install chromium
```

## Excel File Setup

You **do not** need to rename your workbook to `grades.xlsx`.

When the program starts, a Windows file picker lets you select the Excel file you want to use.

### Configuration Prompts

The workbook can contain any worksheet. At runtime, the program asks you for:

- **Sheet name** - Which worksheet to use
- **Student name column** - Column letter (e.g., `C`)
- **Score column** - Column letter (e.g., `R`)
- **Starting row** - First row to process (e.g., `5`)
- **Ending row** - Last row to process (e.g., `52`)

### Example Configuration

```
Sheet name: P1
Student name column: C
Score column: R
Starting row: 5
Ending row: 52
```

This reads student names from column C and scores from column R, rows 5 through 52, on the P1 worksheet.

## Score Display and Decimals

The program enters the same **whole-number value that Excel displays** when the grade cell is formatted to show zero decimal places.

The script uses **proper whole-number rounding** so that the result matches what Excel displays:

| Excel Value | Displayed | Entered in SIS |
|-------------|-----------|----------------|
| 85.25       | 85        | 85             |
| 85.49       | 85        | 85             |
| 85.50       | 86        | 86             |
| 85.75       | 86        | 86             |
| 88.25       | 88        | 88             |

**Note:** The program does not simply remove the decimal portion. It uses normal whole-number rounding to match Excel's display.

## Running the Automation

Execute the script:

```powershell
py sis_grade_old.py
```

### Workflow

1. Program prompts you to select an Excel file
2. You provide the configuration (sheet name, columns, rows, SIS URL)
3. Browser opens with SIS login page
4. **You log in manually** - the script waits for you
5. Script reads Excel data and enters scores into SIS
6. Browser remains open for verification and manual save/submit

## Configuration

Edit the Python file to customize behavior:

```python
# True = pause after every student for verification
# False = continue automatically
TEST_MODE = True

# Wait time (in seconds) after updating each grade
WAIT_AFTER_UPDATE = 1.0
```

## Security & Privacy

- ✅ **Password is never stored** by this script
- ✅ **Manual SIS login** ensures complete control over authentication
- ✅ **Browser remains open** after completion for verification before submission

## Troubleshooting

### "python is not recognized"

Use `py` instead:

```powershell
py -m pip install -r requirements.txt
py -m playwright install chromium
py sis_grade_old.py
```

### Module not found errors

Ensure all dependencies are installed:

```powershell
pip install -r requirements.txt
```

### Playwright/Chromium issues

Reinstall Playwright and Chromium:

```powershell
pip install --upgrade playwright
playwright install chromium
```

## Migration to Latest Version

For new projects or updated features, consider using the version in `sis_global_grade_book` instead.

## Support

For additional help or questions, refer to the main project README at the root folder or consult the inline code comments in `sis_grade_old.py`.

---

**Version:** Old Grade Book (Legacy)  
**Last Updated:** 2024
