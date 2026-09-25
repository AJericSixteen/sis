"""
SIS Grade Automation
--------------------
Fully dynamic Excel + SIS configuration.

At startup you choose:
- Excel .xlsx/.xlsm file
- Worksheet name
- Student-name column
- Score column
- Starting row
- Ending row
- SIS URL

The script then reads the selected Excel range and enters each score
into the corresponding student's "Obtained (%)" field in SIS.

IMPORTANT:
- You log into SIS manually.
- Keep TEST_MODE=True while testing.
"""

import subprocess
import time
import tkinter as tk
from tkinter import filedialog
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


# =========================
# CONFIG
# =========================

# True = pause after every student so you can verify the entry.
# False = continue automatically.
TEST_MODE = True

WAIT_AFTER_UPDATE = 1.0

# Cloudflare/Turnstile detects a CDP debugger attached at launch, so
# Playwright never launches the browser itself. Instead the user opens
# a normal Edge window (below), passes the challenge and logs in by
# hand, and Playwright only attaches afterwards via connect_over_cdp.
CDP_PORT = 9222
EDGE_PROFILE_DIR = str(Path.home() / ".sis_grade_old_edge_profile")

EDGE_EXE_CANDIDATES = [
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
]


def find_edge_exe():
    for candidate in EDGE_EXE_CANDIDATES:
        if Path(candidate).exists():
            return candidate

    return "msedge.exe"


def launch_edge_for_manual_login(target_url):
    edge_exe = find_edge_exe()

    subprocess.Popen(
        [
            edge_exe,
            f"--remote-debugging-port={CDP_PORT}",
            f"--user-data-dir={EDGE_PROFILE_DIR}",
            target_url,
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


# =========================
# FILE SELECTION
# =========================

def select_excel_file():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    file_path = filedialog.askopenfilename(
        title="Select Excel Grade File",
        filetypes=[
            ("Excel files", "*.xlsx"),
            ("Excel macro-enabled files", "*.xlsm"),
            ("All files", "*.*"),
        ],
    )

    root.destroy()

    if not file_path:
        return None

    return file_path


# =========================
# INPUT HELPERS
# =========================

def ask_non_empty(prompt):
    while True:
        value = input(prompt).strip()

        if value:
            return value

        print("This field cannot be empty.")


def ask_positive_integer(prompt):
    while True:
        try:
            value = int(input(prompt))

            if value >= 1:
                return value

            print("Please enter a number greater than 0.")

        except ValueError:
            print("Please enter numbers only.")


def ask_column(prompt):
    while True:
        value = input(prompt).strip().upper()

        # Accept A, B, C, AA, AB, etc.
        if value.isalpha():
            return value

        print("Please enter a valid Excel column, such as C, R, AA, etc.")


def ask_sis_url():
    while True:
        url = input("Paste the SIS URL: ").strip()

        if url.startswith(("http://", "https://")):
            return url

        print("Invalid URL. Please paste the complete SIS URL.")


# =========================
# EXCEL
# =========================

def clean_name(value):
    if value is None:
        return ""

    return " ".join(str(value).strip().upper().split())


def whole_number(score):
    """Remove the decimal portion without rounding."""
    return int(float(score))


def get_excel_settings(excel_file):
    """
    Ask the user for worksheet, name column, score column,
    and row range. Validate everything against the workbook.
    """

    wb = openpyxl.load_workbook(
        excel_file,
        read_only=True,
        data_only=True,
    )

    print()
    print("Available worksheets:")
    for sheet in wb.sheetnames:
        print(f"  - {sheet}")

    print()

    while True:
        sheet_name = ask_non_empty("Enter sheet name: ")

        if sheet_name in wb.sheetnames:
            break

        print(
            f"Sheet '{sheet_name}' was not found."
        )

        print(
            "Please choose one of the worksheets listed above."
        )

    name_column = ask_column(
        "Enter student name column (e.g. C): "
    )

    score_column = ask_column(
        "Enter score column (e.g. R): "
    )

    start_row = ask_positive_integer(
        "Enter starting row: "
    )

    end_row = ask_positive_integer(
        "Enter ending row: "
    )

    if end_row < start_row:
        wb.close()
        raise ValueError(
            "Ending row must be greater than or equal to starting row."
        )

    wb.close()

    return (
        sheet_name,
        name_column,
        score_column,
        start_row,
        end_row,
    )


def load_students(
    excel_file,
    sheet_name,
    name_column,
    score_column,
    start_row,
    end_row,
):
    """
    Read the selected range from the selected worksheet.
    """

    wb = openpyxl.load_workbook(
        excel_file,
        data_only=True,
    )

    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(
            f"Sheet '{sheet_name}' was not found."
        )

    ws = wb[sheet_name]

    last_row = min(end_row, ws.max_row)

    students = []

    for row in range(start_row, last_row + 1):

        raw_name = ws[
            f"{name_column}{row}"
        ].value

        raw_score = ws[
            f"{score_column}{row}"
        ].value

        name = clean_name(raw_name)

        # Ignore blank student rows.
        if not name:
            continue

        # Ignore rows with no score.
        if raw_score is None or str(raw_score).strip() == "":
            continue

        students.append({
            "row": row,
            "name": name,
            "score": whole_number(raw_score),
        })

    wb.close()

    return students


# =========================
# PREVIEW
# =========================

def show_preview(
    excel_file,
    sheet_name,
    name_column,
    score_column,
    start_row,
    end_row,
    sis_url,
    students,
):
    print()
    print("=" * 65)
    print("                         SETTINGS")
    print("=" * 65)

    print(f"Excel file   : {excel_file}")
    print(f"Sheet        : {sheet_name}")
    print(f"Name column  : {name_column}")
    print(f"Score column : {score_column}")
    print(f"Rows         : {start_row} - {end_row}")
    print(f"SIS URL      : {sis_url}")

    print()
    print("-" * 65)
    print("PREVIEW")
    print("-" * 65)

    if not students:
        print("No students with both a name and score were found.")
        print("-" * 65)
        return

    # Show up to 10 rows.
    for student in students[:10]:
        print(
            f"Row {student['row']:<5} "
            f"{student['name']} -> {student['score']}"
        )

    if len(students) > 10:
        print(
            f"... and {len(students) - 10} more students"
        )

    print("-" * 65)
    print(f"Students to process: {len(students)}")


# =========================
# SIS
# =========================

def find_student(page, student_name):
    """
    Try several text matching approaches.
    """

    # Exact visible text.
    locator = page.get_by_text(
        student_name,
        exact=True,
    )

    if locator.count() > 0:
        return locator.first

    # Fallback.
    locator = page.get_by_text(
        student_name,
    )

    if locator.count() > 0:
        return locator.first

    return None


def find_obtained_input(page, student_locator):
    """
    Try to find the Obtained (%) input associated with
    the student's table row.
    """

    # First: nearest table row.
    row = student_locator.locator(
        "xpath=ancestor::tr[1]"
    )

    if row.count() > 0:

        inputs = row.locator("input")

        if inputs.count() > 0:
            return inputs.last

    # Fallback selectors.
    candidates = [
        page.locator(
            'input[placeholder*="Obtained" i]'
        ),
        page.locator(
            'input[name*="obtained" i]'
        ),
        page.locator(
            'input[id*="obtained" i]'
        ),
        page.locator(
            'input[aria-label*="Obtained" i]'
        ),
    ]

    for candidate in candidates:

        if candidate.count() > 0:
            return candidate.first

    return None


# =========================
# MAIN
# =========================

def main():

    print()
    print("=" * 65)
    print("                  SIS GRADE AUTOMATION")
    print("=" * 65)
    print()

    # ---------------------------------
    # Select Excel file
    # ---------------------------------

    print("Select your Excel file...")

    excel_file = select_excel_file()

    if not excel_file:
        print("No file selected. Exiting.")
        return

    print()
    print(f"Selected file: {excel_file}")

    # ---------------------------------
    # Excel settings
    # ---------------------------------

    (
        sheet_name,
        name_column,
        score_column,
        start_row,
        end_row,
    ) = get_excel_settings(excel_file)

    # ---------------------------------
    # SIS URL
    # ---------------------------------

    print()

    sis_url = ask_sis_url()

    # ---------------------------------
    # Load students
    # ---------------------------------

    print()
    print("Reading Excel file...")

    students = load_students(
        excel_file=excel_file,
        sheet_name=sheet_name,
        name_column=name_column,
        score_column=score_column,
        start_row=start_row,
        end_row=end_row,
    )

    # ---------------------------------
    # Preview
    # ---------------------------------

    show_preview(
        excel_file=excel_file,
        sheet_name=sheet_name,
        name_column=name_column,
        score_column=score_column,
        start_row=start_row,
        end_row=end_row,
        sis_url=sis_url,
        students=students,
    )

    if not students:
        print()
        print(
            "No students with both a name and score were found."
        )
        print()
        input("Press ENTER to close...")
        return

    # ---------------------------------
    # Confirmation
    # ---------------------------------

    print()

    confirmation = input(
        "Continue with these settings? (Y/N): "
    ).strip().upper()

    if confirmation != "Y":
        print("Cancelled.")
        return

    # ---------------------------------
    # Open SIS (manual launch, so Cloudflare
    # never sees a CDP debugger attached at
    # browser startup)
    # ---------------------------------

    print()
    print("=" * 65)
    print("                         SIS LOGIN")
    print("=" * 65)
    print()
    print("IMPORTANT: close every open Edge window first.")
    input("Press ENTER once all Edge windows are closed...")

    launch_edge_for_manual_login(sis_url)

    print()
    print("1. Log into SIS manually.")
    print("2. Pass the Cloudflare check yourself if one appears.")
    print("3. Navigate to the page containing the students.")
    print("4. Make sure the 'Obtained (%)' fields are visible.")
    print("5. Return to this PowerShell window.")
    print()

    input(
        "Press ENTER when the SIS page is ready..."
    )

    with sync_playwright() as p:

        try:
            browser = p.chromium.connect_over_cdp(
                f"http://localhost:{CDP_PORT}"
            )
        except Exception as e:
            print(f"Could not connect to Edge on port {CDP_PORT}: {e}")
            input("Press ENTER to close...")
            return

        context = browser.contexts[0] if browser.contexts else browser.new_context()

        page = None
        sis_prefix = sis_url.split("?")[0][:40]

        for candidate in context.pages:
            if candidate.url.startswith(sis_prefix):
                page = candidate
                break

        if page is None:
            page = context.pages[0] if context.pages else context.new_page()

        # ---------------------------------
        # Process students
        # ---------------------------------

        results = []

        for index, student in enumerate(
            students,
            start=1,
        ):

            name = student["name"]
            score = student["score"]
            excel_row = student["row"]

            print()
            print(
                f"[{index}/{len(students)}] "
                f"Excel row {excel_row}"
            )

            print(f"Student: {name}")
            print(f"Score:   {score}")

            try:

                # -------------------------
                # Find student
                # -------------------------

                student_locator = find_student(
                    page,
                    name,
                )

                if student_locator is None:

                    print(
                        "  NOT FOUND"
                    )

                    results.append(
                        (
                            excel_row,
                            name,
                            score,
                            "NOT FOUND",
                        )
                    )

                    continue

                print(
                    "  Student found"
                )

                # -------------------------
                # Scroll / click
                # -------------------------

                student_locator.scroll_into_view_if_needed()

                student_locator.click()

                time.sleep(0.5)

                # -------------------------
                # Find score input
                # -------------------------

                obtained = find_obtained_input(
                    page,
                    student_locator,
                )

                if obtained is None:

                    print(
                        "  Obtained (%) input not found"
                    )

                    results.append(
                        (
                            excel_row,
                            name,
                            score,
                            "INPUT NOT FOUND",
                        )
                    )

                    continue

                # -------------------------
                # Enter score
                # -------------------------

                obtained.scroll_into_view_if_needed()

                obtained.click()

                obtained.fill(
                    str(score)
                )

                # Trigger blur/change.
                obtained.press("Tab")

                print(
                    f"  Entered {score}"
                )

                # -------------------------
                # Test mode
                # -------------------------

                if TEST_MODE:

                    print()
                    print(
                        "  TEST MODE:"
                    )

                    print(
                        "  Verify the score in SIS."
                    )

                    input(
                        "  Press ENTER to continue..."
                    )

                else:

                    time.sleep(
                        WAIT_AFTER_UPDATE
                    )

                results.append(
                    (
                        excel_row,
                        name,
                        score,
                        "UPDATED",
                    )
                )

            except PlaywrightTimeoutError:

                print(
                    "  TIMEOUT"
                )

                results.append(
                    (
                        excel_row,
                        name,
                        score,
                        "TIMEOUT",
                    )
                )

            except Exception as e:

                print(
                    f"  ERROR: {e}"
                )

                results.append(
                    (
                        excel_row,
                        name,
                        score,
                        f"ERROR: {e}",
                    )
                )

        # ---------------------------------
        # Results
        # ---------------------------------

        print()
        print()
        print("=" * 65)
        print("                         RESULTS")
        print("=" * 65)

        for (
            excel_row,
            name,
            score,
            status,
        ) in results:

            print(
                f"Row {excel_row:<5} "
                f"{status:<18} "
                f"{name} -> {score}"
            )

        print()
        print(
            "Review the SIS entries before saving/submitting."
        )

        print()
        print(
            "You can close Edge manually when you're done."
        )

        input(
            "Press ENTER to finish..."
        )


if __name__ == "__main__":
    main()
