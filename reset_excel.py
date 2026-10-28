"""
Reset Helper Script:
Resets Column B ('tick yes after filling form ') to False for all participants or a specific range.
"""

import sys
import openpyxl

EXCEL_FILE = "Google Cloud Learn with Communities - Participant Enrolment Form (Responses).xlsx"

def reset_ticks(start_row=2, end_row=None):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb.active
    
    max_row = end_row if end_row else sheet.max_row
    reset_count = 0
    
    for r in range(start_row, max_row + 1):
        # Keep Row 197 as True if it's the 'No one is supposed to fill this' row
        if r == 197:
            continue
        if sheet.cell(row=r, column=2).value is not False:
            sheet.cell(row=r, column=2, value=False)
            reset_count += 1
            
    wb.save(EXCEL_FILE)
    print(f"Successfully reset {reset_count} rows in '{EXCEL_FILE}' to False (from row {start_row} to {max_row}).")

if __name__ == "__main__":
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 2
    end = int(sys.argv[2]) if len(sys.argv) > 2 else None
    reset_ticks(start, end)
