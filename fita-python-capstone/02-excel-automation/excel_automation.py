"""
Capstone Project 2 - Excel automation using Python (win32com)

Steps practised (same order as the syllabus):
  1. Install and import win32com          pip install pywin32
  2. Open an existing Excel file and navigate to a specific worksheet
  3. Modify data: update values, format cells, insert rows/columns
  4. Save the changes (Save / SaveAs)
  5. Close Excel using Quit

Requires Windows with Microsoft Excel installed (win32com drives the real Excel app).
Run: python excel_automation.py
It creates a sample sales.xlsx first, so the demo works out of the box.
"""
import os

import win32com.client as win32

FOLDER = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(FOLDER, "sales.xlsx")
OUTPUT = os.path.join(FOLDER, "sales_updated.xlsx")

XL_CENTER = -4108      # xlCenter constant
XL_EDGE_BOTTOM = 9     # xlEdgeBottom constant


def create_sample_file(excel):
    """Make an 'existing' workbook so the project can be run immediately."""
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Sales"
    rows = [
        ("Product", "Units", "Price"),
        ("Keyboard", 40, 550),
        ("Mouse", 85, 300),
        ("Monitor", 12, 8200),
        ("Webcam", 30, 1500),
    ]
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row, start=1):
            ws.Cells(r, c).Value = value
    wb.SaveAs(SOURCE)
    wb.Close()


def main():
    excel = win32.Dispatch("Excel.Application")   # Step 1: start Excel via COM
    excel.Visible = True
    excel.DisplayAlerts = False                   # don't stop for overwrite prompts

    try:
        if not os.path.exists(SOURCE):
            create_sample_file(excel)

        # Step 2: open existing file, go to a worksheet
        wb = excel.Workbooks.Open(SOURCE)
        ws = wb.Worksheets("Sales")

        # Step 3a: update values (Mouse price 300 -> 320)
        for r in range(2, ws.UsedRange.Rows.Count + 1):
            if ws.Cells(r, 1).Value == "Mouse":
                ws.Cells(r, 3).Value = 320

        # Step 3b: insert a new column for totals and fill it with formulas
        ws.Columns(4).Insert()
        ws.Cells(1, 4).Value = "Total"
        last_row = ws.UsedRange.Rows.Count
        for r in range(2, last_row + 1):
            ws.Cells(r, 4).Formula = f"=B{r}*C{r}"

        # Step 3c: insert a new row and add a product
        ws.Rows(3).Insert()
        for c, v in enumerate(("Headset", 25, 1800), start=1):
            ws.Cells(3, c).Value = v
        ws.Cells(3, 4).Formula = "=B3*C3"

        # Step 3d: grand total row
        last_row = ws.UsedRange.Rows.Count
        ws.Cells(last_row + 1, 1).Value = "Grand Total"
        ws.Cells(last_row + 1, 4).Formula = f"=SUM(D2:D{last_row})"

        # Step 3e: format cells
        header = ws.Range("A1:D1")
        header.Font.Bold = True
        header.Interior.Color = 0xC77C1F      # BGR blue
        header.Font.Color = 0xFFFFFF
        header.HorizontalAlignment = XL_CENTER
        ws.Range(f"C2:D{last_row + 1}").NumberFormat = "₹#,##0"
        ws.Range(f"A{last_row + 1}:D{last_row + 1}").Font.Bold = True
        ws.Range(f"A{last_row + 1}:D{last_row + 1}").Borders(8).LineStyle = 1
        ws.Columns("A:D").AutoFit()

        # Step 4: save the changes (SaveAs keeps the original untouched; wb.Save() overwrites)
        wb.SaveAs(OUTPUT)
        print(f"Saved: {OUTPUT}")
        wb.Close()
    finally:
        # Step 5: always close Excel
        excel.Quit()


if __name__ == "__main__":
    main()
