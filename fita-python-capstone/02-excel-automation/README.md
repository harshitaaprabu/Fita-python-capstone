# Capstone 2: Excel Automation with Python

Automates the Microsoft Excel application using the `win32com` module.

## Steps covered
1. Install and import `win32com`
2. Open an existing Excel file and go to a specific worksheet
3. Modify data: update values, format cells, insert a column and a row, add formulas
4. Save the changes with `Save` / `SaveAs`
5. Close Excel with `Quit`

## Requirements
- Windows
- Microsoft Excel installed
- Python 3.8+

## Run
```bash
pip install -r requirements.txt
python excel_automation.py
```
A sample `sales.xlsx` is created on the first run, and the modified file is saved as `sales_updated.xlsx`.
