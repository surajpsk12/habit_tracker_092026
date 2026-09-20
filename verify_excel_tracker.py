import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

def test_workbook():
    wb = openpyxl.load_workbook("Habit_Tracker_2026_2027.xlsx", data_only=False)
    sheet_names = wb.sheetnames
    print(f"Total Sheets: {len(sheet_names)}")
    print(f"Sheets: {sheet_names}")
    
    assert "📊 Master Dashboard" in sheet_names
    assert "Sep 2026" in sheet_names
    assert "Oct 2026" in sheet_names
    assert "Dec 2027" in sheet_names
    assert len(sheet_names) == 17

    # Inspect Dashboard Charts
    dash_ws = wb["📊 Master Dashboard"]
    print(f"Dashboard Charts Count: {len(dash_ws._charts)}")
    assert len(dash_ws._charts) == 2, f"Expected 2 charts, got {len(dash_ws._charts)}"
    for idx, c in enumerate(dash_ws._charts, 1):
        print(f"Chart {idx}: Title='{c.title}', Type={type(c).__name__}")

    # Inspect Sep 2026 sheet
    ws = wb["Sep 2026"]
    print(f"Sep 2026 Title: {ws['A2'].value}")
    print(f"Sep 2026 Day 1 col width: {ws.column_dimensions['E'].width}")
    print(f"Sep 2026 Category col width: {ws.column_dimensions['B'].width}")
    print(f"Sep 2026 Habit col width: {ws.column_dimensions['C'].width}")
    print(f"Sep 2026 Validations Count: {len(ws.data_validations.dataValidation)}")
    print(f"Sep 2026 Conditional Formattings Count: {len(ws.conditional_formatting)}")

    # Check formulas
    last_day_col_letter = openpyxl.utils.get_column_letter(4 + 30) # 30 days in Sep
    tot_col_letter = openpyxl.utils.get_column_letter(4 + 30 + 1)
    succ_col_letter = openpyxl.utils.get_column_letter(4 + 30 + 2)
    print(f"Sep 2026 Habit 1 Completed Formula ({tot_col_letter}7): {ws[f'{tot_col_letter}7'].value}")
    print(f"Sep 2026 Habit 1 Success % Formula ({succ_col_letter}7): {ws[f'{succ_col_letter}7'].value}")

    print("ALL VERIFICATIONS COMPLETED SUCCESSFULLY!")

if __name__ == "__main__":
    test_workbook()
