import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

def test_workbook():
    fn = "Habit_Tracker_Sep_Dec_2026.xlsx"
    wb = openpyxl.load_workbook(fn, data_only=False)
    sheet_names = wb.sheetnames
    print(f"\n============================\nVerifying {fn}:")
    print(f"Total Sheets: {len(sheet_names)}")
    print(f"Sheets: {sheet_names}")
    
    expected_sheets = ["Sep 2026", "Oct 2026", "Nov 2026", "Dec 2026"]
    assert sheet_names == expected_sheets, f"Expected {expected_sheets}, got {sheet_names}"

    for name in expected_sheets:
        ws = wb[name]
        print(f"\n--- Checking sheet: {name} ---")
        assert ws["A2"].value == "Habit Tracker"
        
        # 1. Verify Top Section (Before Table)
        print("Verifying Top Section (Before Table)...")
        assert "PUNISHMENT" in str(ws["A5"].value), f"Expected PUNISHMENT at A5, got {ws['A5'].value}"
        assert "REWARD" in str(ws["A6"].value), f"Expected REWARD at A6, got {ws['A6'].value}"
        assert ws["A8"].value == "MONTHLY TARGETS", f"Expected MONTHLY TARGETS at A8, got {ws['A8'].value}"
        assert "1." in str(ws["A9"].value), f"Expected Target 1 at A9, got {ws['A9'].value}"
        assert ws["A14"].value == "NOTES:", f"Expected NOTES: at A14, got {ws['A14'].value}"
        # Yellow notes box
        assert ws["A15"].fill.start_color.rgb == "00FFFFE6" or ws["A15"].fill.start_color.rgb == "FFFFE6"
        print("✓ Top Section (Punishment, Reward, Monthly Targets, Notes) verified before table!")

        # 2. Verify Table Header at Rows 19 & 20
        assert ws["A19"].value == "S.No."
        assert ws["B19"].value == "Protocols"
        assert ws["C19"].value == "Target"
        print("✓ Protocol Table Header at Rows 19 & 20 verified")

        # 3. Verify 20 Protocol rows (Rows 21 to 40)
        for i in range(1, 21):
            row = 20 + i
            assert ws[f"A{row}"].value == i, f"Expected S.No {i} at row {row}, got {ws[f'A{row}'].value}"
            fill = ws[f"A{row}"].fill.start_color.rgb
            assert fill is not None, f"Row {row} missing fill"
        print("✓ 20 Protocol rows verified (S.No. 1 to 20, Rows 21 to 40)")

        # 4. Check Data Validations
        dvs = ws.data_validations.dataValidation
        assert len(dvs) == 2, f"Expected 2 data validations, got {len(dvs)}"
        print("✓ Data validations count verified (Target dropdown & Sleep dropdown)")

        # 5. Check DAILY TOTAL SCORE at Row 41
        assert ws["A41"].value == "DAILY TOTAL SCORE"
        assert "21:D40" in str(ws["D41"].value), f"Expected formula with rows 21 to 40, got {ws['D41'].value}"
        print("✓ DAILY TOTAL SCORE at Row 41 verified with formula referencing rows 21 to 40")

        # 6. Check Sleep Tracking Section
        assert ws["A43"].value == "Sleep Tracking"
        assert "Sleep Hours" in str(ws["A45"].value)
        # Check day column width is wide enough (>= 4.5) to easily display numbers like 1-24 alongside the dropdown button
        d_col_letter = openpyxl.utils.get_column_letter(4)
        col_w = ws.column_dimensions[d_col_letter].width
        print(f"Day column width: {col_w}")
        assert col_w >= 4.5, f"Day column width {col_w} is too small for dropdown number visibility"
        print("✓ Day column width verified for clear in-cell number visibility!")

        # 7. Check Line Chart
        assert len(ws._charts) == 1, f"Expected 1 chart on sheet, got {len(ws._charts)}"
        chart = ws._charts[0]
        assert "Sleep" in str(chart.title)
        print(f"✓ Chart verified: '{chart.title}'")

    print("\nALL VERIFICATIONS PASSED 100% SUCCESSFULLY!")

if __name__ == "__main__":
    test_workbook()
