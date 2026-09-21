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
        
        # 1. Top Section (Before Table): Monthly Targets & Notes
        assert ws["A5"].value == "MONTHLY TARGETS"
        assert "1." in str(ws["A6"].value)
        assert ws["A11"].value == "NOTES:"
        assert ws["A12"].fill.start_color.rgb in ("00FFFFE6", "FFFFE6")
        print("✓ Top Section (Monthly Targets at Row 5, Notes at Row 11) verified")

        # 2. Main Protocol Table Headers (Rows 16 & 17)
        month_days = 30 if "Sep" in name or "Nov" in name else 31
        last_day_col = 3 + month_days
        succ_col_let = openpyxl.utils.get_column_letter(last_day_col + 1)
        done_tgt_col_let = openpyxl.utils.get_column_letter(last_day_col + 2)

        assert ws["A16"].value == "S.No."
        assert ws["B16"].value == "Protocols"
        assert ws["C16"].value == "Target"
        assert ws[f"{succ_col_let}16"].value == "Success %"
        assert ws[f"{done_tgt_col_let}16"].value == "Done / Target"
        print(f"✓ Table Headers verified: 'Success %' at {succ_col_let}16 and new 'Done / Target' column at {done_tgt_col_let}16")

        # 3. 20 Protocol rows (Rows 18 to 37)
        for i in range(1, 21):
            row = 17 + i
            assert ws[f"A{row}"].value == i
            fill = ws[f"A{row}"].fill.start_color.rgb
            assert fill is not None
            # Check Done / Target formula
            formula = str(ws[f"{done_tgt_col_let}{row}"].value)
            assert " / " in formula, f"Expected ' / ' in Done / Target formula at row {row}"
        print("✓ 20 Protocol rows (Rows 18-37) verified with 'Done / Target' formula")

        # 4. DAILY TOTAL SCORE (Row 38)
        assert ws["A38"].value == "DAILY TOTAL SCORE"
        score_done_tgt_formula = str(ws[f"{done_tgt_col_let}38"].value)
        assert " / " in score_done_tgt_formula
        print(f"✓ DAILY TOTAL SCORE (Row 38) verified with Done / Target summary at {done_tgt_col_let}38")

        # 5. DAILY SUCCESS % (Row 39)
        assert ws["A39"].value == "DAILY SUCCESS %"
        assert ws["D39"].number_format == "0.0%"
        print("✓ DAILY SUCCESS % (Row 39) verified")

        # 6. Success and Punishment (Stakes) between Protocol Table and Sleep Table (Rows 41 & 42)
        assert "PUNISHMENT" in str(ws["A41"].value), f"Expected PUNISHMENT at A41, got {ws['A41'].value}"
        assert "REWARD" in str(ws["A42"].value), f"Expected REWARD at A42, got {ws['A42'].value}"
        print("✓ Success & Punishment stakes verified between Protocol Table and Sleep Table (Rows 41 & 42)")

        # 7. Exactly Two-Row Gap (Rows 43 & 44)
        for r_gap in (43, 44):
            for c_idx in range(1, 10):
                c_let = openpyxl.utils.get_column_letter(c_idx)
                assert ws[f"{c_let}{r_gap}"].value is None, f"Expected blank gap at {c_let}{r_gap}"
        print("✓ Exactly two blank gap rows (Rows 43 & 44) verified")

        # 8. Sleep Tracking Section (Starts at Row 45)
        assert ws["A45"].value == "Sleep Tracking"
        assert "Sleep Hours" in str(ws["A47"].value)
        print("✓ Sleep Tracking section starts at Row 45 after the 2-row gap")

        # 9. Sleep Line Chart
        assert len(ws._charts) == 1
        chart = ws._charts[0]
        assert "Sleep" in str(chart.title)
        print(f"✓ Dynamic Sleep Line Chart verified at A49: '{chart.title}'")

    print("\nALL VERIFICATIONS PASSED 100% SUCCESSFULLY!")

if __name__ == "__main__":
    test_workbook()
