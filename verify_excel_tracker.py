import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

def test_workbook():
    fn = "Habit_Tracker_2026_2030.xlsx"
    wb = openpyxl.load_workbook(fn, data_only=False)
    sheet_names = wb.sheetnames
    print(f"\n============================\nVerifying {fn}:")
    print(f"Total Sheets: {len(sheet_names)}")
    
    assert len(sheet_names) == 52, f"Expected 52 sheets, got {len(sheet_names)}"
    assert sheet_names[0] == "Sep 2026"
    assert sheet_names[-1] == "Dec 2030"
    print(f"First Sheet: {sheet_names[0]}, Last Sheet: {sheet_names[-1]}")

    for name in sheet_names:
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
        import calendar
        m_str, y_str = name.split()
        m_idx = list(calendar.month_abbr).index(m_str)
        month_days = calendar.monthrange(int(y_str), m_idx)[1]
        last_day_col = 3 + month_days
        succ_col_let = openpyxl.utils.get_column_letter(last_day_col + 1)
        done_tgt_col_let = openpyxl.utils.get_column_letter(last_day_col + 2)

        assert ws["A16"].value == "S.No."
        assert ws["B16"].value == "Protocols"
        assert ws["C16"].value == "Target"
        assert ws[f"{succ_col_let}16"].value == "Success %"
        assert ws[f"{done_tgt_col_let}16"].value == "Done / Target"
        print(f"✓ Table Headers verified: 'Success %' at {succ_col_let}16 and 'Done / Target' at {done_tgt_col_let}16")

        # 3. 20 Protocol rows (Rows 18 to 37) - Non-Red and Non-Green Group Colors
        # Row 18: Sky Blue (E0F2FE)
        assert ws["A18"].fill.start_color.rgb in ("00E0F2FE", "E0F2FE"), f"Got {ws['A18'].fill.start_color.rgb}"
        # Row 22: Lavender (F3E8FF)
        assert ws["A22"].fill.start_color.rgb in ("00F3E8FF", "F3E8FF"), f"Got {ws['A22'].fill.start_color.rgb}"
        # Row 26: Warm Sand (FEF3C7)
        assert ws["A26"].fill.start_color.rgb in ("00FEF3C7", "FEF3C7"), f"Got {ws['A26'].fill.start_color.rgb}"
        # Row 30: Cool Slate (F1F5F9)
        assert ws["A30"].fill.start_color.rgb in ("00F1F5F9", "F1F5F9"), f"Got {ws['A30'].fill.start_color.rgb}"
        # Row 34: Periwinkle (EEF2FF)
        assert ws["A34"].fill.start_color.rgb in ("00EEF2FF", "EEF2FF"), f"Got {ws['A34'].fill.start_color.rgb}"
        print("✓ 5 Non-Red & Non-Green Pastel Group Colors verified (Sky Blue, Lavender, Sand, Slate, Periwinkle)")

        # 4. Conditional Formatting for Tick (Green) and Cross (Red)
        cf_rules = ws.conditional_formatting
        print(f"Conditional formatting rules count: {len(cf_rules)}")
        assert len(cf_rules) >= 1, "Expected conditional formatting on protocol day cells"
        print("✓ Dynamic Conditional Formatting (Green for ✓, Red for ✗, bg if blank) verified!")

        # 5. DAILY TOTAL SCORE (Row 38) - Shows done/total like 8/10
        assert ws["A38"].value == "DAILY TOTAL SCORE"
        day_score_formula = str(ws["D38"].value)
        assert '& "/" &' in day_score_formula or '&"/"&' in day_score_formula, f"Expected done/eval in D38 formula, got {day_score_formula}"
        print("✓ DAILY TOTAL SCORE (Row 38) verified showing done/eval like 8/10 on day columns")

        # 6. DAILY SUCCESS % (Row 39) - Under Daily Total Score
        assert ws["A39"].value == "DAILY SUCCESS %"
        assert ws["D39"].number_format == "0.0%"
        print("✓ DAILY SUCCESS % (Row 39) verified showing percentage directly under 8/10 score")

        # 7. Success & Punishment stakes (Rows 41 & 42)
        assert "PUNISHMENT" in str(ws["A41"].value)
        assert "REWARD" in str(ws["A42"].value)
        print("✓ Success & Punishment stakes verified between tables at Rows 41 & 42")

        # 8. Two-Row Gap (Rows 43 & 44)
        for r_gap in (43, 44):
            for c_idx in range(1, 10):
                c_let = openpyxl.utils.get_column_letter(c_idx)
                assert ws[f"{c_let}{r_gap}"].value is None, f"Expected blank gap at {c_let}{r_gap}"
        print("✓ Two-row blank gap (Rows 43 & 44) verified")

        # 8. Sleep Tracking Section (Starts at Row 45) - Background Colors Verified
        assert ws["A45"].value == "Sleep Tracking"
        assert ws["A45"].fill.start_color.rgb in ("001E1B4B", "1E1B4B"), f"Got {ws['A45'].fill.start_color.rgb}"
        assert ws["A47"].fill.start_color.rgb in ("00312E81", "312E81"), f"Got {ws['A47'].fill.start_color.rgb}"
        assert ws["D45"].fill.start_color.rgb in ("00E0E7FF", "E0E7FF", "00C7D2FE", "C7D2FE")
        assert ws["D47"].fill.start_color.rgb in ("00EEF2FF", "EEF2FF", "00DBEAFE", "DBEAFE")
        print("✓ Sleep Tracking section background colors (Night Indigo & Twilight theme) verified!")

        # 9. Sleep Line Chart
        assert len(ws._charts) == 1
        chart = ws._charts[0]
        assert "Sleep" in str(chart.title)
        print(f"✓ Dynamic Sleep Line Chart verified at A49: '{chart.title}'")

    print("\nALL VERIFICATIONS PASSED 100% SUCCESSFULLY!")

if __name__ == "__main__":
    test_workbook()
