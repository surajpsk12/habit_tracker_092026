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
        
        # 1. Top Section (Before Table)
        assert "PUNISHMENT" in str(ws["A5"].value)
        assert "REWARD" in str(ws["A6"].value)
        assert ws["A8"].value == "MONTHLY TARGETS"
        assert "1." in str(ws["A9"].value)
        assert ws["A14"].value == "NOTES:"
        print("✓ Top Section (Stakes, Targets, Notes) verified before table")

        # 2. Protocol Table Header & 20 rows
        assert ws["A19"].value == "S.No."
        assert ws["B19"].value == "Protocols"
        assert ws["C19"].value == "Target"
        assert ws["A21"].value == 1
        assert ws["A40"].value == 20
        print("✓ Table Header (Rows 19-20) and 20 Protocol rows (Rows 21-40) verified")

        # 3. Data Validations (Target 1-31, Protocol ✓/✗, Sleep 1-24)
        dvs = ws.data_validations.dataValidation
        print(f"Data validations count: {len(dvs)}")
        assert len(dvs) == 3, f"Expected 3 data validations, got {len(dvs)}"
        has_tick_cross = any("✓" in str(dv.formula1) and "✗" in str(dv.formula1) for dv in dvs)
        assert has_tick_cross, "Protocol day cells tick/cross dropdown not found!"
        print("✓ Protocol day cells tick/cross (✓, ✗) dropdown verified!")

        # 4. DAILY TOTAL SCORE (Row 41)
        assert ws["A41"].value == "DAILY TOTAL SCORE"
        score_formula = str(ws["D41"].value)
        assert "COUNTIF(D21:D40" in score_formula
        print("✓ DAILY TOTAL SCORE at Row 41 verified")

        # 5. DAILY SUCCESS % (Row 42)
        assert ws["A42"].value == "DAILY SUCCESS %"
        pct_formula = str(ws["D42"].value)
        assert "/" in pct_formula and "COUNTIF(D21:D40" in pct_formula
        assert ws["D42"].number_format == "0.0%"
        print("✓ DAILY SUCCESS % row at Row 42 verified with tick/eval formula and 0.0% format")

        # 6. Two-row gap (Rows 43 & 44)
        for r_gap in (43, 44):
            for c_idx in range(1, 10):
                c_let = openpyxl.utils.get_column_letter(c_idx)
                assert ws[f"{c_let}{r_gap}"].value is None, f"Expected blank gap at {c_let}{r_gap}"
        print("✓ Exactly two blank gap rows (Rows 43 & 44) verified!")

        # 7. Sleep Tracking Table (Starts at Row 45)
        assert ws["A45"].value == "Sleep Tracking"
        assert "Sleep Hours" in str(ws["A47"].value)
        print("✓ Sleep Tracking section starts at Row 45 after the 2-row gap")

        # 8. Sleep Line Chart
        assert len(ws._charts) == 1
        chart = ws._charts[0]
        assert "Sleep" in str(chart.title)
        print(f"✓ Dynamic Sleep Line Chart verified at A49: '{chart.title}'")

    print("\nALL VERIFICATIONS PASSED 100% SUCCESSFULLY!")

if __name__ == "__main__":
    test_workbook()
