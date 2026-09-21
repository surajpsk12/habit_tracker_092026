import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

def test_workbook():
    for fn in ["Habit_Tracker_Sep_Dec_2026.xlsx", "Habit_Tracker_2026_2027.xlsx"]:
        wb = openpyxl.load_workbook(fn, data_only=False)
        sheet_names = wb.sheetnames
        print(f"\nVerifying {fn}:")
        print(f"Total Sheets: {len(sheet_names)}")
        print(f"Sheets: {sheet_names}")
        
        expected_sheets = ["Sep 2026", "Oct 2026", "Nov 2026", "Dec 2026"]
        assert sheet_names == expected_sheets, f"Expected {expected_sheets}, got {sheet_names}"

        for name in expected_sheets:
            ws = wb[name]
            print(f"\n--- Checking sheet: {name} ---")
            assert ws["A2"].value == "Habit Tracker"
            print(f"Title: {ws['A2'].value} | Subtitle: {ws['A3'].value}")
            
            # Check S.No, Protocols, Target
            assert ws["A5"].value == "S.No."
            assert ws["B5"].value == "Protocols"
            assert ws["C5"].value == "Target"
            
            # Check Days
            month_days = 30 if "Sep" in name or "Nov" in name else 31
            last_day_col = 3 + month_days
            success_col = last_day_col + 1
            success_col_letter = openpyxl.utils.get_column_letter(success_col)
            assert ws[f"{success_col_letter}5"].value == "Success %"

            # Check 15 protocol rows S.No
            for i in range(1, 16):
                row = 6 + i
                assert ws[f"A{row}"].value == i
                # Check fill
                fill = ws[f"A{row}"].fill.start_color.rgb
                assert fill is not None, f"Row {row} missing fill"

            # Check DAILY TOTAL SCORE
            assert ws["A22"].value == "DAILY TOTAL SCORE"

            # Check Stakes / Reward & Punishment
            assert "PUNISHMENT" in str(ws["A24"].value)
            assert "REWARD" in str(ws["A25"].value)

            # Check Sleep Tracking
            assert ws["A27"].value == "Sleep Tracking"
            for idx, lbl in enumerate(["9hrs", "8hrs", "7hrs", "6hrs", "5hrs"]):
                assert ws[f"A{29+idx}"].value == lbl

            # Check Notes
            assert ws["A35"].value == "NOTES:"

            # Check Monthly Targets
            assert ws["A40"].value == "MONTHLY TARGETS"
            assert "1." in str(ws["A41"].value)
            print(f"Sheet {name} layout and fields 100% verified!")

    print("\nALL VERIFICATIONS PASSED SUCCESSFULLY!")

if __name__ == "__main__":
    test_workbook()
