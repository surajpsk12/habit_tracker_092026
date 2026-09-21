import sys
import calendar
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

def test_workbook():
    fn = "Habit_Tracker_2026_2030.xlsx"
    wb = openpyxl.load_workbook(fn, data_only=False)
    sheet_names = wb.sheetnames
    print(f"\n============================\nVerifying {fn}:")
    print(f"Total Sheets: {len(sheet_names)}")
    
    assert len(sheet_names) == 55, f"Expected 55 sheets, got {len(sheet_names)}"
    assert sheet_names[0] == "Master Dashboard"
    assert sheet_names[1] == "Quarterly Hub"
    assert sheet_names[2] == "Habit Library"
    assert sheet_names[3] == "Sep 2026"
    assert sheet_names[-1] == "Dec 2030"
    print(f"Sheet 1: {sheet_names[0]}, Sheet 2: {sheet_names[1]}, Sheet 3: {sheet_names[2]}, First Month: {sheet_names[3]}, Last Month: {sheet_names[-1]}")

    # 1. Verify Master Dashboard
    ws_master = wb["Master Dashboard"]
    print("\n--- Checking Master Dashboard ---")
    assert "Master Dashboard" in str(ws_master["A2"].value)
    assert ws_master["A5"].value == "GLOBAL SUCCESS RATE"
    assert ws_master["C5"].value == "ALL-TIME CHECKS"
    assert ws_master["E5"].value == "TOTAL TARGET DAYS"
    assert ws_master["G5"].value == "PEAK MONTH ON RECORD"
    assert ws_master["I5"].value == "GLOBAL AVG SLEEP"
    assert ws_master["A4"].hyperlink is not None # Link to Quarterly Hub
    assert ws_master["C4"].hyperlink is not None # Link to Habit Library
    assert len(ws_master._charts) == 2
    assert ws_master.protection.sheet == True
    print("✓ Master Dashboard verified: 5 Hero Cards, Navigation links, 2 Master Charts & Protection")

    # 2. Verify Quarterly Hub
    ws_q = wb["Quarterly Hub"]
    print("\n--- Checking Quarterly Hub ---")
    assert "Quarterly Sprints" in str(ws_q["A2"].value)
    assert ws_q["A4"].hyperlink is not None # Link to Master Dashboard
    assert ws_q["C4"].hyperlink is not None # Link to Habit Library
    assert ws_q["A6"].value == "Quarter"
    assert ws_q["A7"].value == "Q4 2026"
    assert "='Sep 2026'!I92" in str(ws_q["D7"].value)
    assert "='Sep 2026'!G92" in str(ws_q["E7"].value)
    assert ws_q["I7"].hyperlink is not None # Link to Sep 2026
    assert len(ws_q._charts) == 1 # 17-Quarter Trajectory Chart
    assert "90-DAY SPRINT" in str(ws_q["A26"].value)
    assert "HABIT EVOLUTION" in str(ws_q["A33"].value)
    assert ws_q.protection.sheet == True
    print("✓ Quarterly Hub verified: 17 Quarters Scorecard, formulas linked, Quarterly Chart, OKRs & Habit Evolution")

    # 3. Verify Habit Library
    ws_lib = wb["Habit Library"]
    print("\n--- Checking Habit Library ---")
    assert "Habit Mastery Playbook" in str(ws_lib["A2"].value)
    assert ws_lib["A4"].hyperlink is not None
    assert ws_lib["A6"].value == "S.No."
    assert ws_lib["B6"].value == "Life Domain"
    assert ws_lib["C6"].value == "Protocol Name"
    # Check that 60 protocols exist
    assert ws_lib["A66"].value == 60
    assert "ATOMIC HABITS" in str(ws_lib["A68"].value)
    assert ws_lib.protection.sheet == True
    print("✓ Habit Library verified: 60 Curated Protocols across 5 Domains & Atomic Habits framework")

    # 4. Verify all 52 Monthly Sheets
    for name in sheet_names[3:]:
        ws = wb[name]
        print(f"\n--- Checking sheet: {name} ---")

        # Row 1 Navigation Bar (Has 3 hub links + years + prev/next)
        assert "Master Dashboard" in str(ws["A1"].value)
        assert ws["A1"].hyperlink is not None
        assert "Quarterly Hub" in str(ws["C1"].value)
        assert ws["C1"].hyperlink is not None
        assert "Habit Library" in str(ws["D1"].value)
        assert ws["D1"].hyperlink is not None
        assert ws["E1"].value == "2026"
        assert ws["J1"].value == "◀ Prev"
        assert ws["K1"].value == "Next ▶"
        print("✓ Top Navigation Bar (Master, Quarterly, Library, Years, Prev/Next) verified")

        assert ws["A2"].value == "Habit Tracker"
        assert ws["A5"].value == "MONTHLY TARGETS"
        assert ws["A11"].value == "NOTES:"

        # Input cells unlocked
        assert ws["A6"].protection.locked == False # Target input
        assert ws["A12"].protection.locked == False # Notes input
        assert ws["B18"].protection.locked == False # Habit input
        assert ws["C18"].protection.locked == False # Category input
        assert ws["D18"].protection.locked == False # Target input
        assert ws["E18"].protection.locked == False # Day check input
        assert ws["E47"].protection.locked == False # Sleep hours input
        assert ws["A41"].protection.locked == False # Punishment input
        assert ws["A42"].protection.locked == False # Reward input
        assert ws["A96"].protection.locked == False # Retrospective input

        # Formula cells locked
        m_str, y_str = name.split()
        m_idx = list(calendar.month_abbr).index(m_str)
        month_days = calendar.monthrange(int(y_str), m_idx)[1]
        last_day_col = 4 + month_days
        succ_col_let = openpyxl.utils.get_column_letter(last_day_col + 1)
        done_tgt_col_let = openpyxl.utils.get_column_letter(last_day_col + 2)
        status_col_let = openpyxl.utils.get_column_letter(last_day_col + 3)

        assert ws[f"{succ_col_let}18"].protection.locked == True
        assert ws[f"{done_tgt_col_let}18"].protection.locked == True
        assert ws[f"{status_col_let}18"].protection.locked == True

        # Sheet Protection & Print Setup
        assert ws.protection.sheet == True
        assert ws.page_setup.orientation == ws.ORIENTATION_LANDSCAPE
        assert str(ws.page_setup.paperSize) in ('9', ws.PAPERSIZE_A4)
        assert ws.page_setup.fitToWidth == 1

        # Check total charts per monthly sheet
        assert len(ws._charts) == 3

    print("\n🎉 ALL 55 SHEETS (MASTER + QUARTERLY + LIBRARY + 52 MONTHS) VERIFIED 100% SUCCESSFULLY!")

if __name__ == "__main__":
    test_workbook()
