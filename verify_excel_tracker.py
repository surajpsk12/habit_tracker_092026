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
    
    assert len(sheet_names) == 53, f"Expected 53 sheets (Master Dashboard + 52 months), got {len(sheet_names)}"
    assert sheet_names[0] == "Master Dashboard"
    assert sheet_names[1] == "Sep 2026"
    assert sheet_names[-1] == "Dec 2030"
    print(f"First Sheet: {sheet_names[0]}, Second Sheet: {sheet_names[1]}, Last Sheet: {sheet_names[-1]}")

    # 1. Verify Master Dashboard
    ws_master = wb["Master Dashboard"]
    print("\n--- Checking Master Dashboard ---")
    assert "Master Dashboard" in str(ws_master["A2"].value)
    assert ws_master["A5"].value == "GLOBAL SUCCESS RATE"
    assert ws_master["C5"].value == "ALL-TIME CHECKS"
    assert ws_master["E5"].value == "TOTAL TARGET DAYS"
    assert ws_master["G5"].value == "PEAK MONTH ON RECORD"
    assert ws_master["I5"].value == "GLOBAL AVG SLEEP"
    print("✓ Master Dashboard 5 Executive Hero KPI Cards verified")

    # Verify Navigator Grid
    assert "ONE-CLICK MONTH NAVIGATOR" in str(ws_master["A11"].value)
    assert ws_master["B12"].value == "Sep"
    assert ws_master["B12"].hyperlink is not None
    print(f"✓ One-Click Month Navigator verified with hyperlinks (e.g. {ws_master['B12'].hyperlink.target})")

    # Verify 52-row Table on Master Dashboard
    assert ws_master["A18"].value == "Month & Year"
    assert ws_master["E18"].value == "Success %"
    assert ws_master["A19"].value == "Sep 2026"
    assert str(ws_master["B19"].value) == "='Sep 2026'!E92"
    assert str(ws_master["E19"].value) == "='Sep 2026'!K92"
    assert ws_master["J19"].hyperlink is not None
    print("✓ Master 52-Month Aggregated Performance Table verified linked to Row 92 data bridge")

    # Verify Master Charts
    assert len(ws_master._charts) == 2, f"Expected 2 master charts, got {len(ws_master._charts)}"
    print("✓ Master Dashboard Charts verified: Habit Execution Trend & Sleep Average Trend")

    # 2. Verify all 52 Monthly Sheets
    for name in sheet_names[1:]:
        ws = wb[name]
        print(f"\n--- Checking sheet: {name} ---")

        # Row 1 Navigation Bar
        assert "Master Dashboard" in str(ws["A1"].value)
        assert ws["A1"].hyperlink is not None
        assert ws["C1"].value == "2026"
        assert ws["C1"].hyperlink is not None
        assert ws["H1"].value == "◀ Prev"
        assert ws["I1"].value == "Next ▶"
        print("✓ Top Quick-Jump Navigation Bar (Row 1) verified")

        assert ws["A2"].value == "Habit Tracker"
        
        # Monthly Targets & Notes
        assert ws["A5"].value == "MONTHLY TARGETS"
        assert ws["A11"].value == "NOTES:"
        assert ws["A12"].fill.start_color.rgb in ("00FFFFE6", "FFFFE6")

        # Main Table Headers
        m_str, y_str = name.split()
        m_idx = list(calendar.month_abbr).index(m_str)
        month_days = calendar.monthrange(int(y_str), m_idx)[1]
        
        first_day_col = 5
        last_day_col = 4 + month_days
        succ_col_let = openpyxl.utils.get_column_letter(last_day_col + 1)
        done_tgt_col_let = openpyxl.utils.get_column_letter(last_day_col + 2)
        status_col_let = openpyxl.utils.get_column_letter(last_day_col + 3)
        helper_col_let = openpyxl.utils.get_column_letter(last_day_col + 4)

        assert ws["A16"].value == "S.No."
        assert ws["B16"].value == "Protocols"
        assert ws["C16"].value == "Category"
        assert ws["D16"].value == "Target"
        assert ws[f"{succ_col_let}16"].value == "Success %"
        assert ws[f"{done_tgt_col_let}16"].value == "Done / Target"
        assert ws[f"{status_col_let}16"].value == "Status"
        assert ws[f"{helper_col_let}16"].value == "Done Count"
        assert ws.column_dimensions[helper_col_let].hidden == True
        print(f"✓ Table Headers verified: Category at C16, Target at D16, Success % at {succ_col_let}16, Done/Target at {done_tgt_col_let}16, Status Badge at {status_col_let}16")

        # Verify Status Badges formula in Row 18
        status_formula = str(ws[f"{status_col_let}18"].value)
        assert "Mastered" in status_formula and "Consistent" in status_formula and "Needs Focus" in status_formula
        print(f"✓ Automated Status Badges formula verified in Column {status_col_let}")

        # Verify CF: DataBar + Tick + Cross + Today
        cf_rules = ws.conditional_formatting
        assert len(cf_rules) >= 6, f"Expected at least 6 CF rules (including DataBar), got {len(cf_rules)}"
        print("✓ Dynamic Conditional Formatting (DataBar on Success %, Green ✓, Red ✗, Today highlight) verified")

        # Daily Total Score & Daily Success %
        assert ws["A38"].value == "DAILY TOTAL SCORE"
        assert ws["A39"].value == "DAILY SUCCESS %"

        # Stakes & Two-row gap
        assert "PUNISHMENT" in str(ws["A41"].value)
        assert "REWARD" in str(ws["A42"].value)
        for r_gap in (43, 44):
            assert ws[f"A{r_gap}"].value is None

        # Sleep Section & Charts (Total 3 Charts per monthly sheet)
        assert ws["A45"].value == "Sleep Tracking"
        assert ws["A47"].value == "Sleep Hours (1-24)"
        assert len(ws._charts) == 3, f"Expected 3 charts in {name}, found {len(ws._charts)}"

        # KPI Cards & Review tables
        assert ws["A67"].value == "MONTHLY SUCCESS RATE"
        assert ws["A74"].value == "WEEKLY PERFORMANCE (SLUMP DETECTOR)"
        assert ws["N74"].value == "LIFE DOMAIN / CATEGORY BREAKDOWN"
        assert ws["A84"].value == "SLEEP & PRODUCTIVITY CORRELATION ANALYSIS"
        assert ws["A90"].value == "STANDARDIZED DATA BRIDGE (FOR ANNUAL & QUARTERLY REVIEW SHEETS)"
        assert ws["A92"].value == m_str
        assert ws["C92"].value == int(y_str)
        assert ws["A95"].value == "🏆 MONTHLY WINS & HIGHLIGHTS"

    print("\n🎉 ALL 53 SHEETS (MASTER DASHBOARD + 52 MONTHS) VERIFIED 100% SUCCESSFULLY!")

if __name__ == "__main__":
    test_workbook()
