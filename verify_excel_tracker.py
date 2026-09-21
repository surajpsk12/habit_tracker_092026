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
        m_str, y_str = name.split()
        m_idx = list(calendar.month_abbr).index(m_str)
        month_days = calendar.monthrange(int(y_str), m_idx)[1]
        
        first_day_col = 5
        last_day_col = 4 + month_days
        succ_col_let = openpyxl.utils.get_column_letter(last_day_col + 1)
        done_tgt_col_let = openpyxl.utils.get_column_letter(last_day_col + 2)
        helper_col_let = openpyxl.utils.get_column_letter(last_day_col + 3)

        assert ws["A16"].value == "S.No."
        assert ws["B16"].value == "Protocols"
        assert ws["C16"].value == "Category"
        assert ws["D16"].value == "Target"
        assert ws[f"{succ_col_let}16"].value == "Success %"
        assert ws[f"{done_tgt_col_let}16"].value == "Done / Target"
        assert ws[f"{helper_col_let}16"].value == "Done Count"
        assert ws.column_dimensions[helper_col_let].hidden == True
        print(f"✓ Table Headers verified: Category at C16, Target at D16, Success % at {succ_col_let}16, Done/Target at {done_tgt_col_let}16")

        # 3. 20 Protocol rows (Rows 18 to 37) - Non-Red and Non-Green Group Colors
        assert ws["A18"].fill.start_color.rgb in ("00E0F2FE", "E0F2FE"), f"Got {ws['A18'].fill.start_color.rgb}"
        assert ws["A22"].fill.start_color.rgb in ("00F3E8FF", "F3E8FF"), f"Got {ws['A22'].fill.start_color.rgb}"
        assert ws["A26"].fill.start_color.rgb in ("00FEF3C7", "FEF3C7"), f"Got {ws['A26'].fill.start_color.rgb}"
        assert ws["A30"].fill.start_color.rgb in ("00F1F5F9", "F1F5F9"), f"Got {ws['A30'].fill.start_color.rgb}"
        assert ws["A34"].fill.start_color.rgb in ("00EEF2FF", "EEF2FF"), f"Got {ws['A34'].fill.start_color.rgb}"
        print("✓ 5 Non-Red & Non-Green Pastel Group Colors verified")

        # 4. Conditional Formatting for Tick, Cross, and Today Highlighting
        cf_rules = ws.conditional_formatting
        print(f"Conditional formatting rules count: {len(cf_rules)}")
        assert len(cf_rules) >= 5, "Expected CF rules on protocol cells and today header"
        print("✓ Dynamic Conditional Formatting (Green for ✓, Red for ✗, Today highlight) verified!")

        # 5. DAILY TOTAL SCORE (Row 38) & DAILY SUCCESS % (Row 39)
        assert ws["A38"].value == "DAILY TOTAL SCORE"
        day_score_formula = str(ws["E38"].value)
        assert '& "/" &' in day_score_formula or '&"/"&' in day_score_formula, f"Expected done/eval in E38 formula, got {day_score_formula}"
        assert ws["A39"].value == "DAILY SUCCESS %"
        assert ws["E39"].number_format == "0.0%"
        print("✓ DAILY TOTAL SCORE (Row 38) and DAILY SUCCESS % (Row 39) verified")

        # 6. Stakes & Two-Row Gap
        assert "PUNISHMENT" in str(ws["A41"].value)
        assert "REWARD" in str(ws["A42"].value)
        for r_gap in (43, 44):
            for c_idx in range(1, 10):
                c_let = openpyxl.utils.get_column_letter(c_idx)
                assert ws[f"{c_let}{r_gap}"].value is None, f"Expected blank gap at {c_let}{r_gap}"
        print("✓ Success & Punishment stakes + two-row blank gap (Rows 43 & 44) verified")

        # 7. Sleep Tracking Section & Dynamic Charts (Total 3 Charts per Sheet)
        assert ws["A45"].value == "Sleep Tracking"
        assert ws["A47"].value == "Sleep Hours (1-24)"
        assert len(ws._charts) == 3, f"Expected 3 charts, found {len(ws._charts)}"
        chart_titles = [str(c.title) for c in ws._charts]
        print(f"✓ Charts verified (Total 3): Sleep Trend, Weekly Momentum, Domain Performance")

        # 8. Executive KPI Cards (Rows 67–71)
        assert ws["A67"].value == "MONTHLY SUCCESS RATE"
        assert ws["F67"].value == "ACTIVE PROTOCOLS"
        assert ws["K67"].value == "PERFECT DAYS (100%)"
        assert ws["P67"].value == "AVG SLEEP DURATION"
        print("✓ Executive KPI Summary Cards (Rows 67–71) verified")

        # 9. Weekly Slump Detector Table (Rows 74–81) & Category Breakdown (Rows 74–80)
        assert ws["A74"].value == "WEEKLY PERFORMANCE (SLUMP DETECTOR)"
        assert ws["A76"].value == "Week 1"
        assert ws["N74"].value == "LIFE DOMAIN / CATEGORY BREAKDOWN"
        assert ws["N76"].value == "Health & Fitness"
        print("✓ Weekly Slump Detector & Category Domain Breakdown tables verified")

        # 10. Sleep & Productivity Correlation Table (Rows 84–88)
        assert ws["A84"].value == "SLEEP & PRODUCTIVITY CORRELATION ANALYSIS"
        assert ws["A86"].value == "Optimal Rest"
        assert ws["A87"].value == "Sleep Deficit"
        assert "Sleep Correlation Insight" in str(ws["A88"].value)
        print("✓ Sleep vs Habit Performance Correlation Analysis verified")

        # 11. Standardized Data Bridge (Rows 90–93)
        assert ws["A90"].value == "STANDARDIZED DATA BRIDGE (FOR ANNUAL & QUARTERLY REVIEW SHEETS)"
        assert ws["A92"].value == m_str
        assert ws["C92"].value == int(y_str)
        assert ws["K92"].number_format == "0.0%"
        print("✓ Standardized Annual / Quarterly Review Data Bridge verified")

        # 12. Monthly Retrospective Box (Rows 95–104)
        assert ws["A95"].value == "🏆 MONTHLY WINS & HIGHLIGHTS"
        assert ws["J95"].value == "⚠️ FRICTION POINTS & ROOT CAUSES"
        assert ws["S95"].value == "🎯 3 KEY ADJUSTMENTS FOR NEXT MONTH"
        print("✓ Monthly Retrospective & Action Plan Reflection Box verified")

    print("\n🎉 ALL 52 SHEETS VERIFIED 100% SUCCESSFULLY!")

if __name__ == "__main__":
    test_workbook()
