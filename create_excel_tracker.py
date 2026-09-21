import calendar
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.chart import LineChart, BarChart, Reference

def generate_exact_tracker_excel():
    wb = openpyxl.Workbook()
    # Active sheet becomes Master Dashboard
    ws_master = wb.active
    ws_master.title = "Master Dashboard"

    # Typography & Styles
    FONT_FAMILY = "Arial"
    
    # Border definitions
    THIN_BORDER_GRAY = Side(border_style="thin", color="000000")
    MEDIUM_BORDER = Side(border_style="medium", color="000000")
    LIGHT_BORDER = Side(border_style="thin", color="D1D5DB")
    
    BORDER_STANDARD = Border(
        left=THIN_BORDER_GRAY, right=THIN_BORDER_GRAY,
        top=THIN_BORDER_GRAY, bottom=THIN_BORDER_GRAY
    )
    BORDER_SUNDAY = Border(
        left=MEDIUM_BORDER, right=MEDIUM_BORDER,
        top=THIN_BORDER_GRAY, bottom=THIN_BORDER_GRAY
    )
    BORDER_NOTE = Border(
        left=LIGHT_BORDER, right=LIGHT_BORDER,
        top=LIGHT_BORDER, bottom=LIGHT_BORDER
    )
    BORDER_NAV_PILL = Border(
        left=Side(border_style="thin", color="475569"),
        right=Side(border_style="thin", color="475569"),
        top=Side(border_style="thin", color="475569"),
        bottom=Side(border_style="thin", color="475569")
    )
    BORDER_CARD_GREEN = Border(
        left=Side(border_style="medium", color="16A34A"),
        right=Side(border_style="medium", color="16A34A"),
        top=Side(border_style="medium", color="16A34A"),
        bottom=Side(border_style="medium", color="16A34A")
    )
    BORDER_CARD_BLUE = Border(
        left=Side(border_style="medium", color="2563EB"),
        right=Side(border_style="medium", color="2563EB"),
        top=Side(border_style="medium", color="2563EB"),
        bottom=Side(border_style="medium", color="2563EB")
    )
    BORDER_CARD_AMBER = Border(
        left=Side(border_style="medium", color="D97706"),
        right=Side(border_style="medium", color="D97706"),
        top=Side(border_style="medium", color="D97706"),
        bottom=Side(border_style="medium", color="D97706")
    )
    BORDER_CARD_PURPLE = Border(
        left=Side(border_style="medium", color="7C3AED"),
        right=Side(border_style="medium", color="7C3AED"),
        top=Side(border_style="medium", color="7C3AED"),
        bottom=Side(border_style="medium", color="7C3AED")
    )
    BORDER_CARD_INDIGO = Border(
        left=Side(border_style="medium", color="4F46E5"),
        right=Side(border_style="medium", color="4F46E5"),
        top=Side(border_style="medium", color="4F46E5"),
        bottom=Side(border_style="medium", color="4F46E5")
    )

    # 5 Non-Red & Non-Green Pastel Group Colors for 20 protocol rows
    GROUP_COLORS_HEX = [
        "E0F2FE",  # Soft Sky Blue - Rows 1-4
        "F3E8FF",  # Soft Lavender - Rows 5-8
        "FEF3C7",  # Soft Warm Sand/Amber - Rows 9-12
        "F1F5F9",  # Soft Cool Slate - Rows 13-16
        "EEF2FF",  # Soft Periwinkle/Indigo - Rows 17-20
    ]
    
    SUNDAY_GROUP_COLORS_HEX = [
        "BAE6FD",
        "E9D5FF",
        "FDE68A",
        "CBD5E1",
        "C7D2FE",
    ]

    # Dynamic Green & Red Fills & Fonts for Selected Tick / Cross
    CF_TICK_FILL = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    CF_TICK_FONT = Font(name=FONT_FAMILY, size=9, bold=True, color="166534")

    CF_CROSS_FILL = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    CF_CROSS_FONT = Font(name=FONT_FAMILY, size=9, bold=True, color="991B1B")

    # Today Highlight Fill & Font (Gold / Amber Glow)
    CF_TODAY_FILL = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    CF_TODAY_FONT = Font(name=FONT_FAMILY, size=9, bold=True, color="854D0E")

    # Sleep Tracking Theme Colors
    SLEEP_MAIN_HDR_FILL = PatternFill(start_color="1E1B4B", end_color="1E1B4B", fill_type="solid")
    SLEEP_LBL_FILL = PatternFill(start_color="312E81", end_color="312E81", fill_type="solid")
    SLEEP_DAY_HDR_FILL = PatternFill(start_color="E0E7FF", end_color="E0E7FF", fill_type="solid")
    SLEEP_SUN_HDR_FILL = PatternFill(start_color="C7D2FE", end_color="C7D2FE", fill_type="solid")
    SLEEP_CELL_FILL = PatternFill(start_color="EEF2FF", end_color="EEF2FF", fill_type="solid")
    SLEEP_SUN_CELL_FILL = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")

    SUNDAY_HEADER_FILL = PatternFill(start_color="D2D2D2", end_color="D2D2D2", fill_type="solid")
    NOTES_FILL = PatternFill(start_color="FFFFE6", end_color="FFFFE6", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    # KPI Card Fills
    KPI_GREEN_FILL = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
    KPI_BLUE_FILL = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    KPI_AMBER_FILL = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    KPI_PURPLE_FILL = PatternFill(start_color="F3E8FF", end_color="F3E8FF", fill_type="solid")
    KPI_INDIGO_FILL = PatternFill(start_color="EEF2FF", end_color="EEF2FF", fill_type="solid")

    # Section Table Headers
    HDR_DARK_NAVY = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    HDR_SLATE = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    HDR_INDIGO = PatternFill(start_color="1E1B4B", end_color="1E1B4B", fill_type="solid")
    HDR_EMERALD = PatternFill(start_color="047857", end_color="047857", fill_type="solid")
    SUBHDR_GRAY = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    NAV_BAR_FILL = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    NAV_PILL_FILL = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    NAV_ACCENT_FILL = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid")

    QUOTES = [
        "Win the morning, win the day.",
        "Discipline equals freedom.",
        "Focus on the process.",
        "Stay consistent.",
        "Small daily improvements lead to stunning results.",
        "We are what we repeatedly do.",
        "Action is the foundational key to all success.",
        "Success is the sum of small efforts repeated daily.",
        "Energy flows where attention goes.",
        "Your future is created by what you do today.",
        "Mastery is a journey, not a destination.",
        "Consistency creates momentum."
    ]

    CATEGORIES = [
        "Health & Fitness",
        "Mindset & Learning",
        "Career & Focus",
        "Personal & Home",
        "Finance & Wealth",
        "Other"
    ]
    categories_dropdown_str = ",".join(CATEGORIES)

    # Timeline: September 2026 up to December 2030 (52 Months Total)
    target_months = []
    # 2026: Sep (9) to Dec (12)
    for m in range(9, 13):
        target_months.append((2026, m))
    # 2027 to 2030: Full years (1 to 12)
    for y in range(2027, 2031):
        for m in range(1, 13):
            target_months.append((y, m))

    weekdays_abbr = ['M', 'T', 'W', 'Th', 'F', 'Sa', 'Su']
    target_dropdown_str = ",".join(str(i) for i in range(1, 32))
    sleep_dropdown_str = ",".join(str(i) for i in range(1, 25))

    print(f"Generating Master Dashboard + {len(target_months)} monthly sheets with Full UI & Analytics (Sep 2026 to Dec 2030)...")

    month_titles = [f"{calendar.month_abbr[m]} {y}" for y, m in target_months]

    # =========================================================================
    # PART 1: MASTER ANNUAL DASHBOARD SHEET
    # =========================================================================
    ws_master.views.sheetView[0].showGridLines = True
    ws_master.page_setup.orientation = ws_master.ORIENTATION_LANDSCAPE
    ws_master.page_setup.paperSize = ws_master.PAPERSIZE_A4

    # Column dimensions for Master Dashboard
    ws_master.column_dimensions['A'].width = 13.0
    ws_master.column_dimensions['B'].width = 12.5
    ws_master.column_dimensions['C'].width = 12.0
    ws_master.column_dimensions['D'].width = 12.0
    ws_master.column_dimensions['E'].width = 13.0
    ws_master.column_dimensions['F'].width = 12.0
    ws_master.column_dimensions['G'].width = 13.0
    ws_master.column_dimensions['H'].width = 12.5
    ws_master.column_dimensions['I'].width = 17.0
    ws_master.column_dimensions['J'].width = 10.0
    ws_master.column_dimensions['K'].width = 4.0

    # Row 1: Padding
    ws_master.row_dimensions[1].height = 8

    # Row 2 & 3: Master Dashboard Header & Quote
    ws_master.row_dimensions[2].height = 26
    ws_master["A2"] = "Executive Habit & Performance Master Dashboard"
    ws_master["A2"].font = Font(name=FONT_FAMILY, size=18, bold=True, color="0F172A")
    ws_master["A2"].alignment = Alignment(horizontal="left", vertical="center")

    ws_master.merge_cells("F2:J2")
    m_badge = ws_master["F2"]
    m_badge.value = "Multi-Year Executive Suite (2026 – 2030)"
    m_badge.font = Font(name=FONT_FAMILY, size=11, bold=True, color="0284C7")
    m_badge.alignment = Alignment(horizontal="right", vertical="center")

    ws_master.row_dimensions[3].height = 16
    ws_master["A3"] = '"Long-term consistency trumps short-term intensity. Track, review, and elevate."'
    ws_master["A3"].font = Font(name=FONT_FAMILY, size=9.5, italic=True, color="64748B")
    ws_master["A3"].alignment = Alignment(horizontal="left", vertical="center")

    # Row 4: Spacer
    ws_master.row_dimensions[4].height = 6

    # Rows 5 to 9: 5 Master Executive Hero KPI Cards
    kpi_ranges = [
        ("A5:B5", "A6:B8", "A9:B9", "GLOBAL SUCCESS RATE", "=AVERAGEIF(E19:E70, \">0\")", "Overall Execution Rate", KPI_GREEN_FILL, BORDER_CARD_GREEN, "15803D", "166534", "0.0%"),
        ("C5:D5", "C6:D8", "C9:D9", "ALL-TIME CHECKS", "=SUM(D19:D70)", "Total Done Checks Logged", KPI_BLUE_FILL, BORDER_CARD_BLUE, "1D4ED8", "1E40AF", "#,##0"),
        ("E5:F5", "E6:F8", "E9:F9", "TOTAL TARGET DAYS", "=SUM(C19:C70)", "All Commitments Made", KPI_AMBER_FILL, BORDER_CARD_AMBER, "B45309", "92400E", "#,##0"),
        ("G5:H5", "G6:H8", "G9:H9", "PEAK MONTH ON RECORD", '=INDEX(A19:A70, MATCH(MAX(E19:E70), E19:E70, 0))', "Highest Consistency Month", KPI_PURPLE_FILL, BORDER_CARD_PURPLE, "6D28D9", "5B21B6", "@"),
        ("I5:J5", "I6:J8", "I9:J9", "GLOBAL AVG SLEEP", '=IFERROR(TEXT(AVERAGEIF(G19:G70, ">0"), "0.0") & " hrs", "-")', "Multi-Year Sleep Average", KPI_INDIGO_FILL, BORDER_CARD_INDIGO, "4338CA", "312E81", "@"),
    ]

    for top_r, mid_r, bot_r, title_text, val_formula, sub_text, fill_obj, border_obj, col_title, col_val, num_fmt in kpi_ranges:
        ws_master.merge_cells(top_r)
        c_top = ws_master[top_r.split(":")[0]]
        c_top.value = title_text
        c_top.font = Font(name=FONT_FAMILY, size=7.5, bold=True, color=col_title)
        c_top.alignment = Alignment(horizontal="center", vertical="center")

        ws_master.merge_cells(mid_r)
        c_mid = ws_master[mid_r.split(":")[0]]
        c_mid.value = val_formula
        c_mid.font = Font(name=FONT_FAMILY, size=17, bold=True, color=col_val)
        c_mid.alignment = Alignment(horizontal="center", vertical="center")
        c_mid.number_format = num_fmt

        ws_master.merge_cells(bot_r)
        c_bot = ws_master[bot_r.split(":")[0]]
        c_bot.value = sub_text
        c_bot.font = Font(name=FONT_FAMILY, size=7.0, italic=True, color=col_title)
        c_bot.alignment = Alignment(horizontal="center", vertical="center")

        # Apply fills and borders
        c_start_col = ord(top_r[0]) - ord('A') + 1
        c_end_col = ord(top_r[3]) - ord('A') + 1
        for r_k in range(5, 10):
            for c_k in range(c_start_col, c_end_col + 1):
                c_cell = ws_master.cell(row=r_k, column=c_k)
                c_cell.fill = fill_obj
                c_cell.border = border_obj

    # Row 10: Spacer
    ws_master.row_dimensions[10].height = 6

    # Rows 11 to 16: Interactive One-Click Multi-Year Navigator Grid
    ws_master.merge_cells("A11:J11")
    nav_title = ws_master["A11"]
    nav_title.value = "⚡ ONE-CLICK MONTH NAVIGATOR (CLICK ANY MONTH TO JUMP DIRECTLY)"
    nav_title.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="FFFFFF")
    nav_title.fill = HDR_SLATE
    nav_title.alignment = Alignment(horizontal="center", vertical="center")
    for c in range(1, 11):
        ws_master.cell(row=11, column=c).fill = HDR_SLATE
        ws_master.cell(row=11, column=c).border = BORDER_STANDARD

    year_nav_layout = [
        ("2026", range(9, 13), 12),
        ("2027", range(1, 13), 13),
        ("2028", range(1, 13), 14),
        ("2029", range(1, 13), 15),
        ("2030", range(1, 13), 16),
    ]

    for y_str, m_range, r_nav in year_nav_layout:
        ws_master.row_dimensions[r_nav].height = 14
        ws_master[f"A{r_nav}"] = f"Year {y_str}:"
        ws_master[f"A{r_nav}"].font = Font(name=FONT_FAMILY, size=8, bold=True, color="1E293B")
        ws_master[f"A{r_nav}"].fill = SUBHDR_GRAY
        ws_master[f"A{r_nav}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_master[f"A{r_nav}"].border = BORDER_STANDARD

        for idx, m_num in enumerate(m_range):
            col_target = 2 + idx
            c_let = get_column_letter(col_target)
            m_abbr = calendar.month_abbr[m_num]
            target_sheet_name = f"{m_abbr} {y_str}"
            
            nav_cell = ws_master[f"{c_let}{r_nav}"]
            nav_cell.value = m_abbr
            nav_cell.hyperlink = f"#'{target_sheet_name}'!A1"
            nav_cell.font = Font(name=FONT_FAMILY, size=7.5, bold=True, color="0284C7", underline="single")
            nav_cell.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
            nav_cell.alignment = Alignment(horizontal="center", vertical="center")
            nav_cell.border = BORDER_STANDARD

    # Row 17: Spacer
    ws_master.row_dimensions[17].height = 8

    # Rows 18 to 70: 52-Month Aggregated Performance Table
    dash_headers = [
        "Month & Year", "Active Habits", "Target Days", "Done Checks", "Success %",
        "Perfect Days", "Avg Sleep (hrs)", "Best Week", "Top Life Domain", "Direct Link"
    ]
    ws_master.row_dimensions[18].height = 16
    for idx, h_text in enumerate(dash_headers):
        cell = ws_master.cell(row=18, column=idx + 1)
        cell.value = h_text
        cell.font = Font(name=FONT_FAMILY, size=8, bold=True, color="FFFFFF")
        cell.fill = HDR_DARK_NAVY
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER_STANDARD

    for r_idx, sheet_name in enumerate(month_titles):
        r_table = 19 + r_idx
        ws_master.row_dimensions[r_table].height = 14

        ws_master[f"A{r_table}"] = sheet_name
        ws_master[f"A{r_table}"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws_master[f"A{r_table}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_master[f"A{r_table}"].border = BORDER_STANDARD

        # Formula references to the Standardized Data Bridge at Row 92 of each sheet
        ws_master[f"B{r_table}"] = f"='{sheet_name}'!E92"
        ws_master[f"C{r_table}"] = f"='{sheet_name}'!G92"
        ws_master[f"D{r_table}"] = f"='{sheet_name}'!I92"
        ws_master[f"E{r_table}"] = f"='{sheet_name}'!K92"
        ws_master[f"F{r_table}"] = f"='{sheet_name}'!M92"
        ws_master[f"G{r_table}"] = f"='{sheet_name}'!O92"
        ws_master[f"H{r_table}"] = f"='{sheet_name}'!Q92"
        ws_master[f"I{r_table}"] = f"='{sheet_name}'!T92"

        for c_let in ("B", "C", "D", "E", "F", "G", "H", "I"):
            ws_master[f"{c_let}{r_table}"].font = Font(name=FONT_FAMILY, size=8)
            ws_master[f"{c_let}{r_table}"].alignment = Alignment(horizontal="center", vertical="center")
            ws_master[f"{c_let}{r_table}"].border = BORDER_STANDARD

        ws_master[f"E{r_table}"].number_format = "0.0%"
        ws_master[f"G{r_table}"].number_format = "0.0"

        # Direct Hyperlink Jump in Col J
        jump_cell = ws_master[f"J{r_table}"]
        jump_cell.value = "Jump ↗"
        jump_cell.hyperlink = f"#'{sheet_name}'!A1"
        jump_cell.font = Font(name=FONT_FAMILY, size=7.5, bold=True, color="0284C7", underline="single")
        jump_cell.alignment = Alignment(horizontal="center", vertical="center")
        jump_cell.border = BORDER_STANDARD

    # In-Cell Data Bars on Success % in Master Dashboard
    dbar_master = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="10B981")
    ws_master.conditional_formatting.add("E19:E70", dbar_master)

    # Master Charts: Placed on Right Side (Columns L to Z)
    # Chart 1: Multi-Year Habit Performance Trend Line Chart (at L18)
    chart_m_habits = LineChart()
    chart_m_habits.title = "Multi-Year Habit Execution Trend (2026 – 2030)"
    chart_m_habits.style = 13
    chart_m_habits.y_axis.title = "Success Rate (%)"
    chart_m_habits.x_axis.title = "Month"
    chart_m_habits.y_axis.scaling.min = 0
    chart_m_habits.y_axis.scaling.max = 1
    chart_m_habits.width = 16.0
    chart_m_habits.height = 10.5
    chart_m_habits.legend = None

    data_m_habits = Reference(ws_master, min_col=5, min_row=18, max_row=70)
    cats_m_habits = Reference(ws_master, min_col=1, min_row=19, max_row=70)
    chart_m_habits.add_data(data_m_habits, titles_from_data=True)
    chart_m_habits.set_categories(cats_m_habits)

    if chart_m_habits.series:
        s_m = chart_m_habits.series[0]
        s_m.graphicalProperties.line.solidFill = "10B981" # Emerald Green
        s_m.graphicalProperties.line.width = 25000
        s_m.marker.symbol = "circle"
        s_m.marker.size = 4
        s_m.marker.graphicalProperties.solidFill = "047857"

    ws_master.add_chart(chart_m_habits, "L18")

    # Chart 2: Sleep vs Consistency Multi-Year Trend Chart (at L42)
    chart_m_sleep = LineChart()
    chart_m_sleep.title = "Multi-Year Nightly Sleep Average (Hours vs Month)"
    chart_m_sleep.style = 11
    chart_m_sleep.y_axis.title = "Sleep Hours"
    chart_m_sleep.x_axis.title = "Month"
    chart_m_sleep.y_axis.scaling.min = 0
    chart_m_sleep.y_axis.scaling.max = 12
    chart_m_sleep.width = 16.0
    chart_m_sleep.height = 10.5
    chart_m_sleep.legend = None

    data_m_sleep = Reference(ws_master, min_col=7, min_row=18, max_row=70)
    cats_m_sleep = Reference(ws_master, min_col=1, min_row=19, max_row=70)
    chart_m_sleep.add_data(data_m_sleep, titles_from_data=True)
    chart_m_sleep.set_categories(cats_m_sleep)

    if chart_m_sleep.series:
        s_s = chart_m_sleep.series[0]
        s_s.graphicalProperties.line.solidFill = "6366F1" # Indigo
        s_s.graphicalProperties.line.width = 25000
        s_s.marker.symbol = "diamond"
        s_s.marker.size = 4
        s_s.marker.graphicalProperties.solidFill = "4338CA"

    ws_master.add_chart(chart_m_sleep, "L42")

    # =========================================================================
    # PART 2: 52 MONTHLY SHEETS WITH NAVIGATION BAR & STATUS BADGES
    # =========================================================================
    for month_idx, (year, month) in enumerate(target_months):
        month_name = calendar.month_name[month]
        sheet_title = f"{month_name[:3]} {year}"
        ws = wb.create_sheet(title=sheet_title)
        
        # Gridlines and Page setup
        ws.views.sheetView[0].showGridLines = True
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

        days_in_month = calendar.monthrange(year, month)[1]
        first_weekday = calendar.monthrange(year, month)[0] # 0 = Monday, 6 = Sunday
        month_weekdays = [weekdays_abbr[(first_weekday + d) % 7] for d in range(days_in_month)]

        # Column Layout:
        # Col A: S.No. (1)
        # Col B: Protocols (2)
        # Col C: Category (3)
        # Col D: Target (4)
        # Col E to (4 + days_in_month): Days 1 to 30/31
        first_day_col = 5
        last_day_col = 4 + days_in_month
        success_col = last_day_col + 1
        done_tgt_col = last_day_col + 2
        status_col = last_day_col + 3      # NEW: Automated Status Badges Column
        helper_done_col = last_day_col + 4 # Hidden helper column for exact tick counts

        first_day_let = get_column_letter(first_day_col)
        last_day_letter = get_column_letter(last_day_col)
        success_col_letter = get_column_letter(success_col)
        done_tgt_col_letter = get_column_letter(done_tgt_col)
        status_col_letter = get_column_letter(status_col)
        helper_done_col_let = get_column_letter(helper_done_col)

        # Set Column Widths:
        ws.column_dimensions['A'].width = 6.0
        ws.column_dimensions['B'].width = 25.0
        ws.column_dimensions['C'].width = 16.5
        ws.column_dimensions['D'].width = 8.0
        for col_idx in range(first_day_col, last_day_col + 1):
            c_let = get_column_letter(col_idx)
            ws.column_dimensions[c_let].width = 4.8
        ws.column_dimensions[success_col_letter].width = 11.5
        ws.column_dimensions[done_tgt_col_letter].width = 12.0
        ws.column_dimensions[status_col_letter].width = 13.5
        ws.column_dimensions[helper_done_col_let].width = 8.0
        ws.column_dimensions[helper_done_col_let].hidden = True

        # =========================================================================
        # ROW 1: SLEEK TOP QUICK-JUMP NAVIGATION BAR
        # =========================================================================
        ws.row_dimensions[1].height = 18

        # Master Dashboard Jump
        ws.merge_cells("A1:B1")
        nav_home = ws["A1"]
        nav_home.value = "🏠 Master Dashboard"
        nav_home.hyperlink = "#'Master Dashboard'!A1"
        nav_home.font = Font(name=FONT_FAMILY, size=8, bold=True, color="FFFFFF")
        nav_home.fill = NAV_ACCENT_FILL
        nav_home.alignment = Alignment(horizontal="center", vertical="center")
        ws["B1"].fill = NAV_ACCENT_FILL
        ws["A1"].border = BORDER_NAV_PILL
        ws["B1"].border = BORDER_NAV_PILL

        # Year Quick Jumps
        year_jumps = [
            ("C1", "2026", "#'Sep 2026'!A1"),
            ("D1", "2027", "#'Jan 2027'!A1"),
            ("E1", "2028", "#'Jan 2028'!A1"),
            ("F1", "2029", "#'Jan 2029'!A1"),
            ("G1", "2030", "#'Jan 2030'!A1"),
        ]
        for c_pos, y_lbl, y_link in year_jumps:
            y_cell = ws[c_pos]
            y_cell.value = y_lbl
            y_cell.hyperlink = y_link
            is_cur_year = (str(year) == y_lbl)
            y_cell.font = Font(name=FONT_FAMILY, size=7.5, bold=True, color="F8FAFC" if is_cur_year else "CBD5E1")
            y_cell.fill = NAV_BAR_FILL if is_cur_year else NAV_PILL_FILL
            y_cell.alignment = Alignment(horizontal="center", vertical="center")
            y_cell.border = BORDER_NAV_PILL

        # Relative Month Jumps: Prev and Next
        prev_cell = ws["H1"]
        if month_idx > 0:
            prev_name = month_titles[month_idx - 1]
            prev_cell.value = "◀ Prev"
            prev_cell.hyperlink = f"#'{prev_name}'!A1"
            prev_cell.font = Font(name=FONT_FAMILY, size=7.5, bold=True, color="FFFFFF")
        else:
            prev_cell.value = "◀ Prev"
            prev_cell.font = Font(name=FONT_FAMILY, size=7.5, color="94A3B8")
        prev_cell.fill = NAV_PILL_FILL
        prev_cell.alignment = Alignment(horizontal="center", vertical="center")
        prev_cell.border = BORDER_NAV_PILL

        next_cell = ws["I1"]
        if month_idx < len(target_months) - 1:
            next_name = month_titles[month_idx + 1]
            next_cell.value = "Next ▶"
            next_cell.hyperlink = f"#'{next_name}'!A1"
            next_cell.font = Font(name=FONT_FAMILY, size=7.5, bold=True, color="FFFFFF")
        else:
            next_cell.value = "Next ▶"
            next_cell.font = Font(name=FONT_FAMILY, size=7.5, color="94A3B8")
        next_cell.fill = NAV_PILL_FILL
        next_cell.alignment = Alignment(horizontal="center", vertical="center")
        next_cell.border = BORDER_NAV_PILL

        # --- Row 2: Title, Quote, and @surajvansh ---
        ws.row_dimensions[2].height = 24
        ws["A2"] = "Habit Tracker"
        ws["A2"].font = Font(name=FONT_FAMILY, size=18, bold=True, color="000000")
        ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

        quote_text = QUOTES[month_idx % len(QUOTES)]
        quote_start_col = 8
        quote_end_col = last_day_col - 5
        ws.merge_cells(start_row=2, start_column=quote_start_col, end_row=2, end_column=quote_end_col)
        quote_cell = ws.cell(row=2, column=quote_start_col)
        quote_cell.value = f'"{quote_text}"'
        quote_cell.font = Font(name=FONT_FAMILY, size=11, bold=True, color="000000")
        quote_cell.alignment = Alignment(horizontal="center", vertical="center")

        handle_start_col = last_day_col - 4
        ws.merge_cells(start_row=2, start_column=handle_start_col, end_row=2, end_column=status_col)
        handle_cell = ws.cell(row=2, column=handle_start_col)
        handle_cell.value = "@surajvansh"
        handle_cell.font = Font(name=FONT_FAMILY, size=13, bold=True, color="000000")
        handle_cell.alignment = Alignment(horizontal="right", vertical="center")

        # --- Row 3: Month Year and Days Dashboard ---
        ws.row_dimensions[3].height = 18
        ws["A3"] = f"{month_name} {year}"
        ws["A3"].font = Font(name=FONT_FAMILY, size=12, bold=True, color="000000")
        ws["A3"].alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells(start_row=3, start_column=quote_start_col, end_row=3, end_column=quote_end_col)
        dash_cell = ws.cell(row=3, column=quote_start_col)
        dash_cell.value = f"{days_in_month} Days Dashboard"
        dash_cell.font = Font(name=FONT_FAMILY, size=10, bold=True, color="000000")
        dash_cell.alignment = Alignment(horizontal="center", vertical="center")

        # Row 4: Spacing
        ws.row_dimensions[4].height = 5

        # =========================================================================
        # TOP SECTION (BEFORE TABLE):
        # 1. Monthly Targets (8 slots)
        # 2. Notes Yellow Container
        # =========================================================================
        ws.row_dimensions[5].height = 16
        ws["A5"] = "MONTHLY TARGETS"
        ws["A5"].font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        ws["A5"].alignment = Alignment(horizontal="left", vertical="center")

        col2_target_start = 16
        for i in range(1, 5):
            t_row = 5 + i
            ws.row_dimensions[t_row].height = 13

            # Column 1 (Targets 1-4)
            ws.merge_cells(start_row=t_row, start_column=1, end_row=t_row, end_column=col2_target_start - 2)
            t1_cell = ws.cell(row=t_row, column=1)
            t1_cell.value = f"{i}. ____________________ : ____________________"
            t1_cell.font = Font(name=FONT_FAMILY, size=8)
            t1_cell.alignment = Alignment(horizontal="left", vertical="center")

            # Column 2 (Targets 5-8)
            i2 = i + 4
            ws.merge_cells(start_row=t_row, start_column=col2_target_start, end_row=t_row, end_column=status_col)
            t2_cell = ws.cell(row=t_row, column=col2_target_start)
            t2_cell.value = f"{i2}. ____________________ : ____________________"
            t2_cell.font = Font(name=FONT_FAMILY, size=8)
            t2_cell.alignment = Alignment(horizontal="left", vertical="center")

        # Row 10: Spacing
        ws.row_dimensions[10].height = 5

        # --- Row 11: Notes Section Label ---
        ws.row_dimensions[11].height = 15
        ws["A11"] = "NOTES:"
        ws["A11"].font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        ws["A11"].alignment = Alignment(horizontal="left", vertical="center")

        # Rows 12 to 14: Notes Yellow Box
        for r_note in range(12, 15):
            ws.row_dimensions[r_note].height = 13
        ws.merge_cells(f"A12:{status_col_letter}14")
        note_box = ws["A12"]
        note_box.fill = NOTES_FILL
        note_box.font = Font(name=FONT_FAMILY, size=8.5)
        note_box.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        for r in range(12, 15):
            for c in range(1, status_col + 1):
                ws.cell(row=r, column=c).fill = NOTES_FILL
                ws.cell(row=r, column=c).border = BORDER_NOTE

        # Row 15: Spacing before main table
        ws.row_dimensions[15].height = 8

        # =========================================================================
        # MAIN PROTOCOLS TABLE (Rows 16 & 17: Headers)
        # =========================================================================
        ws.row_dimensions[16].height = 15
        ws.row_dimensions[17].height = 15

        # Merge S.No.
        ws.merge_cells("A16:A17")
        ws["A16"] = "S.No."
        ws["A16"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws["A16"].alignment = Alignment(horizontal="center", vertical="center")
        for r in (16, 17):
            ws[f"A{r}"].border = BORDER_STANDARD

        # Merge Protocols
        ws.merge_cells("B16:B17")
        ws["B16"] = "Protocols"
        ws["B16"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws["B16"].alignment = Alignment(horizontal="center", vertical="center")
        for r in (16, 17):
            ws[f"B{r}"].border = BORDER_STANDARD

        # Merge Category
        ws.merge_cells("C16:C17")
        ws["C16"] = "Category"
        ws["C16"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws["C16"].alignment = Alignment(horizontal="center", vertical="center")
        for r in (16, 17):
            ws[f"C{r}"].border = BORDER_STANDARD

        # Merge Target
        ws.merge_cells("D16:D17")
        ws["D16"] = "Target"
        ws["D16"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws["D16"].alignment = Alignment(horizontal="center", vertical="center")
        for r in (16, 17):
            ws[f"D{r}"].border = BORDER_STANDARD

        # Day Number and Weekday Headers
        for d in range(1, days_in_month + 1):
            col = first_day_col + (d - 1)
            c_let = get_column_letter(col)
            is_sunday = (month_weekdays[d - 1] == 'Su')
            header_fill = SUNDAY_HEADER_FILL if is_sunday else WHITE_FILL

            # Row 16: Day number
            cell_d = ws[f"{c_let}16"]
            cell_d.value = d
            cell_d.font = Font(name=FONT_FAMILY, size=8, bold=True)
            cell_d.alignment = Alignment(horizontal="center", vertical="center")
            cell_d.fill = header_fill
            cell_d.border = BORDER_SUNDAY if is_sunday else BORDER_STANDARD

            # Row 17: Weekday abbreviation
            cell_w = ws[f"{c_let}17"]
            cell_w.value = month_weekdays[d - 1]
            cell_w.font = Font(name=FONT_FAMILY, size=7, bold=is_sunday)
            cell_w.alignment = Alignment(horizontal="center", vertical="center")
            cell_w.fill = header_fill
            cell_w.border = BORDER_SUNDAY if is_sunday else BORDER_STANDARD

        # Merge Success %
        ws.merge_cells(f"{success_col_letter}16:{success_col_letter}17")
        ws[f"{success_col_letter}16"] = "Success %"
        ws[f"{success_col_letter}16"].font = Font(name=FONT_FAMILY, size=7.5, bold=True)
        ws[f"{success_col_letter}16"].alignment = Alignment(horizontal="center", vertical="center")
        for r in (16, 17):
            ws[f"{success_col_letter}{r}"].border = BORDER_STANDARD

        # Merge Done / Target
        ws.merge_cells(f"{done_tgt_col_letter}16:{done_tgt_col_letter}17")
        ws[f"{done_tgt_col_letter}16"] = "Done / Target"
        ws[f"{done_tgt_col_letter}16"].font = Font(name=FONT_FAMILY, size=7.5, bold=True)
        ws[f"{done_tgt_col_letter}16"].alignment = Alignment(horizontal="center", vertical="center")
        for r in (16, 17):
            ws[f"{done_tgt_col_letter}{r}"].border = BORDER_STANDARD

        # Merge Status (Badge Column)
        ws.merge_cells(f"{status_col_letter}16:{status_col_letter}17")
        ws[f"{status_col_letter}16"] = "Status"
        ws[f"{status_col_letter}16"].font = Font(name=FONT_FAMILY, size=7.5, bold=True)
        ws[f"{status_col_letter}16"].alignment = Alignment(horizontal="center", vertical="center")
        for r in (16, 17):
            ws[f"{status_col_letter}{r}"].border = BORDER_STANDARD

        # Helper Done Count Header
        ws[f"{helper_done_col_let}16"] = "Done Count"

        # --- Rows 18 to 37: 20 Protocol Rows ---
        first_habit_row = 18
        total_habits = 20
        last_habit_row = first_habit_row + total_habits - 1 # 37

        # DataValidation for Category Dropdown
        dv_category = DataValidation(
            type="list",
            formula1=f'"{categories_dropdown_str}"',
            allow_blank=True,
            promptTitle="Life Domain / Category",
            prompt="Select habit category"
        )
        ws.add_data_validation(dv_category)
        dv_category.add(f"C{first_habit_row}:C{last_habit_row}")

        # DataValidation for Target dropdown (1 to 31)
        dv_target = DataValidation(
            type="list",
            formula1=f'"{target_dropdown_str}"',
            allow_blank=True,
            promptTitle="Target Days",
            prompt="Select monthly target (1-31)"
        )
        ws.add_data_validation(dv_target)
        dv_target.add(f"D{first_habit_row}:D{last_habit_row}")

        # DataValidation for Protocols Day Cells: Tick / Cross ("✓,✗")
        dv_tick_cross = DataValidation(
            type="list",
            formula1='"✓,✗"',
            allow_blank=True,
            promptTitle="Status",
            prompt="Select ✓ (Done) or ✗ (Missed)"
        )
        ws.add_data_validation(dv_tick_cross)
        dv_tick_cross.add(f"{first_day_let}{first_habit_row}:{last_day_letter}{last_habit_row}")

        for i in range(1, total_habits + 1):
            row = first_habit_row + (i - 1)
            ws.row_dimensions[row].height = 14.5
            group_idx = (i - 1) // 4
            group_fill_hex = GROUP_COLORS_HEX[group_idx]
            sunday_fill_hex = SUNDAY_GROUP_COLORS_HEX[group_idx]

            row_fill = PatternFill(start_color=group_fill_hex, end_color=group_fill_hex, fill_type="solid")
            row_sunday_fill = PatternFill(start_color=sunday_fill_hex, end_color=sunday_fill_hex, fill_type="solid")

            # Col A: S.No.
            cell_sn = ws[f"A{row}"]
            cell_sn.value = i
            cell_sn.font = Font(name=FONT_FAMILY, size=8)
            cell_sn.alignment = Alignment(horizontal="center", vertical="center")
            cell_sn.fill = row_fill
            cell_sn.border = BORDER_STANDARD

            # Col B: Protocols (blank for user input)
            cell_proto = ws[f"B{row}"]
            cell_proto.font = Font(name=FONT_FAMILY, size=8.5)
            cell_proto.alignment = Alignment(horizontal="left", vertical="center")
            cell_proto.fill = row_fill
            cell_proto.border = BORDER_STANDARD

            # Col C: Category (Dropdown)
            cell_cat = ws[f"C{row}"]
            cell_cat.font = Font(name=FONT_FAMILY, size=8)
            cell_cat.alignment = Alignment(horizontal="center", vertical="center")
            cell_cat.fill = row_fill
            cell_cat.border = BORDER_STANDARD

            # Col D: Target (Dropdown 1-31)
            cell_tgt = ws[f"D{row}"]
            cell_tgt.font = Font(name=FONT_FAMILY, size=8)
            cell_tgt.alignment = Alignment(horizontal="center", vertical="center")
            cell_tgt.fill = row_fill
            cell_tgt.border = BORDER_STANDARD

            # Cols E..: Day cells with dropdown ✓ / ✗
            for d in range(1, days_in_month + 1):
                col = first_day_col + (d - 1)
                c_let = get_column_letter(col)
                is_sunday = (month_weekdays[d - 1] == 'Su')
                cell_day = ws[f"{c_let}{row}"]
                cell_day.font = Font(name=FONT_FAMILY, size=9, bold=True)
                cell_day.alignment = Alignment(horizontal="center", vertical="center")
                cell_day.fill = row_sunday_fill if is_sunday else row_fill
                cell_day.border = BORDER_SUNDAY if is_sunday else BORDER_STANDARD

            # Col Helper: Total Done Count
            cell_helper = ws[f"{helper_done_col_let}{row}"]
            cell_helper.value = f'=COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "✓")'

            # Expressions for Row calculations
            row_ticks = f'(COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "✓") + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "✔") + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "v") + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "V"))'
            row_eval = f'({row_ticks} + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "✗") + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "x") + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "X"))'

            # Col Success %:
            cell_succ = ws[f"{success_col_letter}{row}"]
            cell_succ.value = f'=IF(ISBLANK(B{row}), "", IF(ISNUMBER(D{row}), IF(D{row}>0, {row_ticks}/D{row}, 0), IF({row_eval}>0, {row_ticks}/{row_eval}, "")))'
            cell_succ.font = Font(name=FONT_FAMILY, size=8)
            cell_succ.alignment = Alignment(horizontal="center", vertical="center")
            cell_succ.fill = row_fill
            cell_succ.border = BORDER_STANDARD
            cell_succ.number_format = "0.0%"

            # Col Done / Target: Shows exact number of days done vs target at denominator
            cell_done_tgt = ws[f"{done_tgt_col_letter}{row}"]
            cell_done_tgt.value = f'=IF(ISBLANK(B{row}), "", IF(ISNUMBER(D{row}), IF(D{row}>0, {row_ticks} & "/" & D{row}, {row_ticks} & "/0"), IF({row_eval}>0, {row_ticks} & "/" & {row_eval}, "")))'
            cell_done_tgt.font = Font(name=FONT_FAMILY, size=8, bold=True)
            cell_done_tgt.alignment = Alignment(horizontal="center", vertical="center")
            cell_done_tgt.fill = row_fill
            cell_done_tgt.border = BORDER_STANDARD

            # Col Status: Automated Status Badges
            cell_status = ws[f"{status_col_letter}{row}"]
            cell_status.value = f'=IF(ISBLANK(B{row}), "", IF({cell_succ.coordinate}>=0.85, "🔥 Mastered", IF({cell_succ.coordinate}>=0.7, "⚡ Consistent", "⚠️ Needs Focus")))'
            cell_status.font = Font(name=FONT_FAMILY, size=8, bold=True)
            cell_status.alignment = Alignment(horizontal="center", vertical="center")
            cell_status.fill = row_fill
            cell_status.border = BORDER_STANDARD

        # Conditional Formatting: In-Cell DataBars on Success % Column
        dbar_rule = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="10B981")
        ws.conditional_formatting.add(f"{success_col_letter}{first_habit_row}:{success_col_letter}{last_habit_row}", dbar_rule)

        # Dynamic Tick / Cross conditional formatting
        rule_tick = CellIsRule(operator='equal', formula=['"✓"'], fill=CF_TICK_FILL, font=CF_TICK_FONT)
        rule_tick_alt = CellIsRule(operator='equal', formula=['"✔"'], fill=CF_TICK_FILL, font=CF_TICK_FONT)
        rule_cross = CellIsRule(operator='equal', formula=['"✗"'], fill=CF_CROSS_FILL, font=CF_CROSS_FONT)
        rule_cross_alt = CellIsRule(operator='equal', formula=['"x"'], fill=CF_CROSS_FILL, font=CF_CROSS_FONT)
        rule_cross_upper = CellIsRule(operator='equal', formula=['"X"'], fill=CF_CROSS_FILL, font=CF_CROSS_FONT)

        day_cf_range = f"{first_day_let}{first_habit_row}:{last_day_letter}{last_habit_row}"
        ws.conditional_formatting.add(day_cf_range, rule_tick)
        ws.conditional_formatting.add(day_cf_range, rule_tick_alt)
        ws.conditional_formatting.add(day_cf_range, rule_cross)
        ws.conditional_formatting.add(day_cf_range, rule_cross_alt)
        ws.conditional_formatting.add(day_cf_range, rule_cross_upper)

        # Dynamic "TODAY" Column Highlighting
        for d in range(1, days_in_month + 1):
            col = first_day_col + (d - 1)
            c_let = get_column_letter(col)
            today_formula = f'AND(DAY(TODAY())={d}, MONTH(TODAY())={month}, YEAR(TODAY())={year})'
            rule_today_hdr = FormulaRule(formula=[today_formula], fill=CF_TODAY_FILL, font=CF_TODAY_FONT)
            ws.conditional_formatting.add(f"{c_let}16:{c_let}17", rule_today_hdr)

        # --- Row 38: DAILY TOTAL SCORE (Displays done/target like 8/10) ---
        score_row = 38
        ws.row_dimensions[score_row].height = 18
        
        ws.merge_cells(f"A{score_row}:D{score_row}")
        score_label = ws[f"A{score_row}"]
        score_label.value = "DAILY TOTAL SCORE"
        score_label.font = Font(name=FONT_FAMILY, size=8, bold=True)
        score_label.alignment = Alignment(horizontal="right", vertical="center")
        for c in ("A", "B", "C", "D"):
            ws[f"{c}{score_row}"].border = BORDER_STANDARD

        for d in range(1, days_in_month + 1):
            col = first_day_col + (d - 1)
            c_let = get_column_letter(col)
            cell_score = ws[f"{c_let}{score_row}"]
            
            day_ticks = f'(COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "✓") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "✔") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "v") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "V"))'
            day_eval = f'({day_ticks} + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "✗") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "x") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "X"))'
            
            cell_score.value = f'=IF({day_eval} > 0, {day_ticks} & "/" & {day_eval}, "")'
            cell_score.font = Font(name=FONT_FAMILY, size=8, bold=True)
            cell_score.alignment = Alignment(horizontal="center", vertical="center")
            cell_score.border = BORDER_STANDARD

        # Row 38 Success Column (Overall Month Success Rate)
        all_month_ticks = f'SUM({helper_done_col_let}{first_habit_row}:{helper_done_col_let}{last_habit_row})'
        all_month_targets = f'SUM(D{first_habit_row}:D{last_habit_row})'

        score_summary = ws[f"{success_col_letter}{score_row}"]
        score_summary.value = f'=IF({all_month_targets} > 0, {all_month_ticks} / {all_month_targets}, "")'
        score_summary.font = Font(name=FONT_FAMILY, size=8, bold=True)
        score_summary.alignment = Alignment(horizontal="center", vertical="center")
        score_summary.border = BORDER_STANDARD
        score_summary.number_format = "0.0%"

        # Row 38 Done / Target Column: Total Done vs Total Target in Month
        done_tgt_summary = ws[f"{done_tgt_col_letter}{score_row}"]
        done_tgt_summary.value = f'=IF({all_month_targets} > 0, {all_month_ticks} & "/" & {all_month_targets}, "")'
        done_tgt_summary.font = Font(name=FONT_FAMILY, size=8, bold=True)
        done_tgt_summary.alignment = Alignment(horizontal="center", vertical="center")
        done_tgt_summary.border = BORDER_STANDARD

        # Status Col Row 38
        ws[f"{status_col_letter}{score_row}"].border = BORDER_STANDARD

        # --- Row 39: DAILY SUCCESS % (Under Daily Total Score) ---
        pct_row = 39
        ws.row_dimensions[pct_row].height = 18

        ws.merge_cells(f"A{pct_row}:D{pct_row}")
        pct_label = ws[f"A{pct_row}"]
        pct_label.value = "DAILY SUCCESS %"
        pct_label.font = Font(name=FONT_FAMILY, size=8, bold=True)
        pct_label.alignment = Alignment(horizontal="right", vertical="center")
        for c in ("A", "B", "C", "D"):
            ws[f"{c}{pct_row}"].border = BORDER_STANDARD

        for d in range(1, days_in_month + 1):
            col = first_day_col + (d - 1)
            c_let = get_column_letter(col)
            cell_pct = ws[f"{c_let}{pct_row}"]
            
            day_ticks = f'(COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "✓") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "✔") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "v") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "V"))'
            day_eval = f'({day_ticks} + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "✗") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "x") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "X"))'
            
            cell_pct.value = f'=IF({day_eval} > 0, {day_ticks} / {day_eval}, "")'
            cell_pct.font = Font(name=FONT_FAMILY, size=7.5, bold=True)
            cell_pct.alignment = Alignment(horizontal="center", vertical="center")
            cell_pct.border = BORDER_STANDARD
            cell_pct.number_format = "0.0%"

        pct_summary = ws[f"{success_col_letter}{pct_row}"]
        pct_summary.value = f'=IF(COUNT({first_day_let}{pct_row}:{last_day_letter}{pct_row}) > 0, AVERAGE({first_day_let}{pct_row}:{last_day_letter}{pct_row}), "")'
        pct_summary.font = Font(name=FONT_FAMILY, size=8, bold=True)
        pct_summary.alignment = Alignment(horizontal="center", vertical="center")
        pct_summary.border = BORDER_STANDARD
        pct_summary.number_format = "0.0%"

        ws[f"{done_tgt_col_letter}{pct_row}"].border = BORDER_STANDARD
        ws[f"{status_col_letter}{pct_row}"].border = BORDER_STANDARD

        # =========================================================================
        # BETWEEN PROTOCOL TABLE AND SLEEP TABLE:
        # Success and Punishment (Stakes) + Two-row gap
        # =========================================================================
        ws.row_dimensions[40].height = 6

        # --- Rows 41 & 42: Punishment & Success (Reward) ---
        ws.row_dimensions[41].height = 15
        ws.row_dimensions[42].height = 15

        ws.merge_cells(f"A41:{status_col_letter}41")
        punish_cell = ws["A41"]
        punish_cell.value = "IF SUCCESS < 80%, PUNISHMENT: ________________________________________________"
        punish_cell.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="C80000")
        punish_cell.alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells(f"A42:{status_col_letter}42")
        reward_cell = ws["A42"]
        reward_cell.value = "IF SUCCESS > 90%, REWARD: __________________________________________________"
        reward_cell.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="007800")
        reward_cell.alignment = Alignment(horizontal="left", vertical="center")

        # --- Rows 43 & 44: Exactly Two-Row Blank Gap ---
        ws.row_dimensions[43].height = 10
        ws.row_dimensions[44].height = 10

        # =========================================================================
        # SLEEP TRACKING SECTION (Starts at Row 45) - Sophisticated Twilight Theme
        # =========================================================================
        ws.row_dimensions[45].height = 15
        ws.row_dimensions[46].height = 15

        ws.merge_cells("A45:D46")
        sleep_hdr = ws["A45"]
        sleep_hdr.value = "Sleep Tracking"
        sleep_hdr.font = Font(name=FONT_FAMILY, size=9, bold=True, color="FFFFFF")
        sleep_hdr.alignment = Alignment(horizontal="center", vertical="center")
        for r in (45, 46):
            for c in ("A", "B", "C", "D"):
                ws[f"{c}{r}"].fill = SLEEP_MAIN_HDR_FILL
                ws[f"{c}{r}"].border = BORDER_STANDARD

        for d in range(1, days_in_month + 1):
            col = first_day_col + (d - 1)
            c_let = get_column_letter(col)
            is_sunday = (month_weekdays[d - 1] == 'Su')
            header_fill = SLEEP_SUN_HDR_FILL if is_sunday else SLEEP_DAY_HDR_FILL

            # Row 45: Day number
            cell_d = ws[f"{c_let}45"]
            cell_d.value = d
            cell_d.font = Font(name=FONT_FAMILY, size=8, bold=True, color="1E1B4B")
            cell_d.alignment = Alignment(horizontal="center", vertical="center")
            cell_d.fill = header_fill
            cell_d.border = BORDER_SUNDAY if is_sunday else BORDER_STANDARD

            # Row 46: Weekday
            cell_w = ws[f"{c_let}46"]
            cell_w.value = month_weekdays[d - 1]
            cell_w.font = Font(name=FONT_FAMILY, size=7, bold=True, color="3730A3" if not is_sunday else "1E1B4B")
            cell_w.alignment = Alignment(horizontal="center", vertical="center")
            cell_w.fill = header_fill
            cell_w.border = BORDER_SUNDAY if is_sunday else BORDER_STANDARD

        # --- Row 47: Sleep Hours with Dropdown (1 to 24) for each day ---
        sleep_row = 47
        ws.row_dimensions[sleep_row].height = 20

        ws.merge_cells(f"A{sleep_row}:D{sleep_row}")
        sleep_lbl = ws[f"A{sleep_row}"]
        sleep_lbl.value = "Sleep Hours (1-24)"
        sleep_lbl.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="FFFFFF")
        sleep_lbl.alignment = Alignment(horizontal="center", vertical="center")
        for c in ("A", "B", "C", "D"):
            ws[f"{c}{sleep_row}"].fill = SLEEP_LBL_FILL
            ws[f"{c}{sleep_row}"].border = BORDER_STANDARD

        # DataValidation for Sleep Hours (1 to 24)
        dv_sleep = DataValidation(
            type="list",
            formula1=f'"{sleep_dropdown_str}"',
            allow_blank=True,
            promptTitle="Sleep Hours",
            prompt="Select sleep hours (1-24)"
        )
        ws.add_data_validation(dv_sleep)
        dv_sleep.add(f"{first_day_let}{sleep_row}:{last_day_letter}{sleep_row}")

        for d in range(1, days_in_month + 1):
            col = first_day_col + (d - 1)
            c_let = get_column_letter(col)
            is_sunday = (month_weekdays[d - 1] == 'Su')
            cell_s = ws[f"{c_let}{sleep_row}"]
            cell_s.font = Font(name=FONT_FAMILY, size=9, bold=True, color="1E1B4B")
            cell_s.alignment = Alignment(horizontal="center", vertical="center")
            cell_s.number_format = "0"
            if is_sunday:
                cell_s.fill = SLEEP_SUN_CELL_FILL
                cell_s.border = BORDER_SUNDAY
            else:
                cell_s.fill = SLEEP_CELL_FILL
                cell_s.border = BORDER_STANDARD

        # --- Sleep Tracking Line Chart (Row 49 to Row 64) ---
        ws.row_dimensions[48].height = 6

        chart_sleep = LineChart()
        chart_sleep.title = "Daily Sleep Trend (Hours vs Day)"
        chart_sleep.style = 13
        chart_sleep.y_axis.title = "Hours"
        chart_sleep.x_axis.title = "Day"
        chart_sleep.y_axis.scaling.min = 0
        chart_sleep.y_axis.scaling.max = 24
        chart_sleep.width = 24
        chart_sleep.height = 7.5
        chart_sleep.legend = None

        data_sleep = Reference(ws, min_col=first_day_col, min_row=sleep_row, max_col=last_day_col, max_row=sleep_row)
        cats_sleep = Reference(ws, min_col=first_day_col, min_row=45, max_col=last_day_col, max_row=45)

        chart_sleep.add_data(data_sleep, titles_from_data=False, from_rows=True)
        chart_sleep.set_categories(cats_sleep)

        if chart_sleep.series:
            series = chart_sleep.series[0]
            series.graphicalProperties.line.solidFill = "3B82F6"
            series.graphicalProperties.line.width = 25000
            series.marker.symbol = "circle"
            series.marker.size = 5
            series.marker.graphicalProperties.solidFill = "1D4ED8"

        ws.add_chart(chart_sleep, "A49")

        # Padding before End-of-Month Analytics
        for r_pad in range(49, 66):
            if ws.row_dimensions[r_pad].height is None:
                ws.row_dimensions[r_pad].height = 14

        # =========================================================================
        # END-OF-MONTH ADVANCED ANALYTICS & REVIEWS (Row 66 Onwards)
        # =========================================================================
        ws.row_dimensions[66].height = 8

        # --- SECTION 1: EXECUTIVE MONTHLY KPI CARDS (Rows 67–71) ---
        # Card 1: Monthly Success Rate (Cols A-D)
        ws.merge_cells("A67:D67")
        ws["A67"] = "MONTHLY SUCCESS RATE"
        ws["A67"].font = Font(name=FONT_FAMILY, size=8, bold=True, color="15803D")
        ws["A67"].alignment = Alignment(horizontal="center", vertical="center")
        
        ws.merge_cells("A68:D70")
        ws["A68"] = f'=IF({all_month_targets}>0, {all_month_ticks}/{all_month_targets}, 0)'
        ws["A68"].font = Font(name=FONT_FAMILY, size=18, bold=True, color="166534")
        ws["A68"].alignment = Alignment(horizontal="center", vertical="center")
        ws["A68"].number_format = "0.0%"

        ws.merge_cells("A71:D71")
        ws["A71"] = "Target Completion"
        ws["A71"].font = Font(name=FONT_FAMILY, size=7.5, italic=True, color="15803D")
        ws["A71"].alignment = Alignment(horizontal="center", vertical="center")

        for r_c in range(67, 72):
            for c_c in range(1, 5):
                c_cell = ws.cell(row=r_c, column=c_c)
                c_cell.fill = KPI_GREEN_FILL
                c_cell.border = BORDER_CARD_GREEN

        # Card 2: Active Habits (Cols F-I)
        ws.merge_cells("F67:I67")
        ws["F67"] = "ACTIVE PROTOCOLS"
        ws["F67"].font = Font(name=FONT_FAMILY, size=8, bold=True, color="1D4ED8")
        ws["F67"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("F68:I70")
        ws["F68"] = f'=COUNTA(B{first_habit_row}:B{last_habit_row}) & " / 20"'
        ws["F68"].font = Font(name=FONT_FAMILY, size=18, bold=True, color="1E40AF")
        ws["F68"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("F71:I71")
        ws["F71"] = "Habits Tracked"
        ws["F71"].font = Font(name=FONT_FAMILY, size=7.5, italic=True, color="1D4ED8")
        ws["F71"].alignment = Alignment(horizontal="center", vertical="center")

        for r_c in range(67, 72):
            for c_c in range(6, 10):
                c_cell = ws.cell(row=r_c, column=c_c)
                c_cell.fill = KPI_BLUE_FILL
                c_cell.border = BORDER_CARD_BLUE

        # Card 3: Perfect Days (Cols K-N)
        ws.merge_cells("K67:N67")
        ws["K67"] = "PERFECT DAYS (100%)"
        ws["K67"].font = Font(name=FONT_FAMILY, size=8, bold=True, color="B45309")
        ws["K67"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("K68:N70")
        ws["K68"] = f'=COUNTIF({first_day_let}{pct_row}:{last_day_letter}{pct_row}, 1) & " Days"'
        ws["K68"].font = Font(name=FONT_FAMILY, size=18, bold=True, color="92400E")
        ws["K68"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("K71:N71")
        ws["K71"] = "All Marked Habits Completed"
        ws["K71"].font = Font(name=FONT_FAMILY, size=7.5, italic=True, color="B45309")
        ws["K71"].alignment = Alignment(horizontal="center", vertical="center")

        for r_c in range(67, 72):
            for c_c in range(11, 15):
                c_cell = ws.cell(row=r_c, column=c_c)
                c_cell.fill = KPI_AMBER_FILL
                c_cell.border = BORDER_CARD_AMBER

        # Card 4: Avg Sleep Duration (Cols P-S)
        ws.merge_cells("P67:S67")
        ws["P67"] = "AVG SLEEP DURATION"
        ws["P67"].font = Font(name=FONT_FAMILY, size=8, bold=True, color="6D28D9")
        ws["P67"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("P68:S70")
        ws["P68"] = f'=IFERROR(TEXT(AVERAGE({first_day_let}{sleep_row}:{last_day_letter}{sleep_row}), "0.0") & " hrs", "-")'
        ws["P68"].font = Font(name=FONT_FAMILY, size=18, bold=True, color="5B21B6")
        ws["P68"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("P71:S71")
        ws["P71"] = "Nightly Rest Average"
        ws["P71"].font = Font(name=FONT_FAMILY, size=7.5, italic=True, color="6D28D9")
        ws["P71"].alignment = Alignment(horizontal="center", vertical="center")

        for r_c in range(67, 72):
            for c_c in range(16, 20):
                c_cell = ws.cell(row=r_c, column=c_c)
                c_cell.fill = KPI_PURPLE_FILL
                c_cell.border = BORDER_CARD_PURPLE

        # Row 72: Spacing
        ws.row_dimensions[72].height = 8

        # --- SECTION 2: WEEKLY BREAKDOWN (SLUMP DETECTOR) & CATEGORY BREAKDOWN (Rows 74–82) ---
        ws.row_dimensions[73].height = 6

        # Left Table: Weekly Performance (Cols A-E, Rows 74-82)
        ws.merge_cells("A74:E74")
        ws["A74"] = "WEEKLY PERFORMANCE (SLUMP DETECTOR)"
        ws["A74"].font = Font(name=FONT_FAMILY, size=9, bold=True, color="FFFFFF")
        ws["A74"].fill = HDR_SLATE
        ws["A74"].alignment = Alignment(horizontal="center", vertical="center")
        for c in range(1, 6):
            ws.cell(row=74, column=c).fill = HDR_SLATE
            ws.cell(row=74, column=c).border = BORDER_STANDARD

        w_headers = ["Week", "Days Range", "Checks Done", "Evaluated", "Success %"]
        ws.row_dimensions[75].height = 15
        for idx, h_text in enumerate(w_headers):
            cell = ws.cell(row=75, column=idx + 1)
            cell.value = h_text
            cell.font = Font(name=FONT_FAMILY, size=8, bold=True)
            cell.fill = SUBHDR_GRAY
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER_STANDARD

        # Week ranges calculation
        week_ranges = [
            ("Week 1", "Days 1 - 7", first_day_col, first_day_col + 6),
            ("Week 2", "Days 8 - 14", first_day_col + 7, first_day_col + 13),
            ("Week 3", "Days 15 - 21", first_day_col + 14, first_day_col + 20),
            ("Week 4", "Days 22 - 28", first_day_col + 21, first_day_col + 27),
            ("Week 5", f"Days 29 - {days_in_month}", first_day_col + 28, last_day_col if days_in_month > 28 else first_day_col + 28),
        ]

        for w_idx, (w_name, w_range, col_s, col_e) in enumerate(week_ranges):
            r_w = 76 + w_idx
            ws.row_dimensions[r_w].height = 14
            let_s = get_column_letter(col_s)
            let_e = get_column_letter(col_e)

            ws[f"A{r_w}"] = w_name
            ws[f"A{r_w}"].font = Font(name=FONT_FAMILY, size=8, bold=True)
            ws[f"A{r_w}"].alignment = Alignment(horizontal="center", vertical="center")
            ws[f"A{r_w}"].border = BORDER_STANDARD

            ws[f"B{r_w}"] = w_range if (w_idx < 4 or days_in_month > 28) else "N/A"
            ws[f"B{r_w}"].font = Font(name=FONT_FAMILY, size=8)
            ws[f"B{r_w}"].alignment = Alignment(horizontal="center", vertical="center")
            ws[f"B{r_w}"].border = BORDER_STANDARD

            if w_idx < 4 or days_in_month > 28:
                ws[f"C{r_w}"] = f'=COUNTIF({let_s}{first_habit_row}:{let_e}{last_habit_row}, "✓")'
                ws[f"D{r_w}"] = f'=COUNTIF({let_s}{first_habit_row}:{let_e}{last_habit_row}, "✓") + COUNTIF({let_s}{first_habit_row}:{let_e}{last_habit_row}, "✗")'
                ws[f"E{r_w}"] = f'=IFERROR(C{r_w}/D{r_w}, 0)'
            else:
                ws[f"C{r_w}"] = 0
                ws[f"D{r_w}"] = 0
                ws[f"E{r_w}"] = 0

            for c_let, align_val in (("C", "center"), ("D", "center"), ("E", "center")):
                ws[f"{c_let}{r_w}"].font = Font(name=FONT_FAMILY, size=8)
                ws[f"{c_let}{r_w}"].alignment = Alignment(horizontal=align_val, vertical="center")
                ws[f"{c_let}{r_w}"].border = BORDER_STANDARD
            ws[f"E{r_w}"].number_format = "0.0%"

        # Total / Average Row for Weeks
        r_tot_w = 81
        ws.row_dimensions[r_tot_w].height = 14
        ws[f"A{r_tot_w}"] = "Total / Avg"
        ws[f"A{r_tot_w}"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws[f"A{r_tot_w}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"A{r_tot_w}"].border = BORDER_STANDARD

        ws[f"B{r_tot_w}"] = "Month Total"
        ws[f"B{r_tot_w}"].font = Font(name=FONT_FAMILY, size=8, italic=True)
        ws[f"B{r_tot_w}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"B{r_tot_w}"].border = BORDER_STANDARD

        ws[f"C{r_tot_w}"] = "=SUM(C76:C80)"
        ws[f"C{r_tot_w}"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws[f"C{r_tot_w}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"C{r_tot_w}"].border = BORDER_STANDARD

        ws[f"D{r_tot_w}"] = "=SUM(D76:D80)"
        ws[f"D{r_tot_w}"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws[f"D{r_tot_w}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"D{r_tot_w}"].border = BORDER_STANDARD

        ws[f"E{r_tot_w}"] = f'=IFERROR(C{r_tot_w}/D{r_tot_w}, 0)'
        ws[f"E{r_tot_w}"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws[f"E{r_tot_w}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"E{r_tot_w}"].border = BORDER_STANDARD
        ws[f"E{r_tot_w}"].number_format = "0.0%"

        # Weekly Consistency Column Chart (Placed at G74)
        chart_weekly = BarChart()
        chart_weekly.type = "col"
        chart_weekly.title = "Weekly Momentum Trend (%)"
        chart_weekly.style = 10
        chart_weekly.y_axis.title = "Success %"
        chart_weekly.x_axis.title = "Week"
        chart_weekly.width = 13.5
        chart_weekly.height = 7.0
        chart_weekly.legend = None

        data_weekly = Reference(ws, min_col=5, min_row=75, max_row=80)
        cats_weekly = Reference(ws, min_col=1, min_row=76, max_row=80)
        chart_weekly.add_data(data_weekly, titles_from_data=True)
        chart_weekly.set_categories(cats_weekly)
        ws.add_chart(chart_weekly, "G74")

        # Right Table: Category Domain Breakdown (Cols N-S, Rows 74-81)
        ws.merge_cells("N74:S74")
        ws["N74"] = "LIFE DOMAIN / CATEGORY BREAKDOWN"
        ws["N74"].font = Font(name=FONT_FAMILY, size=9, bold=True, color="FFFFFF")
        ws["N74"].fill = HDR_DARK_NAVY
        ws["N74"].alignment = Alignment(horizontal="center", vertical="center")
        for c in range(14, 20):
            ws.cell(row=74, column=c).fill = HDR_DARK_NAVY
            ws.cell(row=74, column=c).border = BORDER_STANDARD

        ws.merge_cells("N75:O75")
        ws["N75"] = "Domain / Category"
        ws["N75"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws["N75"].fill = SUBHDR_GRAY
        ws["N75"].alignment = Alignment(horizontal="center", vertical="center")
        ws["N75"].border = BORDER_STANDARD
        ws["O75"].border = BORDER_STANDARD

        ws["P75"] = "Active Habits"
        ws["Q75"] = "Target Days"
        ws["R75"] = "Done Checks"
        ws["S75"] = "Success %"
        for c_let in ("P", "Q", "R", "S"):
            ws[f"{c_let}75"].font = Font(name=FONT_FAMILY, size=8, bold=True)
            ws[f"{c_let}75"].fill = SUBHDR_GRAY
            ws[f"{c_let}75"].alignment = Alignment(horizontal="center", vertical="center")
            ws[f"{c_let}75"].border = BORDER_STANDARD

        eval_categories = [
            "Health & Fitness",
            "Mindset & Learning",
            "Career & Focus",
            "Personal & Home",
            "Finance & Wealth",
        ]

        for c_idx, cat_name in enumerate(eval_categories):
            r_cat = 76 + c_idx
            ws.row_dimensions[r_cat].height = 14
            ws.merge_cells(f"N{r_cat}:O{r_cat}")
            ws[f"N{r_cat}"] = cat_name
            ws[f"N{r_cat}"].font = Font(name=FONT_FAMILY, size=8, bold=True)
            ws[f"N{r_cat}"].alignment = Alignment(horizontal="left", vertical="center")
            ws[f"N{r_cat}"].border = BORDER_STANDARD
            ws[f"O{r_cat}"].border = BORDER_STANDARD

            ws[f"P{r_cat}"] = f'=COUNTIF($C${first_habit_row}:$C${last_habit_row}, "{cat_name}")'
            ws[f"Q{r_cat}"] = f'=SUMIF($C${first_habit_row}:$C${last_habit_row}, "{cat_name}", $D${first_habit_row}:$D${last_habit_row})'
            ws[f"R{r_cat}"] = f'=SUMIF($C${first_habit_row}:$C${last_habit_row}, "{cat_name}", ${helper_done_col_let}${first_habit_row}:${helper_done_col_let}${last_habit_row})'
            ws[f"S{r_cat}"] = f'=IFERROR(R{r_cat}/Q{r_cat}, 0)'

            for c_let in ("P", "Q", "R", "S"):
                ws[f"{c_let}{r_cat}"].font = Font(name=FONT_FAMILY, size=8)
                ws[f"{c_let}{r_cat}"].alignment = Alignment(horizontal="center", vertical="center")
                ws[f"{c_let}{r_cat}"].border = BORDER_STANDARD
            ws[f"S{r_cat}"].number_format = "0.0%"

        # Category Performance Bar Chart (Placed at U74)
        chart_cat = BarChart()
        chart_cat.type = "col"
        chart_cat.title = "Performance by Domain (%)"
        chart_cat.style = 11
        chart_cat.y_axis.title = "Success %"
        chart_cat.x_axis.title = "Domain"
        chart_cat.width = 13.5
        chart_cat.height = 7.0
        chart_cat.legend = None

        data_cat = Reference(ws, min_col=19, min_row=75, max_row=80)
        cats_cat = Reference(ws, min_col=14, min_row=76, max_row=80)
        chart_cat.add_data(data_cat, titles_from_data=True)
        chart_cat.set_categories(cats_cat)
        ws.add_chart(chart_cat, "U74")

        # Row 83: Spacing
        ws.row_dimensions[82].height = 10
        ws.row_dimensions[83].height = 10

        # --- SECTION 3: SLEEP & PRODUCTIVITY CORRELATION ANALYSIS (Rows 84–88) ---
        ws.merge_cells(f"A84:{status_col_letter}84")
        corr_hdr = ws["A84"]
        corr_hdr.value = "SLEEP & PRODUCTIVITY CORRELATION ANALYSIS"
        corr_hdr.font = Font(name=FONT_FAMILY, size=9, bold=True, color="FFFFFF")
        corr_hdr.fill = HDR_INDIGO
        corr_hdr.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(1, status_col + 1):
            ws.cell(row=84, column=c).fill = HDR_INDIGO
            ws.cell(row=84, column=c).border = BORDER_STANDARD

        # Correlation Table Headers (Row 85)
        ws.row_dimensions[85].height = 15
        corr_cols = [
            ("A85:C85", "Rest Category"),
            ("D85:F85", "Day Criteria"),
            ("G85:I85", "Days Logged"),
            ("J85:M85", "Avg Sleep Duration"),
            ("N85:Q85", "Avg Habit Success %"),
            ("R85:U85", "Productivity Delta"),
        ]
        for c_range, c_title in corr_cols:
            ws.merge_cells(c_range)
            top_cell = ws[c_range.split(":")[0]]
            top_cell.value = c_title
            top_cell.font = Font(name=FONT_FAMILY, size=8, bold=True)
            top_cell.fill = SUBHDR_GRAY
            top_cell.alignment = Alignment(horizontal="center", vertical="center")

        for r_corr in (85, 86, 87, 88):
            for c_corr in range(1, 22):
                ws.cell(row=r_corr, column=c_corr).border = BORDER_STANDARD

        # Row 86: Optimal Rest (>= 7 Hours)
        ws.row_dimensions[86].height = 15
        ws.merge_cells("A86:C86")
        ws["A86"] = "Optimal Rest"
        ws["A86"].font = Font(name=FONT_FAMILY, size=8, bold=True, color="166534")
        ws["A86"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("D86:F86")
        ws["D86"] = ">= 7 Hours / Night"
        ws["D86"].font = Font(name=FONT_FAMILY, size=8)
        ws["D86"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("G86:I86")
        ws["G86"] = f'=COUNTIF({first_day_let}{sleep_row}:{last_day_letter}{sleep_row}, ">=7")'
        ws["G86"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws["G86"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("J86:M86")
        ws["J86"] = f'=IFERROR(TEXT(AVERAGEIF({first_day_let}{sleep_row}:{last_day_letter}{sleep_row}, ">=7"), "0.0") & " hrs", "-")'
        ws["J86"].font = Font(name=FONT_FAMILY, size=8)
        ws["J86"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("N86:Q86")
        ws["N86"] = f'=IFERROR(AVERAGEIFS({first_day_let}{pct_row}:{last_day_letter}{pct_row}, {first_day_let}{sleep_row}:{last_day_letter}{sleep_row}, ">=7", {first_day_let}{pct_row}:{last_day_letter}{pct_row}, ">0"), 0)'
        ws["N86"].font = Font(name=FONT_FAMILY, size=8, bold=True, color="166534")
        ws["N86"].alignment = Alignment(horizontal="center", vertical="center")
        ws["N86"].number_format = "0.0%"

        ws.merge_cells("R86:U86")
        ws["R86"] = "Baseline (Target)"
        ws["R86"].font = Font(name=FONT_FAMILY, size=8, italic=True)
        ws["R86"].alignment = Alignment(horizontal="center", vertical="center")

        # Row 87: Sleep Deficit (< 7 Hours)
        ws.row_dimensions[87].height = 15
        ws.merge_cells("A87:C87")
        ws["A87"] = "Sleep Deficit"
        ws["A87"].font = Font(name=FONT_FAMILY, size=8, bold=True, color="991B1B")
        ws["A87"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("D87:F87")
        ws["D87"] = "< 7 Hours / Night"
        ws["D87"].font = Font(name=FONT_FAMILY, size=8)
        ws["D87"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("G87:I87")
        ws["G87"] = f'=COUNTIF({first_day_let}{sleep_row}:{last_day_letter}{sleep_row}, "<7") - COUNTIF({first_day_let}{sleep_row}:{last_day_letter}{sleep_row}, "")'
        ws["G87"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws["G87"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("J87:M87")
        ws["J87"] = f'=IFERROR(TEXT(AVERAGEIFS({first_day_let}{sleep_row}:{last_day_letter}{sleep_row}, {first_day_let}{sleep_row}:{last_day_letter}{sleep_row}, "<7", {first_day_let}{sleep_row}:{last_day_letter}{sleep_row}, ">0"), "0.0") & " hrs", "-")'
        ws["J87"].font = Font(name=FONT_FAMILY, size=8)
        ws["J87"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("N87:Q87")
        ws["N87"] = f'=IFERROR(AVERAGEIFS({first_day_let}{pct_row}:{last_day_letter}{pct_row}, {first_day_let}{sleep_row}:{last_day_letter}{sleep_row}, "<7", {first_day_let}{pct_row}:{last_day_letter}{pct_row}, ">0"), 0)'
        ws["N87"].font = Font(name=FONT_FAMILY, size=8, bold=True, color="991B1B")
        ws["N87"].alignment = Alignment(horizontal="center", vertical="center")
        ws["N87"].number_format = "0.0%"

        ws.merge_cells("R87:U87")
        ws["R87"] = '=IFERROR(TEXT(N86 - N87, "+0.0%;-0.0%;0.0%") & " Rest Advantage", "-")'
        ws["R87"].font = Font(name=FONT_FAMILY, size=8, bold=True, color="166534")
        ws["R87"].alignment = Alignment(horizontal="center", vertical="center")

        # Row 88: Insight Callout
        ws.row_dimensions[88].height = 18
        ws.merge_cells(f"A88:{status_col_letter}88")
        insight_cell = ws["A88"]
        insight_cell.value = '=IF(AND(ISNUMBER(N86), ISNUMBER(N87), N86>0, N87>0), "💡 Sleep Correlation Insight: You complete " & TEXT(ABS(N86-N87), "0.0%") & IF(N86>=N87, " MORE habits on days with 7+ hours of sleep!", " fewer habits on days with 7+ hours of sleep."), "💡 Log both habits and sleep hours daily to unlock automatic AI correlation insights!")'
        insight_cell.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="1E1B4B")
        insight_cell.fill = PatternFill(start_color="EEF2FF", end_color="EEF2FF", fill_type="solid")
        insight_cell.alignment = Alignment(horizontal="center", vertical="center")

        # Row 89: Spacing
        ws.row_dimensions[89].height = 10

        # --- SECTION 4: STANDARDIZED ANNUAL / QUARTERLY REVIEW DATA BRIDGE (Rows 90–93) ---
        ws.merge_cells(f"A90:{status_col_letter}90")
        bridge_hdr = ws["A90"]
        bridge_hdr.value = "STANDARDIZED DATA BRIDGE (FOR ANNUAL & QUARTERLY REVIEW SHEETS)"
        bridge_hdr.font = Font(name=FONT_FAMILY, size=9, bold=True, color="FFFFFF")
        bridge_hdr.fill = HDR_EMERALD
        bridge_hdr.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(1, status_col + 1):
            ws.cell(row=90, column=c).fill = HDR_EMERALD
            ws.cell(row=90, column=c).border = BORDER_STANDARD

        ws.row_dimensions[91].height = 15
        bridge_headers = [
            ("A91:B91", "Month"),
            ("C91:D91", "Year"),
            ("E91:F91", "Active Habits"),
            ("G91:H91", "Target Days"),
            ("I91:J91", "Done Checks"),
            ("K91:L91", "Success %"),
            ("M91:N91", "Perfect Days"),
            ("O91:P91", "Avg Sleep"),
            ("Q91:S91", "Best Week"),
            ("T91:W91", "Top Life Domain"),
        ]
        for b_range, b_title in bridge_headers:
            ws.merge_cells(b_range)
            b_cell = ws[b_range.split(":")[0]]
            b_cell.value = b_title
            b_cell.font = Font(name=FONT_FAMILY, size=8, bold=True)
            b_cell.fill = SUBHDR_GRAY
            b_cell.alignment = Alignment(horizontal="center", vertical="center")

        for r_b in (91, 92):
            for c_b in range(1, 24):
                ws.cell(row=r_b, column=c_b).border = BORDER_STANDARD

        # Row 92: Standardized Bridge Values
        ws.row_dimensions[92].height = 16
        ws.merge_cells("A92:B92")
        ws["A92"] = f"{month_name[:3]}"
        ws["A92"].font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        ws["A92"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("C92:D92")
        ws["C92"] = year
        ws["C92"].font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        ws["C92"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("E92:F92")
        ws["E92"] = f'=COUNTA(B{first_habit_row}:B{last_habit_row})'
        ws["E92"].font = Font(name=FONT_FAMILY, size=8.5)
        ws["E92"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("G92:H92")
        ws["G92"] = f'={all_month_targets}'
        ws["G92"].font = Font(name=FONT_FAMILY, size=8.5)
        ws["G92"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("I92:J92")
        ws["I92"] = f'={all_month_ticks}'
        ws["I92"].font = Font(name=FONT_FAMILY, size=8.5)
        ws["I92"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("K92:L92")
        ws["K92"] = f'=IF(G92>0, I92/G92, 0)'
        ws["K92"].font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="166534")
        ws["K92"].alignment = Alignment(horizontal="center", vertical="center")
        ws["K92"].number_format = "0.0%"

        ws.merge_cells("M92:N92")
        ws["M92"] = f'=COUNTIF({first_day_let}{pct_row}:{last_day_letter}{pct_row}, 1)'
        ws["M92"].font = Font(name=FONT_FAMILY, size=8.5)
        ws["M92"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("O92:P92")
        ws["O92"] = f'=IFERROR(AVERAGE({first_day_let}{sleep_row}:{last_day_letter}{sleep_row}), 0)'
        ws["O92"].font = Font(name=FONT_FAMILY, size=8.5)
        ws["O92"].alignment = Alignment(horizontal="center", vertical="center")
        ws["O92"].number_format = "0.0"

        ws.merge_cells("Q92:S92")
        ws["Q92"] = '=INDEX(A76:A80, MATCH(MAX(E76:E80), E76:E80, 0))'
        ws["Q92"].font = Font(name=FONT_FAMILY, size=8.5)
        ws["Q92"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("T92:W92")
        ws["T92"] = '=INDEX(N76:N80, MATCH(MAX(S76:S80), S76:S80, 0))'
        ws["T92"].font = Font(name=FONT_FAMILY, size=8.5)
        ws["T92"].alignment = Alignment(horizontal="center", vertical="center")

        # Row 93: Explanatory Caption
        ws.row_dimensions[93].height = 14
        ws.merge_cells(f"A93:{status_col_letter}93")
        cap_cell = ws["A93"]
        cap_cell.value = "📋 Note: Reference or copy Row 92 directly into an Annual / Quarterly review dashboard to track yearly habit trends."
        cap_cell.font = Font(name=FONT_FAMILY, size=7.5, italic=True, color="4B5563")
        cap_cell.alignment = Alignment(horizontal="left", vertical="center")

        # Row 94: Spacing
        ws.row_dimensions[94].height = 8

        # --- SECTION 5: MONTHLY RETROSPECTIVE & ACTION PLAN (Rows 95–104) ---
        ws.merge_cells("A95:H95")
        ws["A95"] = "🏆 MONTHLY WINS & HIGHLIGHTS"
        ws["A95"].font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="166534")
        ws["A95"].alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells("J95:Q95")
        ws["J95"] = "⚠️ FRICTION POINTS & ROOT CAUSES"
        ws["J95"].font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="991B1B")
        ws["J95"].alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells("S95:Z95")
        ws["S95"] = "🎯 3 KEY ADJUSTMENTS FOR NEXT MONTH"
        ws["S95"].font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="1D4ED8")
        ws["S95"].alignment = Alignment(horizontal="left", vertical="center")

        # Box 1: Wins (A96:H103)
        ws.merge_cells("A96:H103")
        box_wins = ws["A96"]
        box_wins.value = "1. What went exceptionally well this month?\n2. Which habit felt effortless and consistent?\n3. Major milestone achieved:"
        box_wins.font = Font(name=FONT_FAMILY, size=8, color="374151")
        box_wins.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        for r_b in range(96, 104):
            for c_b in range(1, 9):
                ws.cell(row=r_b, column=c_b).fill = KPI_GREEN_FILL
                ws.cell(row=r_b, column=c_b).border = BORDER_NOTE

        # Box 2: Friction (J96:Q103)
        ws.merge_cells("J96:Q103")
        box_fric = ws["J96"]
        box_fric.value = "1. What caused missed or skipped days?\n2. Main triggers / energy drains encountered:\n3. Which habit had the lowest consistency and why?"
        box_fric.font = Font(name=FONT_FAMILY, size=8, color="374151")
        box_fric.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        for r_b in range(96, 104):
            for c_b in range(10, 18):
                ws.cell(row=r_b, column=c_b).fill = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
                ws.cell(row=r_b, column=c_b).border = BORDER_NOTE

        # Box 3: Adjustments (S96:Z103)
        ws.merge_cells("S96:Z103")
        box_adj = ws["S96"]
        box_adj.value = "1. Habit Stacking or time adjustment for next month:\n2. Environment design change to eliminate friction:\n3. One non-negotiable priority protocol:"
        box_adj.font = Font(name=FONT_FAMILY, size=8, color="374151")
        box_adj.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        for r_b in range(96, 104):
            for c_b in range(19, 27):
                ws.cell(row=r_b, column=c_b).fill = KPI_BLUE_FILL
                ws.cell(row=r_b, column=c_b).border = BORDER_NOTE

    output_filename = "Habit_Tracker_2026_2030.xlsx"
    wb.save(output_filename)
    print(f"Workbook successfully saved to {output_filename}")

if __name__ == "__main__":
    generate_exact_tracker_excel()
