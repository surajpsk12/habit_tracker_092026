import calendar
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import BarChart, LineChart, Reference

def build_habit_tracker():
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet

    # Typography & Styles (Modern Premium Executive Theme)
    FONT_FAMILY = "Segoe UI"
    
    HEADER_FILL = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid") # Dark Slate
    HEADER_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
    
    SUBHEADER_FILL = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    SUBHEADER_FONT = Font(name=FONT_FAMILY, size=9.5, bold=True, color="F8FAFC")
    
    TITLE_FONT = Font(name=FONT_FAMILY, size=16, bold=True, color="0F172A")
    SUBTITLE_FONT = Font(name=FONT_FAMILY, size=10, italic=True, color="475569")
    
    CARD_TITLE_FONT = Font(name=FONT_FAMILY, size=9, bold=True, color="64748B")
    CARD_VALUE_FONT = Font(name=FONT_FAMILY, size=18, bold=True, color="0F172A")
    
    SUNDAY_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    SUNDAY_HEADER_FILL = PatternFill(start_color="CBD5E1", end_color="CBD5E1", fill_type="solid")
    
    GROUP_FILLS = [
        PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid"), # Soft Red (Health)
        PatternFill(start_color="F0F9FF", end_color="F0F9FF", fill_type="solid"), # Soft Sky (Deep Work)
        PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid"), # Soft Emerald (Mindset)
        PatternFill(start_color="FEFCE8", end_color="FEFCE8", fill_type="solid"), # Soft Amber (Routine)
        PatternFill(start_color="FAF5FF", end_color="FAF5FF", fill_type="solid")  # Soft Violet (Night)
    ]
    
    # Borders
    THIN_BORDER_SIDE = Side(border_style="thin", color="CBD5E1")
    THICK_BOTTOM_SIDE = Side(border_style="medium", color="475569")
    DOUBLE_BOTTOM_SIDE = Side(border_style="double", color="1E293B")
    CARD_BORDER_SIDE = Side(border_style="thin", color="94A3B8")
    
    BORDER_ALL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
    BORDER_HEADER = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THICK_BOTTOM_SIDE)
    BORDER_TOTAL = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=DOUBLE_BOTTOM_SIDE)
    BORDER_CARD = Border(left=CARD_BORDER_SIDE, right=CARD_BORDER_SIDE, top=CARD_BORDER_SIDE, bottom=CARD_BORDER_SIDE)

    # Preset 15 Protocols (Grouped in 5 Categories)
    DEFAULT_HABITS = [
        # Category 1: Health & Fitness
        ("Health & Fitness", "Morning Workout / Gym", 25),
        ("Health & Fitness", "Drink 3.5L Water Daily", 30),
        ("Health & Fitness", "Clean Nutrition / No Junk", 28),
        
        # Category 2: Deep Work & Skill
        ("Deep Work & Skill", "3 Hours Focused Coding", 25),
        ("Deep Work & Skill", "Read 20 Pages Book", 28),
        ("Deep Work & Skill", "Project Work / Commit", 26),
        
        # Category 3: Mindset & Focus
        ("Mindset & Focus", "10 Mins Meditation / Breathwork", 28),
        ("Mindset & Focus", "Morning Sunlight (15 mins)", 30),
        ("Mindset & Focus", "Daily Journaling & Planning", 28),
        
        # Category 4: Routine & Habits
        ("Routine & Ops", "Wake Up by 6:00 AM", 26),
        ("Routine & Ops", "Cold Shower", 25),
        ("Routine & Ops", "Zero Screen First 30 Mins", 28),
        
        # Category 5: Night & Recovery
        ("Night Protocol", "No Screens After 10:00 PM", 26),
        ("Night Protocol", "Night Reflection & Review", 28),
        ("Night Protocol", "In Bed by 10:30 PM", 26),
    ]

    CATEGORIES_LIST = [
        "Health & Fitness",
        "Deep Work & Skill",
        "Mindset & Focus",
        "Routine & Ops",
        "Night Protocol",
        "Personal Growth",
        "Finance & Career"
    ]

    QUOTES = [
        "Discipline equals freedom. - Jocko Willink",
        "Win the morning, win the day. - Tim Ferriss",
        "We are what we repeatedly do. Excellence, then, is not an act, but a habit. - Aristotle",
        "You do not rise to the level of your goals. You fall to the level of your systems. - James Clear",
        "Focus on the process, results will follow.",
        "Small daily improvements over time lead to stunning results."
    ]

    # Timeline: Sep 2026 to Dec 2027 (16 Months total!)
    months_timeline = []
    # 2026: Sep (9), Oct (10), Nov (11), Dec (12)
    for m in range(9, 13):
        months_timeline.append((2026, m))
    # 2027: Jan (1) to Dec (12)
    for m in range(1, 13):
        months_timeline.append((2027, m))

    # --- 1. BUILD MASTER DASHBOARD ---
    dash_ws = wb.create_sheet(title="📊 Master Dashboard")
    dash_ws.views.sheetView[0].showGridLines = True
    
    # Title Banner
    dash_ws["A2"] = "🏆 MASTER HABIT TRACKER EXECUTIVE DASHBOARD (2026 - 2027)"
    dash_ws["A2"].font = Font(name=FONT_FAMILY, size=16, bold=True, color="0F172A")
    dash_ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    dash_ws.merge_cells("A2:R2")
    
    dash_ws["A3"] = "Real-time annual consistency analytics, category performance matrix, and trend visualization."
    dash_ws["A3"].font = SUBTITLE_FONT

    # KPI Summary Cards (Rows 5 to 7)
    kpis = [
        ("A5", "C5", "A6", "C7", "TOTAL MONTHS TRACKED", "16 Months", "F8FAFC", "0F172A"),
        ("E5", "G5", "E6", "G7", "OVERALL YEAR CONSISTENCY", "=R26", "F0FDF4", "16A34A"),
        ("I5", "K5", "I6", "K7", "TOTAL HABIT PROTOCOLS", "15 Habits", "EFF6FF", "2563EB"),
        ("M5", "P5", "M6", "P7", "TARGET SUCCESS THRESHOLD", "90% for Reward", "FEF2F2", "DC2626"),
    ]

    for top_l_t, bot_r_t, top_l_v, bot_r_v, title_text, val_text, fill_hex, text_hex in kpis:
        card_fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")
        dash_ws[top_l_t] = title_text
        dash_ws[top_l_t].font = CARD_TITLE_FONT
        dash_ws[top_l_t].alignment = Alignment(horizontal="center", vertical="center")
        dash_ws[top_l_t].fill = card_fill
        dash_ws.merge_cells(f"{top_l_t}:{bot_r_t}")

        dash_ws[top_l_v] = val_text
        dash_ws[top_l_v].font = Font(name=FONT_FAMILY, size=16, bold=True, color=text_hex)
        dash_ws[top_l_v].alignment = Alignment(horizontal="center", vertical="center")
        dash_ws[top_l_v].fill = card_fill
        if str(val_text).startswith("="):
            dash_ws[top_l_v].number_format = "0.0%"
        dash_ws.merge_cells(f"{top_l_v}:{bot_r_v}")

        # Set card borders
        col_start = openpyxl.utils.column_index_from_string(top_l_t[0])
        col_end = openpyxl.utils.column_index_from_string(bot_r_t[0])
        for r in range(5, 8):
            for c in range(col_start, col_end + 1):
                dash_ws.cell(row=r, column=c).border = BORDER_CARD

    # Section Label (Row 9)
    dash_ws["A9"] = "MONTHLY HABIT PERFORMANCE MATRIX"
    dash_ws["A9"].font = Font(name=FONT_FAMILY, size=11, bold=True, color="1E293B")

    # Table Header (Row 10)
    dash_ws["A10"] = "Habit Protocol"
    dash_ws["A10"].font = HEADER_FONT
    dash_ws["A10"].fill = HEADER_FILL
    dash_ws["A10"].alignment = Alignment(horizontal="left", vertical="center")
    dash_ws["A10"].border = BORDER_HEADER
    
    dash_ws["B10"] = "Category"
    dash_ws["B10"].font = HEADER_FONT
    dash_ws["B10"].fill = HEADER_FILL
    dash_ws["B10"].alignment = Alignment(horizontal="left", vertical="center")
    dash_ws["B10"].border = BORDER_HEADER

    col_idx = 3 # Col C
    for year, month in months_timeline:
        m_short = f"{calendar.month_name[month][:3]} '{str(year)[2:]}"
        col_letter = get_column_letter(col_idx)
        cell = dash_ws[f"{col_letter}10"]
        cell.value = m_short
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER_HEADER
        col_idx += 1

    # Avg Column
    avg_col_letter = get_column_letter(col_idx)
    dash_ws[f"{avg_col_letter}10"] = "Year Avg"
    dash_ws[f"{avg_col_letter}10"].font = HEADER_FONT
    dash_ws[f"{avg_col_letter}10"].fill = HEADER_FILL
    dash_ws[f"{avg_col_letter}10"].alignment = Alignment(horizontal="center", vertical="center")
    dash_ws[f"{avg_col_letter}10"].border = BORDER_HEADER

    # Fill Habit Rows in Dashboard (Rows 11 to 25)
    for idx, (cat, habit_name, target) in enumerate(DEFAULT_HABITS, start=11):
        dash_ws.row_dimensions[idx].height = 22
        
        dash_ws[f"A{idx}"] = habit_name
        dash_ws[f"A{idx}"].font = Font(name=FONT_FAMILY, size=9.5, bold=True)
        dash_ws[f"A{idx}"].border = BORDER_ALL
        
        dash_ws[f"B{idx}"] = cat
        dash_ws[f"B{idx}"].font = Font(name=FONT_FAMILY, size=9, color="475569")
        dash_ws[f"B{idx}"].border = BORDER_ALL
        
        c_idx = 3
        for year, month in months_timeline:
            m_name = f"{calendar.month_name[month][:3]} {year}"
            c_letter = get_column_letter(c_idx)
            cell = dash_ws[f"{c_letter}{idx}"]
            
            month_row = idx - 4 # Maps 11..25 to 7..21
            days_count = calendar.monthrange(year, month)[1]
            success_col_letter = get_column_letter(4 + days_count + 2) # Success col
            
            cell.value = f"='{m_name}'!{success_col_letter}{month_row}"
            cell.number_format = "0.0%"
            cell.font = Font(name=FONT_FAMILY, size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER_ALL
            c_idx += 1
            
        first_m_col = get_column_letter(3)
        last_m_col = get_column_letter(c_idx - 1)
        avg_cell = dash_ws[f"{avg_col_letter}{idx}"]
        avg_cell.value = f"=AVERAGE({first_m_col}{idx}:{last_m_col}{idx})"
        avg_cell.number_format = "0.0%"
        avg_cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True)
        avg_cell.alignment = Alignment(horizontal="center", vertical="center")
        avg_cell.border = BORDER_ALL

    # Dashboard Summary Row (Row 26)
    dash_ws.row_dimensions[26].height = 24
    dash_ws["A26"] = "OVERALL MONTHLY SCORE"
    dash_ws["A26"].font = Font(name=FONT_FAMILY, size=10, bold=True, color="1E293B")
    dash_ws["A26"].border = BORDER_TOTAL
    dash_ws["B26"] = "All Categories"
    dash_ws["B26"].font = Font(name=FONT_FAMILY, size=9, italic=True, color="64748B")
    dash_ws["B26"].border = BORDER_TOTAL

    c_idx = 3
    for year, month in months_timeline:
        c_letter = get_column_letter(c_idx)
        cell = dash_ws[f"{c_letter}26"]
        cell.value = f"=AVERAGE({c_letter}11:{c_letter}25)"
        cell.number_format = "0.0%"
        cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="047857")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER_TOTAL
        c_idx += 1

    avg_summary_cell = dash_ws[f"{avg_col_letter}26"]
    avg_summary_cell.value = f"=AVERAGE({first_m_col}26:{last_m_col}26)"
    avg_summary_cell.number_format = "0.0%"
    avg_summary_cell.font = Font(name=FONT_FAMILY, size=10.5, bold=True, color="047857")
    avg_summary_cell.alignment = Alignment(horizontal="center", vertical="center")
    avg_summary_cell.border = BORDER_TOTAL

    # --- Charts at the Bottom of Master Dashboard ---
    # 1. Bar Chart: Monthly Consistency % Trend
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "Monthly Habit Consistency % Trend"
    chart1.y_axis.title = "Average Success Rate"
    chart1.x_axis.title = "Month"
    chart1.width = 24
    chart1.height = 12
    chart1.legend = None

    # Data: Row 26 (OVERALL MONTHLY SCORE), cols C to R
    data1 = Reference(dash_ws, min_col=3, min_row=26, max_col=c_idx-1, max_row=26)
    cats1 = Reference(dash_ws, min_col=3, min_row=10, max_col=c_idx-1, max_row=10)
    chart1.add_data(data1, titles_from_data=False, from_rows=True)
    chart1.set_categories(cats1)
    dash_ws.add_chart(chart1, "A29")

    # 2. Line Chart: Category / Top Habits Performance
    chart2 = LineChart()
    chart2.title = "Habits Performance Comparison Across Months"
    chart2.style = 13
    chart2.y_axis.title = "Completion Rate"
    chart2.x_axis.title = "Month"
    chart2.width = 24
    chart2.height = 12

    data2 = Reference(dash_ws, min_col=2, min_row=10, max_col=c_idx-1, max_row=17) # Top 7 habits
    cats2 = Reference(dash_ws, min_col=3, min_row=10, max_col=c_idx-1, max_row=10)
    chart2.add_data(data2, titles_from_data=True, from_rows=True)
    chart2.set_categories(cats2)
    dash_ws.add_chart(chart2, "I29")

    # Dashboard Column Widths
    dash_ws.column_dimensions['A'].width = 32
    dash_ws.column_dimensions['B'].width = 22
    for c in range(3, col_idx + 1):
        c_letter = get_column_letter(c)
        dash_ws.column_dimensions[c_letter].width = 11

    # --- 2. BUILD ALL 16 MONTHLY SHEETS (Sep 2026 to Dec 2027) ---
    weekdays_abbr = ['M', 'T', 'W', 'Th', 'F', 'Sa', 'Su']

    # Data Validations
    # Dropdown for Category:
    dv_cat = DataValidation(
        type="list", 
        formula1=f'"{",".join(CATEGORIES_LIST)}"', 
        allow_blank=True,
        prompt="Select Category",
        promptTitle="Habit Category"
    )

    # Dropdown for Daily Habit Completion: "Done,Not Done"
    dv_status = DataValidation(
        type="list", 
        formula1='"Done,Not Done"', 
        allow_blank=True,
        prompt="Select Status",
        promptTitle="Habit Status"
    )

    # Dropdown for Monthly Targets:
    dv_target_status = DataValidation(
        type="list", 
        formula1='"Not Started,In Progress,Achieved"', 
        allow_blank=True
    )

    # Conditional Formatting Rules for Done / Not Done
    DONE_FILL = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid") # Soft Emerald
    DONE_FONT = Font(name=FONT_FAMILY, size=9.5, bold=True, color="166534")
    
    NOT_DONE_FILL = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid") # Soft Red
    NOT_DONE_FONT = Font(name=FONT_FAMILY, size=9.5, color="991B1B")

    rule_done = CellIsRule(operator='equal', formula=['"Done"'], fill=DONE_FILL, font=DONE_FONT)
    rule_not_done = CellIsRule(operator='equal', formula=['"Not Done"'], fill=NOT_DONE_FILL, font=NOT_DONE_FONT)

    for year, month in months_timeline:
        month_name_full = calendar.month_name[month]
        sheet_title = f"{month_name_full[:3]} {year}"
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True
        
        days_in_month = calendar.monthrange(year, month)[1]
        first_weekday = calendar.monthrange(year, month)[0] # 0 = Monday, 6 = Sunday
        
        month_weekdays = [weekdays_abbr[(first_weekday + d) % 7] for d in range(days_in_month)]
        last_day_col = 4 + days_in_month
        last_day_letter = get_column_letter(last_day_col)

        # Register validations on sheet
        ws.add_data_validation(dv_cat)
        ws.add_data_validation(dv_status)
        ws.add_data_validation(dv_target_status)

        # --- Row 2: Title & Header Banner ---
        ws["A2"] = f"HABIT TRACKER - {month_name_full.upper()} {year}"
        ws["A2"].font = TITLE_FONT
        ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells("A2:C2")

        # Name / Handle
        ws["D2"] = "@surajvansh"
        ws["D2"].font = Font(name=FONT_FAMILY, size=11, bold=True, color="475569")
        ws["D2"].alignment = Alignment(horizontal="right", vertical="center")

        # Row 3: Motivational Quote
        quote_text = QUOTES[(month - 1) % len(QUOTES)]
        ws["A3"] = f'"{quote_text}"'
        ws["A3"].font = SUBTITLE_FONT
        ws["A3"].alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(f"A3:{last_day_letter}3")

        # Row heights
        ws.row_dimensions[2].height = 25
        ws.row_dimensions[3].height = 18
        ws.row_dimensions[5].height = 20
        ws.row_dimensions[6].height = 20

        # --- Row 5 & 6: Table Headers ---
        ws["A5"] = "S.No."
        ws["A5"].font = HEADER_FONT; ws["A5"].fill = HEADER_FILL
        ws["A5"].alignment = Alignment(horizontal="center", vertical="center")
        ws["A5"].border = BORDER_HEADER
        ws.merge_cells("A5:A6")

        ws["B5"] = "Category"
        ws["B5"].font = HEADER_FONT; ws["B5"].fill = HEADER_FILL
        ws["B5"].alignment = Alignment(horizontal="center", vertical="center")
        ws["B5"].border = BORDER_HEADER
        ws.merge_cells("B5:B6")

        ws["C5"] = "Habit / Protocol"
        ws["C5"].font = HEADER_FONT; ws["C5"].fill = HEADER_FILL
        ws["C5"].alignment = Alignment(horizontal="left", vertical="center")
        ws["C5"].border = BORDER_HEADER
        ws.merge_cells("C5:C6")

        ws["D5"] = "Target"
        ws["D5"].font = HEADER_FONT; ws["D5"].fill = HEADER_FILL
        ws["D5"].alignment = Alignment(horizontal="center", vertical="center")
        ws["D5"].border = BORDER_HEADER
        ws.merge_cells("D5:D6")

        # Day Headers (Columns E to ...)
        for d in range(1, days_in_month + 1):
            col_letter = get_column_letter(4 + d)
            is_sunday = (month_weekdays[d - 1] == 'Su')
            
            # Row 5: Day Number
            cell_num = ws[f"{col_letter}5"]
            cell_num.value = d
            cell_num.font = HEADER_FONT
            cell_num.fill = SUNDAY_HEADER_FILL if is_sunday else HEADER_FILL
            cell_num.alignment = Alignment(horizontal="center", vertical="center")
            cell_num.border = BORDER_ALL
            
            # Row 6: Day Name (M, T, W...)
            cell_name = ws[f"{col_letter}6"]
            cell_name.value = month_weekdays[d - 1]
            cell_name.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="1E293B" if is_sunday else "FFFFFF")
            cell_name.fill = SUNDAY_HEADER_FILL if is_sunday else SUBHEADER_FILL
            cell_name.alignment = Alignment(horizontal="center", vertical="center")
            cell_name.border = BORDER_ALL

        # Right Headers (Total, Success %, Status)
        tot_col_letter = get_column_letter(last_day_col + 1)
        ws[f"{tot_col_letter}5"] = "Completed"
        ws[f"{tot_col_letter}5"].font = HEADER_FONT; ws[f"{tot_col_letter}5"].fill = HEADER_FILL
        ws[f"{tot_col_letter}5"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{tot_col_letter}5"].border = BORDER_HEADER
        ws.merge_cells(f"{tot_col_letter}5:{tot_col_letter}6")

        succ_col_letter = get_column_letter(last_day_col + 2)
        ws[f"{succ_col_letter}5"] = "Success %"
        ws[f"{succ_col_letter}5"].font = HEADER_FONT; ws[f"{succ_col_letter}5"].fill = HEADER_FILL
        ws[f"{succ_col_letter}5"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{succ_col_letter}5"].border = BORDER_HEADER
        ws.merge_cells(f"{succ_col_letter}5:{succ_col_letter}6")

        stat_col_letter = get_column_letter(last_day_col + 3)
        ws[f"{stat_col_letter}5"] = "Status"
        ws[f"{stat_col_letter}5"].font = HEADER_FONT; ws[f"{stat_col_letter}5"].fill = HEADER_FILL
        ws[f"{stat_col_letter}5"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{stat_col_letter}5"].border = BORDER_HEADER
        ws.merge_cells(f"{stat_col_letter}5:{stat_col_letter}6")

        first_day_letter = get_column_letter(5)

        # Apply Data Validations to Ranges
        dv_cat.add(f"B7:B21")
        dv_status.add(f"{first_day_letter}7:{last_day_letter}21")

        # --- Rows 7 to 21: 15 Habit Rows ---
        for i, (cat, habit_name, target) in enumerate(DEFAULT_HABITS, start=1):
            row = 6 + i
            ws.row_dimensions[row].height = 24
            group_idx = (i - 1) // 3
            group_fill = GROUP_FILLS[group_idx]
            
            ws[f"A{row}"] = i
            ws[f"A{row}"].alignment = Alignment(horizontal="center", vertical="center")
            ws[f"A{row}"].fill = group_fill
            ws[f"A{row}"].border = BORDER_ALL
            ws[f"A{row}"].font = Font(name=FONT_FAMILY, size=9.5, bold=True)
            
            # Category with Dropdown validation
            ws[f"B{row}"] = cat
            ws[f"B{row}"].alignment = Alignment(horizontal="left", vertical="center")
            ws[f"B{row}"].fill = group_fill
            ws[f"B{row}"].border = BORDER_ALL
            ws[f"B{row}"].font = Font(name=FONT_FAMILY, size=9, bold=True, color="334155")
            
            # Habit Name
            ws[f"C{row}"] = habit_name
            ws[f"C{row}"].alignment = Alignment(horizontal="left", vertical="center")
            ws[f"C{row}"].fill = group_fill
            ws[f"C{row}"].border = BORDER_ALL
            ws[f"C{row}"].font = Font(name=FONT_FAMILY, size=9.5, bold=True)
            
            # Target
            ws[f"D{row}"] = target
            ws[f"D{row}"].alignment = Alignment(horizontal="center", vertical="center")
            ws[f"D{row}"].fill = group_fill
            ws[f"D{row}"].border = BORDER_ALL
            ws[f"D{row}"].font = Font(name=FONT_FAMILY, size=9.5)

            # Day columns: Default to blank "" so no messy FALSE text, ready for "Done" / "Not Done" dropdown!
            for d in range(1, days_in_month + 1):
                col_letter = get_column_letter(4 + d)
                cell = ws[f"{col_letter}{row}"]
                cell.value = "" # Clean, unselected state
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name=FONT_FAMILY, size=9)
                cell.border = BORDER_ALL
                if month_weekdays[d - 1] == 'Su':
                    cell.fill = SUNDAY_FILL

            # Total Completed Formula: Counts "Done" (or TRUE for backward compatibility)
            tot_cell = ws[f"{tot_col_letter}{row}"]
            tot_cell.value = f'=COUNTIF({first_day_letter}{row}:{last_day_letter}{row}, "Done") + COUNTIF({first_day_letter}{row}:{last_day_letter}{row}, TRUE)'
            tot_cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True)
            tot_cell.alignment = Alignment(horizontal="center", vertical="center")
            tot_cell.fill = group_fill
            tot_cell.border = BORDER_ALL

            # Success % Formula
            succ_cell = ws[f"{succ_col_letter}{row}"]
            succ_cell.value = f"=IF(D{row}>0, {tot_col_letter}{row}/D{row}, 0)"
            succ_cell.number_format = "0.0%"
            succ_cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True)
            succ_cell.alignment = Alignment(horizontal="center", vertical="center")
            succ_cell.fill = group_fill
            succ_cell.border = BORDER_ALL

            # Status Tag Formula
            stat_cell = ws[f"{stat_col_letter}{row}"]
            stat_cell.value = f'=IF({succ_col_letter}{row}>=0.9, "🔥 Outstanding", IF({succ_col_letter}{row}>=0.8, "✅ On Track", "⚠️ Needs Focus"))'
            stat_cell.font = Font(name=FONT_FAMILY, size=8.5, bold=True)
            stat_cell.alignment = Alignment(horizontal="center", vertical="center")
            stat_cell.fill = group_fill
            stat_cell.border = BORDER_ALL

        # Add Conditional Formatting to Day Grid
        ws.conditional_formatting.add(f"{first_day_letter}7:{last_day_letter}21", rule_done)
        ws.conditional_formatting.add(f"{first_day_letter}7:{last_day_letter}21", rule_not_done)

        # --- Row 22: DAILY TOTAL SCORE ---
        ws.row_dimensions[22].height = 24
        ws["A22"] = "DAILY TOTAL SCORE (HABITS DONE)"
        ws["A22"].font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="1E293B")
        ws["A22"].alignment = Alignment(horizontal="right", vertical="center")
        ws["A22"].border = BORDER_TOTAL
        ws.merge_cells("A22:D22")
        ws["B22"].border = BORDER_TOTAL; ws["C22"].border = BORDER_TOTAL; ws["D22"].border = BORDER_TOTAL

        for d in range(1, days_in_month + 1):
            col_letter = get_column_letter(4 + d)
            cell = ws[f"{col_letter}22"]
            cell.value = f'=COUNTIF({col_letter}7:{col_letter}21, "Done") + COUNTIF({col_letter}7:{col_letter}21, TRUE)'
            cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="0F172A")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER_TOTAL
            if month_weekdays[d - 1] == 'Su':
                cell.fill = SUNDAY_HEADER_FILL

        ws[f"{tot_col_letter}22"] = f"=SUM({tot_col_letter}7:{tot_col_letter}21)"
        ws[f"{tot_col_letter}22"].font = Font(name=FONT_FAMILY, size=9.5, bold=True)
        ws[f"{tot_col_letter}22"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{tot_col_letter}22"].border = BORDER_TOTAL

        ws[f"{succ_col_letter}22"] = f"=AVERAGE({succ_col_letter}7:{succ_col_letter}21)"
        ws[f"{succ_col_letter}22"].number_format = "0.0%"
        ws[f"{succ_col_letter}22"].font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="047857")
        ws[f"{succ_col_letter}22"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{succ_col_letter}22"].border = BORDER_TOTAL

        ws[f"{stat_col_letter}22"] = f'=IF({succ_col_letter}22>=0.9, "🔥 EXCELLENT", IF({succ_col_letter}22>=0.8, "✅ GOOD", "⚠️ IMPROVE"))'
        ws[f"{stat_col_letter}22"].font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        ws[f"{stat_col_letter}22"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{stat_col_letter}22"].border = BORDER_TOTAL

        # --- Row 24 to 27: MONTHLY PERFORMANCE & STAKES SECTION ---
        ws.row_dimensions[24].height = 20
        ws.row_dimensions[25].height = 18
        ws.row_dimensions[26].height = 18
        ws.row_dimensions[27].height = 22

        ws["A24"] = "⚡ MONTHLY PERFORMANCE & STAKES ENGINE"
        ws["A24"].font = Font(name=FONT_FAMILY, size=10.5, bold=True, color="1E293B")
        ws.merge_cells(f"A24:{last_day_letter}24")

        ws["A25"] = "IF SUCCESS < 80% ➔ PUNISHMENT: No gaming/entertainment weekends + 50 burpees daily next month"
        ws["A25"].font = Font(name=FONT_FAMILY, size=9, bold=True, color="DC2626") # Red
        ws.merge_cells(f"A25:{last_day_letter}25")

        ws["A26"] = "IF SUCCESS > 90% ➔ REWARD: Buy new tech item / custom reward + weekend getaway trip"
        ws["A26"].font = Font(name=FONT_FAMILY, size=9, bold=True, color="16A34A") # Green
        ws.merge_cells(f"A26:{last_day_letter}26")

        ws["A27"] = f'=IF({succ_col_letter}22<0.8, "🚨 PUNISHMENT TRIGGERED - MONTHLY SUCCESS IS BELOW 80%", IF({succ_col_letter}22>=0.9, "🎉 REWARD UNLOCKED! MONTHLY SUCCESS EXCEEDS 90%", "👍 ON TRACK - KEEP PUSHING TO REACH 90% REWARD"))'
        ws["A27"].font = Font(name=FONT_FAMILY, size=10, bold=True, color="1E293B")
        ws["A27"].alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(f"A27:{last_day_letter}27")

        # --- Row 29 to 34: SLEEP TRACKING MATRIX ---
        ws.row_dimensions[29].height = 20
        ws.row_dimensions[30].height = 18
        ws.row_dimensions[31].height = 20
        ws.row_dimensions[32].height = 18
        ws.row_dimensions[34].height = 20

        ws["A29"] = "🌙 SLEEP LOGGING & RECOVERY TRACKER"
        ws["A29"].font = Font(name=FONT_FAMILY, size=10.5, bold=True, color="1E293B")
        ws.merge_cells(f"A29:{last_day_letter}29")

        ws["A30"] = "Sleep Metric"
        ws["A30"].font = HEADER_FONT; ws["A30"].fill = HEADER_FILL
        ws["A30"].alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells("A30:D30")

        for d in range(1, days_in_month + 1):
            col_letter = get_column_letter(4 + d)
            cell = ws[f"{col_letter}30"]
            cell.value = d
            cell.font = HEADER_FONT; cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws["A31"] = "Hours Slept (Enter e.g. 7.5)"
        ws["A31"].font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        ws["A31"].border = BORDER_ALL
        ws.merge_cells("A31:D31")

        for d in range(1, days_in_month + 1):
            col_letter = get_column_letter(4 + d)
            cell = ws[f"{col_letter}31"]
            cell.value = 8.0 # Default benchmark
            cell.number_format = "0.0"
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name=FONT_FAMILY, size=8.5)
            cell.border = BORDER_ALL

        ws["A32"] = "Target Hours (8.0h)"
        ws["A32"].font = Font(name=FONT_FAMILY, size=8.5, italic=True, color="475569")
        ws["A32"].border = BORDER_ALL
        ws.merge_cells("A32:D32")

        for d in range(1, days_in_month + 1):
            col_letter = get_column_letter(4 + d)
            cell = ws[f"{col_letter}32"]
            cell.value = 8.0
            cell.number_format = "0.0"
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name=FONT_FAMILY, size=8, color="64748B")
            cell.border = BORDER_ALL

        # Sleep Metrics Card (Row 34)
        ws["A34"] = "Monthly Avg Sleep Hours:"
        ws["A34"].font = Font(name=FONT_FAMILY, size=9, bold=True)
        ws.merge_cells("A34:D34")
        
        sleep_first_col = get_column_letter(5)
        sleep_last_col = get_column_letter(4 + days_in_month)
        
        ws["E34"] = f"=AVERAGE({sleep_first_col}31:{sleep_last_col}31)"
        ws["E34"].number_format = "0.0 hrs"
        ws["E34"].font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="1D4ED8")

        # --- Row 36 to 46: MONTHLY TARGETS (8 TARGETS) ---
        ws.row_dimensions[36].height = 20
        ws.row_dimensions[37].height = 20

        ws["A36"] = "🎯 MONTHLY TARGETS & MILESTONES"
        ws["A36"].font = Font(name=FONT_FAMILY, size=10.5, bold=True, color="1E293B")
        ws.merge_cells(f"A36:{last_day_letter}36")

        ws["A37"] = "S.No."
        ws["A37"].font = HEADER_FONT; ws["A37"].fill = HEADER_FILL; ws["A37"].alignment = Alignment(horizontal="center", vertical="center")
        
        ws["B37"] = "Target Goal Description"
        ws["B37"].font = HEADER_FONT; ws["B37"].fill = HEADER_FILL; ws["B37"].alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells("B37:D37")

        ws["E37"] = "Target Metric / KPI"
        ws["E37"].font = HEADER_FONT; ws["E37"].fill = HEADER_FILL; ws["E37"].alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells("E37:H37")

        ws["I37"] = "Status (Dropdown)"
        ws["I37"].font = HEADER_FONT; ws["I37"].fill = HEADER_FILL; ws["I37"].alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells("I37:K37")

        ws["L37"] = "Notes / Action Plan"
        ws["L37"].font = HEADER_FONT; ws["L37"].fill = HEADER_FILL; ws["L37"].alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells("L37:R37")

        sample_targets = [
            ("Complete 1 full coding project module", "1 Module", "In Progress", "Focus on clean architecture"),
            ("Read 1 self-improvement book", "1 Book", "Not Started", "Atomic Habits or Deep Work"),
            ("Hit 25 days workout consistency", "25 Days", "In Progress", "Track via habit grid above"),
            ("Maintain average 7.5+ hrs sleep", "7.5 Hours", "In Progress", "Sleep grid benchmark"),
            ("Zero missed morning routines", "100% Routine", "In Progress", "No morning screen time"),
            ("Save 20% of monthly income", "20% Savings", "Not Started", "Auto-transfer to investment"),
            ("Complete 10 deep focus study sessions", "10 Sessions", "Not Started", "Pomodoro method 90m"),
            ("Review and refine habit system", "1 Review", "Not Started", "End of month reflection"),
        ]

        # Add target validation to column I
        dv_target_status.add("I38:I45")

        for idx, (t_desc, t_metric, t_status, t_notes) in enumerate(sample_targets, start=38):
            ws.row_dimensions[idx].height = 20
            t_no = idx - 37
            ws[f"A{idx}"] = t_no
            ws[f"A{idx}"].alignment = Alignment(horizontal="center", vertical="center")
            ws[f"A{idx}"].font = Font(name=FONT_FAMILY, size=8.5, bold=True)
            ws[f"A{idx}"].border = BORDER_ALL

            ws["B" + str(idx)] = t_desc
            ws["B" + str(idx)].font = Font(name=FONT_FAMILY, size=9)
            ws["B" + str(idx)].border = BORDER_ALL
            ws.merge_cells(f"B{idx}:D{idx}")

            ws["E" + str(idx)] = t_metric
            ws["E" + str(idx)].font = Font(name=FONT_FAMILY, size=8.5, italic=True)
            ws["E" + str(idx)].border = BORDER_ALL
            ws.merge_cells(f"E{idx}:H{idx}")

            ws["I" + str(idx)] = t_status
            ws["I" + str(idx)].font = Font(name=FONT_FAMILY, size=8.5, bold=True)
            ws["I" + str(idx)].alignment = Alignment(horizontal="center", vertical="center")
            ws["I" + str(idx)].border = BORDER_ALL
            ws.merge_cells(f"I{idx}:K{idx}")

            ws["L" + str(idx)] = t_notes
            ws["L" + str(idx)].font = Font(name=FONT_FAMILY, size=8, color="475569")
            ws["L" + str(idx)].border = BORDER_ALL
            ws.merge_cells(f"L{idx}:R{idx}")

        # --- Row 48 to 54: MONTHLY REFLECTION & NOTES ---
        ws.row_dimensions[48].height = 20
        ws.row_dimensions[49].height = 18
        ws.row_dimensions[50].height = 40

        ws["A48"] = "📝 MONTHLY REFLECTION & RETROSPECTIVE"
        ws["A48"].font = Font(name=FONT_FAMILY, size=10.5, bold=True, color="1E293B")
        ws.merge_cells(f"A48:{last_day_letter}48")

        reflection_boxes = [
            ("A49", "A49:K49", "A50", "A50:K53", "🏆 KEY WINS & VICTORIES", "1. Maintained high workout consistency.\n2. Built advanced digital habit tracker system."),
            ("M49", "M49:W49", "M50", "M50:W53", "⚠️ FRICTION POINTS & CHALLENGES", "1. Late night screen usage on weekends.\n2. Hydration goal missed on travel days."),
            ("Y49", "Y49:AI49", "Y50", "Y50:AI53", "🚀 NEXT MONTH ACTION PLAN", "1. Phone outside bedroom after 10 PM.\n2. Prepare water bottle every morning.")
        ]

        for head_cell, head_range, body_cell, body_range, box_title, sample_text in reflection_boxes:
            ws[head_cell] = box_title
            ws[head_cell].font = SUBHEADER_FONT
            ws[head_cell].fill = SUBHEADER_FILL
            ws[head_cell].alignment = Alignment(horizontal="left", vertical="center")
            ws.merge_cells(head_range)

            ws[body_cell] = sample_text
            ws[body_cell].font = Font(name=FONT_FAMILY, size=8.5, color="334155")
            ws[body_cell].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            ws.merge_cells(body_range)

        # Generous Column Widths so text is never cut off!
        ws.column_dimensions['A'].width = 7
        ws.column_dimensions['B'].width = 24  # Category dropdown fits completely
        ws.column_dimensions['C'].width = 36  # Habit Name fits completely
        ws.column_dimensions['D'].width = 11  # Target fits completely

        # Day columns: width 8 so "Done", "Not Done" and dropdown arrow fit comfortably
        for d in range(1, days_in_month + 1):
            col_letter = get_column_letter(4 + d)
            ws.column_dimensions[col_letter].width = 8.5

        ws.column_dimensions[tot_col_letter].width = 14
        ws.column_dimensions[succ_col_letter].width = 14
        ws.column_dimensions[stat_col_letter].width = 18

    # Save Workbook
    filename = "Habit_Tracker_2026_2027.xlsx"
    wb.save(filename)
    print(f"Successfully generated enhanced '{filename}' with {len(wb.sheetnames)} sheets.")

if __name__ == "__main__":
    build_habit_tracker()
