import calendar
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import LineChart, Reference

def generate_exact_tracker_excel():
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Typography & Styles
    FONT_FAMILY = "Arial"
    
    # Border definitions
    THIN_BORDER_GRAY = Side(border_style="thin", color="000000")
    MEDIUM_BORDER = Side(border_style="medium", color="000000")
    
    BORDER_STANDARD = Border(
        left=THIN_BORDER_GRAY, right=THIN_BORDER_GRAY,
        top=THIN_BORDER_GRAY, bottom=THIN_BORDER_GRAY
    )
    BORDER_SUNDAY = Border(
        left=MEDIUM_BORDER, right=MEDIUM_BORDER,
        top=THIN_BORDER_GRAY, bottom=THIN_BORDER_GRAY
    )
    BORDER_NOTE = Border(
        left=Side(border_style="thin", color="D1D5DB"),
        right=Side(border_style="thin", color="D1D5DB"),
        top=Side(border_style="thin", color="D1D5DB"),
        bottom=Side(border_style="thin", color="D1D5DB")
    )

    # 5 Group colors for 20 protocol rows (4 rows per group)
    GROUP_COLORS_HEX = [
        "FFEBEB",  # Soft Red/Pink: RGB(255, 235, 235) - Rows 1-4
        "EBF5FF",  # Soft Sky/Blue: RGB(235, 245, 255) - Rows 5-8
        "EBFFEB",  # Soft Mint/Green: RGB(235, 255, 235) - Rows 9-12
        "FFFFEB",  # Soft Cream/Yellow: RGB(255, 255, 235) - Rows 13-16
        "F5EBFF",  # Soft Lavender/Purple: RGB(245, 235, 255) - Rows 17-20
    ]
    
    # Sunday deeper tint in protocol rows
    SUNDAY_GROUP_COLORS_HEX = [
        "E6D2D2",  # RGB(230, 210, 210)
        "D2DCE6",  # RGB(210, 220, 230)
        "D2E6D2",  # RGB(210, 230, 210)
        "E6E6D2",  # RGB(230, 230, 210)
        "DCD2E6",  # RGB(220, 210, 230)
    ]

    SUNDAY_HEADER_FILL = PatternFill(start_color="D2D2D2", end_color="D2D2D2", fill_type="solid") # RGB(210, 210, 210)
    NOTES_FILL = PatternFill(start_color="FFFFE6", end_color="FFFFE6", fill_type="solid") # RGB(255, 255, 230)
    WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    MONTH_QUOTES = {
        9: "Win the morning, win the day.",
        10: "Discipline equals freedom.",
        11: "Focus on the process.",
        12: "Stay consistent."
    }

    # Only September 2026 to December 2026
    target_months = [
        (2026, 9),
        (2026, 10),
        (2026, 11),
        (2026, 12)
    ]

    weekdays_abbr = ['M', 'T', 'W', 'Th', 'F', 'Sa', 'Su']

    # Target Dropdown values: 1 to 31
    target_dropdown_str = ",".join(str(i) for i in range(1, 32))
    # Sleep Hours Dropdown values: 1 to 24
    sleep_dropdown_str = ",".join(str(i) for i in range(1, 25))

    for year, month in target_months:
        month_name = calendar.month_name[month]
        sheet_title = f"{month_name[:3]} {year}"
        ws = wb.create_sheet(title=sheet_title)
        
        # Gridlines and Page setup
        ws.views.sheetView[0].showGridLines = True
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1

        days_in_month = calendar.monthrange(year, month)[1]
        first_weekday = calendar.monthrange(year, month)[0] # 0 = Monday, 6 = Sunday
        month_weekdays = [weekdays_abbr[(first_weekday + d) % 7] for d in range(days_in_month)]

        first_day_col = 4
        last_day_col = 3 + days_in_month
        success_col = last_day_col + 1

        first_day_let = get_column_letter(first_day_col)
        last_day_letter = get_column_letter(last_day_col)
        success_col_letter = get_column_letter(success_col)

        # Set Column Widths:
        # Increase day columns to 4.8 so numbers selected from dropdowns (1-24) are clearly visible with the arrow!
        ws.column_dimensions['A'].width = 6.5
        ws.column_dimensions['B'].width = 28
        ws.column_dimensions['C'].width = 8.5
        for col_idx in range(first_day_col, last_day_col + 1):
            c_let = get_column_letter(col_idx)
            ws.column_dimensions[c_let].width = 4.8
        ws.column_dimensions[success_col_letter].width = 11.5

        # Row 1: Top spacing
        ws.row_dimensions[1].height = 8

        # --- Row 2: Title, Quote, and @surajvansh ---
        ws.row_dimensions[2].height = 24
        ws["A2"] = "Habit Tracker"
        ws["A2"].font = Font(name=FONT_FAMILY, size=18, bold=True, color="000000")
        ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

        quote_text = MONTH_QUOTES.get(month, "Win the morning, win the day.")
        quote_start_col = 8
        quote_end_col = last_day_col - 5
        ws.merge_cells(start_row=2, start_column=quote_start_col, end_row=2, end_column=quote_end_col)
        quote_cell = ws.cell(row=2, column=quote_start_col)
        quote_cell.value = f'"{quote_text}"'
        quote_cell.font = Font(name=FONT_FAMILY, size=11, bold=True, color="000000")
        quote_cell.alignment = Alignment(horizontal="center", vertical="center")

        handle_start_col = last_day_col - 4
        ws.merge_cells(start_row=2, start_column=handle_start_col, end_row=2, end_column=success_col)
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
        # REPOSITIONED TO TOP (BEFORE TABLE):
        # 1. Punishment & Success (Stakes)
        # 2. Monthly Targets
        # 3. Notes
        # =========================================================================

        # --- Row 5 & 6: Punishment & Success (Reward) ---
        ws.row_dimensions[5].height = 15
        ws.row_dimensions[6].height = 15

        ws.merge_cells(f"A5:{success_col_letter}5")
        punish_cell = ws["A5"]
        punish_cell.value = "IF SUCCESS < 80%, PUNISHMENT: ________________________________________________"
        punish_cell.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="C80000")
        punish_cell.alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells(f"A6:{success_col_letter}6")
        reward_cell = ws["A6"]
        reward_cell.value = "IF SUCCESS > 90%, REWARD: __________________________________________________"
        reward_cell.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="007800")
        reward_cell.alignment = Alignment(horizontal="left", vertical="center")

        # Row 7: Spacing
        ws.row_dimensions[7].height = 5

        # --- Row 8: Monthly Targets Label ---
        ws.row_dimensions[8].height = 16
        ws["A8"] = "MONTHLY TARGETS"
        ws["A8"].font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        ws["A8"].alignment = Alignment(horizontal="left", vertical="center")

        # Rows 9 to 12: Monthly Targets (2 columns: 1-4 on left, 5-8 on right)
        col2_target_start = 16
        for i in range(1, 5):
            t_row = 8 + i
            ws.row_dimensions[t_row].height = 13

            # Column 1 (Targets 1-4)
            ws.merge_cells(start_row=t_row, start_column=1, end_row=t_row, end_column=col2_target_start - 2)
            t1_cell = ws.cell(row=t_row, column=1)
            t1_cell.value = f"{i}. ____________________ : ____________________"
            t1_cell.font = Font(name=FONT_FAMILY, size=8)
            t1_cell.alignment = Alignment(horizontal="left", vertical="center")

            # Column 2 (Targets 5-8)
            i2 = i + 4
            ws.merge_cells(start_row=t_row, start_column=col2_target_start, end_row=t_row, end_column=last_day_col)
            t2_cell = ws.cell(row=t_row, column=col2_target_start)
            t2_cell.value = f"{i2}. ____________________ : ____________________"
            t2_cell.font = Font(name=FONT_FAMILY, size=8)
            t2_cell.alignment = Alignment(horizontal="left", vertical="center")

        # Row 13: Spacing
        ws.row_dimensions[13].height = 5

        # --- Row 14: Notes Section Label ---
        ws.row_dimensions[14].height = 15
        ws["A14"] = "NOTES:"
        ws["A14"].font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        ws["A14"].alignment = Alignment(horizontal="left", vertical="center")

        # Rows 15 to 17: Notes Yellow Box (Top Before Table)
        for r_note in range(15, 18):
            ws.row_dimensions[r_note].height = 13
        ws.merge_cells(f"A15:{last_day_letter}17")
        note_box = ws["A15"]
        note_box.fill = NOTES_FILL
        note_box.font = Font(name=FONT_FAMILY, size=8.5)
        note_box.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        for r in range(15, 18):
            for c in range(1, last_day_col + 1):
                ws.cell(row=r, column=c).fill = NOTES_FILL
                ws.cell(row=r, column=c).border = BORDER_NOTE

        # Row 18: Spacing before main table
        ws.row_dimensions[18].height = 8

        # =========================================================================
        # MAIN PROTOCOLS TABLE
        # =========================================================================

        # --- Rows 19 & 20: Protocol Table Header ---
        ws.row_dimensions[19].height = 15
        ws.row_dimensions[20].height = 15

        # Merge S.No.
        ws.merge_cells("A19:A20")
        ws["A19"] = "S.No."
        ws["A19"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws["A19"].alignment = Alignment(horizontal="center", vertical="center")
        for r in (19, 20):
            ws[f"A{r}"].border = BORDER_STANDARD

        # Merge Protocols
        ws.merge_cells("B19:B20")
        ws["B19"] = "Protocols"
        ws["B19"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws["B19"].alignment = Alignment(horizontal="center", vertical="center")
        for r in (19, 20):
            ws[f"B{r}"].border = BORDER_STANDARD

        # Merge Target
        ws.merge_cells("C19:C20")
        ws["C19"] = "Target"
        ws["C19"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws["C19"].alignment = Alignment(horizontal="center", vertical="center")
        for r in (19, 20):
            ws[f"C{r}"].border = BORDER_STANDARD

        # Day Number and Weekday Headers
        for d in range(1, days_in_month + 1):
            col = first_day_col + (d - 1)
            c_let = get_column_letter(col)
            is_sunday = (month_weekdays[d - 1] == 'Su')
            header_fill = SUNDAY_HEADER_FILL if is_sunday else WHITE_FILL

            # Row 19: Day number
            cell_d = ws[f"{c_let}19"]
            cell_d.value = d
            cell_d.font = Font(name=FONT_FAMILY, size=8, bold=True)
            cell_d.alignment = Alignment(horizontal="center", vertical="center")
            cell_d.fill = header_fill
            cell_d.border = BORDER_SUNDAY if is_sunday else BORDER_STANDARD

            # Row 20: Weekday abbreviation
            cell_w = ws[f"{c_let}20"]
            cell_w.value = month_weekdays[d - 1]
            cell_w.font = Font(name=FONT_FAMILY, size=7, bold=is_sunday)
            cell_w.alignment = Alignment(horizontal="center", vertical="center")
            cell_w.fill = header_fill
            cell_w.border = BORDER_SUNDAY if is_sunday else BORDER_STANDARD

        # Merge Success %
        ws.merge_cells(f"{success_col_letter}19:{success_col_letter}20")
        ws[f"{success_col_letter}19"] = "Success %"
        ws[f"{success_col_letter}19"].font = Font(name=FONT_FAMILY, size=7.5, bold=True)
        ws[f"{success_col_letter}19"].alignment = Alignment(horizontal="center", vertical="center")
        for r in (19, 20):
            ws[f"{success_col_letter}{r}"].border = BORDER_STANDARD

        # --- Rows 21 to 40: 20 Protocol Rows (5 color groups of 4 rows each) ---
        first_habit_row = 21
        total_habits = 20
        last_habit_row = first_habit_row + total_habits - 1 # 40

        # Setup DataValidation for Target dropdown (1 to 31)
        dv_target = DataValidation(
            type="list",
            formula1=f'"{target_dropdown_str}"',
            allow_blank=True,
            promptTitle="Target Days",
            prompt="Select monthly target (1-31)"
        )
        ws.add_data_validation(dv_target)
        dv_target.add(f"C{first_habit_row}:C{last_habit_row}")

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

            # Col C: Target (has dropdown 1-31)
            cell_tgt = ws[f"C{row}"]
            cell_tgt.font = Font(name=FONT_FAMILY, size=8)
            cell_tgt.alignment = Alignment(horizontal="center", vertical="center")
            cell_tgt.fill = row_fill
            cell_tgt.border = BORDER_STANDARD

            # Cols D..: Day cells
            for d in range(1, days_in_month + 1):
                col = first_day_col + (d - 1)
                c_let = get_column_letter(col)
                is_sunday = (month_weekdays[d - 1] == 'Su')
                cell_day = ws[f"{c_let}{row}"]
                cell_day.font = Font(name=FONT_FAMILY, size=8)
                cell_day.alignment = Alignment(horizontal="center", vertical="center")
                cell_day.fill = row_sunday_fill if is_sunday else row_fill
                cell_day.border = BORDER_SUNDAY if is_sunday else BORDER_STANDARD

            # Col Success %: Dynamic formula supporting numbers (1) and checkmarks (✓, x)
            cell_succ = ws[f"{success_col_letter}{row}"]
            calc_expr = f'(COUNT({first_day_let}{row}:{last_day_letter}{row}) + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "✓") + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "x") + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "X"))'
            cell_succ.value = f'=IF(ISBLANK(B{row}), "", IF(ISNUMBER(C{row}), IF(C{row}>0, {calc_expr}/C{row}, 0), {calc_expr}/{days_in_month}))'
            cell_succ.font = Font(name=FONT_FAMILY, size=8)
            cell_succ.alignment = Alignment(horizontal="center", vertical="center")
            cell_succ.fill = row_fill
            cell_succ.border = BORDER_STANDARD
            cell_succ.number_format = "0.0%"

        # --- Row 41: DAILY TOTAL SCORE ---
        score_row = 41
        ws.row_dimensions[score_row].height = 20
        
        ws.merge_cells(f"A{score_row}:C{score_row}")
        score_label = ws[f"A{score_row}"]
        score_label.value = "DAILY TOTAL SCORE"
        score_label.font = Font(name=FONT_FAMILY, size=8, bold=True)
        score_label.alignment = Alignment(horizontal="right", vertical="center")
        for c in ("A", "B", "C"):
            ws[f"{c}{score_row}"].border = BORDER_STANDARD

        for d in range(1, days_in_month + 1):
            col = first_day_col + (d - 1)
            c_let = get_column_letter(col)
            cell_score = ws[f"{c_let}{score_row}"]
            cell_score.value = f'=IF((COUNT({c_let}{first_habit_row}:{c_let}{last_habit_row}) + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "✓") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "x") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "X")) > 0, SUM({c_let}{first_habit_row}:{c_let}{last_habit_row}) + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "✓") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "x") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "X"), "")'
            cell_score.font = Font(name=FONT_FAMILY, size=8, bold=True)
            cell_score.alignment = Alignment(horizontal="center", vertical="center")
            cell_score.border = BORDER_STANDARD

        succ_summary = ws[f"{success_col_letter}{score_row}"]
        succ_summary.value = f'=IF(COUNT({success_col_letter}{first_habit_row}:{success_col_letter}{last_habit_row}) > 0, AVERAGE({success_col_letter}{first_habit_row}:{success_col_letter}{last_habit_row}), "")'
        succ_summary.font = Font(name=FONT_FAMILY, size=8, bold=True)
        succ_summary.alignment = Alignment(horizontal="center", vertical="center")
        succ_summary.border = BORDER_STANDARD
        succ_summary.number_format = "0.0%"

        # Row 42: Spacing
        ws.row_dimensions[42].height = 6

        # =========================================================================
        # SLEEP TRACKING SECTION & LIVE CHART
        # =========================================================================

        # --- Sleep Tracking Header (Rows 43 & 44) ---
        ws.row_dimensions[43].height = 15
        ws.row_dimensions[44].height = 15

        ws.merge_cells("A43:C44")
        sleep_hdr = ws["A43"]
        sleep_hdr.value = "Sleep Tracking"
        sleep_hdr.font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        sleep_hdr.alignment = Alignment(horizontal="center", vertical="center")
        for r in (43, 44):
            for c in ("A", "B", "C"):
                ws[f"{c}{r}"].border = BORDER_STANDARD

        for d in range(1, days_in_month + 1):
            col = first_day_col + (d - 1)
            c_let = get_column_letter(col)
            is_sunday = (month_weekdays[d - 1] == 'Su')
            header_fill = SUNDAY_HEADER_FILL if is_sunday else WHITE_FILL

            # Row 43: Day number
            cell_d = ws[f"{c_let}43"]
            cell_d.value = d
            cell_d.font = Font(name=FONT_FAMILY, size=8, bold=True)
            cell_d.alignment = Alignment(horizontal="center", vertical="center")
            cell_d.fill = header_fill
            cell_d.border = BORDER_SUNDAY if is_sunday else BORDER_STANDARD

            # Row 44: Weekday
            cell_w = ws[f"{c_let}44"]
            cell_w.value = month_weekdays[d - 1]
            cell_w.font = Font(name=FONT_FAMILY, size=7, bold=is_sunday)
            cell_w.alignment = Alignment(horizontal="center", vertical="center")
            cell_w.fill = header_fill
            cell_w.border = BORDER_SUNDAY if is_sunday else BORDER_STANDARD

        # --- Row 45: Sleep Hours with Dropdown (1 to 24) for each day ---
        # Height 20 and width 4.8 ensure numbers (1 to 24) are completely clear & visible like Target box!
        sleep_row = 45
        ws.row_dimensions[sleep_row].height = 20

        ws.merge_cells(f"A{sleep_row}:C{sleep_row}")
        sleep_lbl = ws[f"A{sleep_row}"]
        sleep_lbl.value = "Sleep Hours (1-24)"
        sleep_lbl.font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        sleep_lbl.alignment = Alignment(horizontal="center", vertical="center")
        for c in ("A", "B", "C"):
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
            cell_s.font = Font(name=FONT_FAMILY, size=9, bold=True, color="000000")
            cell_s.alignment = Alignment(horizontal="center", vertical="center")
            cell_s.number_format = "0"
            if is_sunday:
                cell_s.fill = SUNDAY_HEADER_FILL
                cell_s.border = BORDER_SUNDAY
            else:
                cell_s.border = BORDER_STANDARD

        # --- Sleep Tracking Line Chart (Row 47 to Row 60) ---
        chart = LineChart()
        chart.title = "Daily Sleep Trend (Hours vs Day)"
        chart.style = 13
        chart.y_axis.title = "Hours"
        chart.x_axis.title = "Day"
        chart.y_axis.scaling.min = 0
        chart.y_axis.scaling.max = 24
        chart.width = 24
        chart.height = 7.5
        chart.legend = None

        # Data from Row 45 (Sleep Hours)
        data = Reference(ws, min_col=first_day_col, min_row=sleep_row, max_col=last_day_col, max_row=sleep_row)
        # Categories from Row 43 (Day Numbers)
        categories = Reference(ws, min_col=first_day_col, min_row=43, max_col=last_day_col, max_row=43)

        chart.add_data(data, titles_from_data=False, from_rows=True)
        chart.set_categories(categories)

        if chart.series:
            series = chart.series[0]
            series.graphicalProperties.line.solidFill = "3B82F6"
            series.graphicalProperties.line.width = 25000
            series.marker.symbol = "circle"
            series.marker.size = 5
            series.marker.graphicalProperties.solidFill = "1D4ED8"

        ws.add_chart(chart, "A47")

    output_filename = "Habit_Tracker_Sep_Dec_2026.xlsx"
    wb.save(output_filename)
    print(f"Workbook successfully saved to {output_filename}")

if __name__ == "__main__":
    generate_exact_tracker_excel()
