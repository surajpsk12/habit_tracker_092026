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
        done_tgt_col = last_day_col + 2  # New column right after Success %

        first_day_let = get_column_letter(first_day_col)
        last_day_letter = get_column_letter(last_day_col)
        success_col_letter = get_column_letter(success_col)
        done_tgt_col_letter = get_column_letter(done_tgt_col)

        # Set Column Widths:
        ws.column_dimensions['A'].width = 6.5
        ws.column_dimensions['B'].width = 28
        ws.column_dimensions['C'].width = 8.5
        for col_idx in range(first_day_col, last_day_col + 1):
            c_let = get_column_letter(col_idx)
            ws.column_dimensions[c_let].width = 4.8
        ws.column_dimensions[success_col_letter].width = 11.5
        ws.column_dimensions[done_tgt_col_letter].width = 13.5

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
        ws.merge_cells(start_row=2, start_column=handle_start_col, end_row=2, end_column=done_tgt_col)
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
        # 1. Monthly Targets
        # 2. Notes
        # =========================================================================

        # --- Row 5: Monthly Targets Label ---
        ws.row_dimensions[5].height = 16
        ws["A5"] = "MONTHLY TARGETS"
        ws["A5"].font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        ws["A5"].alignment = Alignment(horizontal="left", vertical="center")

        # Rows 6 to 9: Monthly Targets (2 columns: 1-4 on left, 5-8 on right)
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
            ws.merge_cells(start_row=t_row, start_column=col2_target_start, end_row=t_row, end_column=done_tgt_col)
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
        ws.merge_cells(f"A12:{done_tgt_col_letter}14")
        note_box = ws["A12"]
        note_box.fill = NOTES_FILL
        note_box.font = Font(name=FONT_FAMILY, size=8.5)
        note_box.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        for r in range(12, 15):
            for c in range(1, done_tgt_col + 1):
                ws.cell(row=r, column=c).fill = NOTES_FILL
                ws.cell(row=r, column=c).border = BORDER_NOTE

        # Row 15: Spacing before main table
        ws.row_dimensions[15].height = 8

        # =========================================================================
        # MAIN PROTOCOLS TABLE
        # =========================================================================

        # --- Rows 16 & 17: Protocol Table Header ---
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

        # Merge Target
        ws.merge_cells("C16:C17")
        ws["C16"] = "Target"
        ws["C16"].font = Font(name=FONT_FAMILY, size=8, bold=True)
        ws["C16"].alignment = Alignment(horizontal="center", vertical="center")
        for r in (16, 17):
            ws[f"C{r}"].border = BORDER_STANDARD

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

        # Merge Done / Target (New Column right after Success %)
        ws.merge_cells(f"{done_tgt_col_letter}16:{done_tgt_col_letter}17")
        ws[f"{done_tgt_col_letter}16"] = "Done / Target"
        ws[f"{done_tgt_col_letter}16"].font = Font(name=FONT_FAMILY, size=7.5, bold=True)
        ws[f"{done_tgt_col_letter}16"].alignment = Alignment(horizontal="center", vertical="center")
        for r in (16, 17):
            ws[f"{done_tgt_col_letter}{r}"].border = BORDER_STANDARD

        # --- Rows 18 to 37: 20 Protocol Rows (5 color groups of 4 rows each) ---
        first_habit_row = 18
        total_habits = 20
        last_habit_row = first_habit_row + total_habits - 1 # 37

        # DataValidation for Target dropdown (1 to 31)
        dv_target = DataValidation(
            type="list",
            formula1=f'"{target_dropdown_str}"',
            allow_blank=True,
            promptTitle="Target Days",
            prompt="Select monthly target (1-31)"
        )
        ws.add_data_validation(dv_target)
        dv_target.add(f"C{first_habit_row}:C{last_habit_row}")

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

            # Col C: Target (has dropdown 1-31)
            cell_tgt = ws[f"C{row}"]
            cell_tgt.font = Font(name=FONT_FAMILY, size=8)
            cell_tgt.alignment = Alignment(horizontal="center", vertical="center")
            cell_tgt.fill = row_fill
            cell_tgt.border = BORDER_STANDARD

            # Cols D..: Day cells with dropdown ✓ / ✗
            for d in range(1, days_in_month + 1):
                col = first_day_col + (d - 1)
                c_let = get_column_letter(col)
                is_sunday = (month_weekdays[d - 1] == 'Su')
                cell_day = ws[f"{c_let}{row}"]
                cell_day.font = Font(name=FONT_FAMILY, size=9, bold=True)
                cell_day.alignment = Alignment(horizontal="center", vertical="center")
                cell_day.fill = row_sunday_fill if is_sunday else row_fill
                cell_day.border = BORDER_SUNDAY if is_sunday else BORDER_STANDARD

            # Expressions for Row calculations
            row_ticks = f'(COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "✓") + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "✔") + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "v") + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "V"))'
            row_eval = f'({row_ticks} + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "✗") + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "x") + COUNTIF({first_day_let}{row}:{last_day_letter}{row}, "X"))'

            # Col Success %:
            cell_succ = ws[f"{success_col_letter}{row}"]
            cell_succ.value = f'=IF(ISBLANK(B{row}), "", IF(ISNUMBER(C{row}), IF(C{row}>0, {row_ticks}/C{row}, 0), IF({row_eval}>0, {row_ticks}/{row_eval}, "")))'
            cell_succ.font = Font(name=FONT_FAMILY, size=8)
            cell_succ.alignment = Alignment(horizontal="center", vertical="center")
            cell_succ.fill = row_fill
            cell_succ.border = BORDER_STANDARD
            cell_succ.number_format = "0.0%"

            # Col Done / Target: Shows exact number of days done vs target at denominator
            cell_done_tgt = ws[f"{done_tgt_col_letter}{row}"]
            cell_done_tgt.value = f'=IF(ISBLANK(B{row}), "", IF(ISNUMBER(C{row}), IF(C{row}>0, {row_ticks} & " / " & C{row}, {row_ticks} & " / 0"), IF({row_eval}>0, {row_ticks} & " / " & {row_eval}, "")))'
            cell_done_tgt.font = Font(name=FONT_FAMILY, size=8, bold=True)
            cell_done_tgt.alignment = Alignment(horizontal="center", vertical="center")
            cell_done_tgt.fill = row_fill
            cell_done_tgt.border = BORDER_STANDARD

        # --- Row 38: DAILY TOTAL SCORE ---
        score_row = 38
        ws.row_dimensions[score_row].height = 18
        
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
            
            day_ticks = f'(COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "✓") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "✔") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "v") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "V"))'
            day_eval = f'({day_ticks} + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "✗") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "x") + COUNTIF({c_let}{first_habit_row}:{c_let}{last_habit_row}, "X"))'
            
            # Displays tasks done on that day (or done vs evaluated)
            cell_score.value = f'=IF({day_eval} > 0, {day_ticks}, "")'
            cell_score.font = Font(name=FONT_FAMILY, size=8, bold=True)
            cell_score.alignment = Alignment(horizontal="center", vertical="center")
            cell_score.border = BORDER_STANDARD

        # Row 38 Success Column (Overall Month Success Rate)
        all_month_ticks = f'(COUNTIF({first_day_let}{first_habit_row}:{last_day_letter}{last_habit_row}, "✓") + COUNTIF({first_day_let}{first_habit_row}:{last_day_letter}{last_habit_row}, "✔") + COUNTIF({first_day_let}{first_habit_row}:{last_day_letter}{last_habit_row}, "v") + COUNTIF({first_day_let}{first_habit_row}:{last_day_letter}{last_habit_row}, "V"))'
        all_month_eval = f'({all_month_ticks} + COUNTIF({first_day_let}{first_habit_row}:{last_day_letter}{last_habit_row}, "✗") + COUNTIF({first_day_let}{first_habit_row}:{last_day_letter}{last_habit_row}, "x") + COUNTIF({first_day_let}{first_habit_row}:{last_day_letter}{last_habit_row}, "X"))'
        all_month_targets = f'SUM(C{first_habit_row}:C{last_habit_row})'

        score_summary = ws[f"{success_col_letter}{score_row}"]
        score_summary.value = f'=IF({all_month_ticks} > 0, IF({all_month_targets} > 0, {all_month_ticks} / {all_month_targets}, {all_month_ticks} / {all_month_eval}), "")'
        score_summary.font = Font(name=FONT_FAMILY, size=8, bold=True)
        score_summary.alignment = Alignment(horizontal="center", vertical="center")
        score_summary.border = BORDER_STANDARD
        score_summary.number_format = "0.0%"

        # Row 38 Done / Target Column: Total Done vs Total Target in Month
        done_tgt_summary = ws[f"{done_tgt_col_letter}{score_row}"]
        done_tgt_summary.value = f'=IF({all_month_ticks} > 0, {all_month_ticks} & " / " & IF({all_month_targets} > 0, {all_month_targets}, {all_month_eval}), "")'
        done_tgt_summary.font = Font(name=FONT_FAMILY, size=8, bold=True)
        done_tgt_summary.alignment = Alignment(horizontal="center", vertical="center")
        done_tgt_summary.border = BORDER_STANDARD

        # --- Row 39: DAILY SUCCESS % ---
        pct_row = 39
        ws.row_dimensions[pct_row].height = 18

        ws.merge_cells(f"A{pct_row}:C{pct_row}")
        pct_label = ws[f"A{pct_row}"]
        pct_label.value = "DAILY SUCCESS %"
        pct_label.font = Font(name=FONT_FAMILY, size=8, bold=True)
        pct_label.alignment = Alignment(horizontal="right", vertical="center")
        for c in ("A", "B", "C"):
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

        pct_done_tgt = ws[f"{done_tgt_col_letter}{pct_row}"]
        pct_done_tgt.value = ""
        pct_done_tgt.border = BORDER_STANDARD

        # =========================================================================
        # BETWEEN PROTOCOL TABLE AND SLEEP TABLE:
        # Success and Punishment (Stakes) + Two-row gap
        # =========================================================================
        ws.row_dimensions[40].height = 6

        # --- Rows 41 & 42: Punishment & Success (Reward) ---
        ws.row_dimensions[41].height = 15
        ws.row_dimensions[42].height = 15

        ws.merge_cells(f"A41:{done_tgt_col_letter}41")
        punish_cell = ws["A41"]
        punish_cell.value = "IF SUCCESS < 80%, PUNISHMENT: ________________________________________________"
        punish_cell.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="C80000")
        punish_cell.alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells(f"A42:{done_tgt_col_letter}42")
        reward_cell = ws["A42"]
        reward_cell.value = "IF SUCCESS > 90%, REWARD: __________________________________________________"
        reward_cell.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color="007800")
        reward_cell.alignment = Alignment(horizontal="left", vertical="center")

        # --- Rows 43 & 44: Exactly Two-Row Blank Gap ---
        ws.row_dimensions[43].height = 10
        ws.row_dimensions[44].height = 10

        # =========================================================================
        # SLEEP TRACKING SECTION (Starts at Row 45)
        # =========================================================================

        # --- Sleep Tracking Header (Rows 45 & 46) ---
        ws.row_dimensions[45].height = 15
        ws.row_dimensions[46].height = 15

        ws.merge_cells("A45:C46")
        sleep_hdr = ws["A45"]
        sleep_hdr.value = "Sleep Tracking"
        sleep_hdr.font = Font(name=FONT_FAMILY, size=8.5, bold=True)
        sleep_hdr.alignment = Alignment(horizontal="center", vertical="center")
        for r in (45, 46):
            for c in ("A", "B", "C"):
                ws[f"{c}{r}"].border = BORDER_STANDARD

        for d in range(1, days_in_month + 1):
            col = first_day_col + (d - 1)
            c_let = get_column_letter(col)
            is_sunday = (month_weekdays[d - 1] == 'Su')
            header_fill = SUNDAY_HEADER_FILL if is_sunday else WHITE_FILL

            # Row 45: Day number
            cell_d = ws[f"{c_let}45"]
            cell_d.value = d
            cell_d.font = Font(name=FONT_FAMILY, size=8, bold=True)
            cell_d.alignment = Alignment(horizontal="center", vertical="center")
            cell_d.fill = header_fill
            cell_d.border = BORDER_SUNDAY if is_sunday else BORDER_STANDARD

            # Row 46: Weekday
            cell_w = ws[f"{c_let}46"]
            cell_w.value = month_weekdays[d - 1]
            cell_w.font = Font(name=FONT_FAMILY, size=7, bold=is_sunday)
            cell_w.alignment = Alignment(horizontal="center", vertical="center")
            cell_w.fill = header_fill
            cell_w.border = BORDER_SUNDAY if is_sunday else BORDER_STANDARD

        # --- Row 47: Sleep Hours with Dropdown (1 to 24) for each day ---
        sleep_row = 47
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

        # --- Sleep Tracking Line Chart (Row 49 to Row 62) ---
        ws.row_dimensions[48].height = 6

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

        # Data from Row 47 (Sleep Hours)
        data = Reference(ws, min_col=first_day_col, min_row=sleep_row, max_col=last_day_col, max_row=sleep_row)
        # Categories from Row 45 (Day Numbers)
        categories = Reference(ws, min_col=first_day_col, min_row=45, max_col=last_day_col, max_row=45)

        chart.add_data(data, titles_from_data=False, from_rows=True)
        chart.set_categories(categories)

        if chart.series:
            series = chart.series[0]
            series.graphicalProperties.line.solidFill = "3B82F6"
            series.graphicalProperties.line.width = 25000
            series.marker.symbol = "circle"
            series.marker.size = 5
            series.marker.graphicalProperties.solidFill = "1D4ED8"

        ws.add_chart(chart, "A49")

    output_filename = "Habit_Tracker_Sep_Dec_2026.xlsx"
    wb.save(output_filename)
    print(f"Workbook successfully saved to {output_filename}")

if __name__ == "__main__":
    generate_exact_tracker_excel()
