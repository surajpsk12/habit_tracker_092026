import calendar
import random
from fpdf import FPDF

class HabitTracker(FPDF):
    def header(self):
        pass

    def footer(self):
        # Position at 22 mm from bottom
        self.set_y(-22)
        self.set_font('Arial', 'B', 9)
        self.set_text_color(0, 0, 0)
        
        # Heading for Targets
        self.cell(0, 5, "MONTHLY TARGETS", 0, 1, 'L')
        
        start_y = self.get_y()
        col1_x = 10
        col2_x = 90  # Adjusted spacing
        
        self.set_font('Arial', '', 8)
        # Column 1 (Targets 1-4)
        for i in range(1, 5):
            self.set_xy(col1_x, start_y + (i-1) * 4)
            self.cell(70, 4, f"{i}. ____________________ : ____________________", 0, 0, 'L')
            
        # Column 2 (Targets 5-8)
        for i in range(5, 9):
            self.set_xy(col2_x, start_y + (i-5) * 4)
            self.cell(70, 4, f"{i}. ____________________ : ____________________", 0, 0, 'L')

def draw_protocol_header(pdf_obj, col_width_sn, col_width_label, col_width_target, col_width_day, col_width_success, row_height, days, weekdays):
    start_x = pdf_obj.get_x()
    start_y = pdf_obj.get_y()
    merged_height = row_height * 2
    
    pdf_obj.set_font('Arial', 'B', 8)
    pdf_obj.set_x(start_x + col_width_sn + col_width_label + col_width_target)
    for day in range(1, days + 1):
         if weekdays[day-1] == 'Su':
             pdf_obj.set_fill_color(210, 210, 210)
             pdf_obj.cell(col_width_day, row_height, str(day), 1, 0, 'C', True)
         else:
             pdf_obj.cell(col_width_day, row_height, str(day), 1, 0, 'C')
    
    pdf_obj.set_font('Arial', 'B', 7)
    pdf_obj.cell(col_width_success, row_height, "", 1, 0, 'C')
    pdf_obj.ln()
    
    pdf_obj.set_font('Arial', '', 7)
    pdf_obj.set_x(start_x + col_width_sn + col_width_label + col_width_target)
    for i, day_char in enumerate(weekdays):
        if day_char == 'Su':
            pdf_obj.set_fill_color(210, 210, 210)
            pdf_obj.cell(col_width_day, row_height, day_char, 1, 0, 'C', True)
        else:
            pdf_obj.cell(col_width_day, row_height, day_char, 1, 0, 'C')
    
    pdf_obj.cell(col_width_success, row_height, "", 1, 0, 'C')
    pdf_obj.ln()
    
    y_after_headers = pdf_obj.get_y()
    pdf_obj.set_font('Arial', 'B', 8)
    font_size_mm = 8 * 0.352778 
    text_y_pos = start_y + (merged_height / 2) - (font_size_mm / 2)

    pdf_obj.rect(start_x, start_y, col_width_sn, merged_height)
    pdf_obj.set_xy(start_x, text_y_pos)
    pdf_obj.cell(col_width_sn, 0, "S.No.", 0, 0, 'C') 

    pdf_obj.rect(start_x + col_width_sn, start_y, col_width_label, merged_height)
    pdf_obj.set_xy(start_x + col_width_sn, text_y_pos)
    pdf_obj.cell(col_width_label, 0, "Protocols", 0, 0, 'C')

    pdf_obj.rect(start_x + col_width_sn + col_width_label, start_y, col_width_target, merged_height)
    pdf_obj.set_xy(start_x + col_width_sn + col_width_label, text_y_pos)
    pdf_obj.cell(col_width_target, 0, "Target", 0, 0, 'C')

    pdf_obj.rect(start_x + col_width_sn + col_width_label + col_width_target + (col_width_day * days), start_y, col_width_success, merged_height)
    pdf_obj.set_xy(start_x + col_width_sn + col_width_label + col_width_target + (col_width_day * days), text_y_pos)
    pdf_obj.cell(col_width_success, 0, "Success %", 0, 0, 'C')

    pdf_obj.set_y(y_after_headers)

def draw_sleep_header(pdf_obj, col_width_label, col_width_day, row_height, days, weekdays):
    start_x = pdf_obj.get_x()
    start_y = pdf_obj.get_y()
    merged_height = row_height * 2
    pdf_obj.set_font('Arial', 'B', 8)
    pdf_obj.set_x(start_x + col_width_label) 
    for day in range(1, days + 1):
         if weekdays[day-1] == 'Su':
             pdf_obj.set_fill_color(210, 210, 210)
             pdf_obj.cell(col_width_day, row_height, str(day), 1, 0, 'C', True)
         else:
             pdf_obj.cell(col_width_day, row_height, str(day), 1, 0, 'C')
    pdf_obj.ln()
    pdf_obj.set_font('Arial', '', 7)
    pdf_obj.set_x(start_x + col_width_label) 
    for i, day_char in enumerate(weekdays):
        if day_char == 'Su':
            pdf_obj.set_fill_color(210, 210, 210)
            pdf_obj.cell(col_width_day, row_height, day_char, 1, 0, 'C', True)
        else:
            pdf_obj.cell(col_width_day, row_height, day_char, 1, 0, 'C')
    pdf_obj.ln()
    y_after_headers = pdf_obj.get_y()
    pdf_obj.set_font('Arial', 'B', 8)
    pdf_obj.rect(start_x, start_y, col_width_label, merged_height)
    pdf_obj.set_xy(start_x, start_y + (merged_height/2) - 1.5)
    pdf_obj.cell(col_width_label, 0, "Sleep Tracking", 0, 0, 'C')
    pdf_obj.set_y(y_after_headers)

def create_tracker():
    try:
        year = int(input("Enter Year (e.g. 2026): "))
        month = int(input("Enter Month (1-12): "))
    except ValueError:
        return

    month_name = calendar.month_name[month]
    days_in_month = calendar.monthrange(year, month)[1]
    first_weekday = calendar.monthrange(year, month)[0]
    weekdays_list = ['M', 'T', 'W', 'Th', 'F', 'Sa', 'Su']
    month_weekdays = [weekdays_list[(first_weekday + d) % 7] for d in range(days_in_month)]

    quotes = ["Discipline equals freedom.", "Win the morning, win the day.", "Focus on the process.", "Stay consistent."]
    random_quote = random.choice(quotes)

    group_colors = [(255,235,235), (235,245,255), (235,255,235), (255,255,235), (245,235,255)]

    pdf = HabitTracker(orientation='L', unit='mm', format='A4')
    pdf.add_page()

    col_width_day = 6.4; col_width_sn = 10; col_width_label = 38; col_width_target = 14; col_width_success = 18 
    row_height = 4.8; total_row_height = 7 
    
    # --- Header Section ---
    pdf.set_font('Arial', 'B', 20)
    pdf.cell(100, 8, 'Habit Tracker', 0, 0, 'L')
    
    # Name at Top Right Extreme
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(0, 8, '@surajvansh', 0, 1, 'R')
    
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 5, f"{month_name} {year}", 0, 1, 'L')
    
    # Quote centered
    pdf.set_xy(60, 10); pdf.set_font('Arial', 'B', 11); pdf.cell(170, 8, f'"{random_quote}"', 0, 0, 'C')
    
    # Dashboard line
    pdf.set_xy(pdf.l_margin, 15); pdf.set_font('Arial', 'B', 10); pdf.cell(0, 10, f"{days_in_month} Days Dashboard", 0, 1, 'C')

    pdf.set_y(26)
    draw_protocol_header(pdf, col_width_sn, col_width_label, col_width_target, col_width_day, col_width_success, row_height, days_in_month, month_weekdays)
    
    # Protocol Rows
    pdf.set_font('Arial', '', 8)
    for i in range(1, 16):
        group_idx = (i - 1) // 3
        r, g, b = group_colors[group_idx]
        pdf.set_fill_color(r, g, b)
        pdf.cell(col_width_sn, row_height, str(i), 1, 0, 'C', True) 
        pdf.cell(col_width_label, row_height, "", 1, 0, 'L', True) 
        pdf.cell(col_width_target, row_height, "", 1, 0, 'C', True) 
        for d in range(days_in_month):
            if month_weekdays[d] == 'Su':
                pdf.set_line_width(0.5); pdf.set_fill_color(max(0,r-25), max(0,g-25), max(0,b-25))
                pdf.cell(col_width_day, row_height, "", 1, 0, 'C', True)
                pdf.set_fill_color(r, g, b); pdf.set_line_width(0.2)
            else:
                pdf.cell(col_width_day, row_height, "", 1, 0, 'C', True)
        pdf.cell(col_width_success, row_height, "", 1, 1, 'C', True)
        
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(col_width_sn + col_width_label + col_width_target, total_row_height, "DAILY TOTAL SCORE", 1, 0, 'R')
    for _ in range(days_in_month):
        pdf.cell(col_width_day, total_row_height, "", 1, 0, 'C')
    pdf.cell(col_width_success, total_row_height, "", 1, 1, 'C')

    # Stakes
    pdf.ln(1)
    pdf.set_font('Arial', 'B', 9)
    pdf.set_text_color(200, 0, 0); pdf.cell(0, 4, "IF SUCCESS < 80%, PUNISHMENT: ________________________________________________", 0, 1, 'L')
    pdf.set_text_color(0, 120, 0); pdf.cell(0, 4, "IF SUCCESS > 90%, REWARD: __________________________________________________", 0, 1, 'L')
    pdf.set_text_color(0, 0, 0)

    # Sleep Tracking
    sleep_label_width = col_width_sn + col_width_label + col_width_target
    draw_sleep_header(pdf, sleep_label_width, col_width_day, row_height, days_in_month, month_weekdays)
    for label in ["9hrs", "8hrs", "7hrs", "6hrs", "5hrs"]:
        pdf.cell(sleep_label_width, row_height, label, 1, 0, 'C')
        for d in range(days_in_month):
             pdf.cell(col_width_day, row_height, "", 1, 0, 'C', (month_weekdays[d]=='Su'))
        pdf.ln()

    # Notes Section
    pdf.ln(1)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(0, 5, "NOTES:", 0, 1, 'L')
    pdf.set_fill_color(255, 255, 230) 
    note_width = sleep_label_width + (col_width_day * days_in_month)
    pdf.rect(pdf.get_x(), pdf.get_y(), note_width, 14, 'F')
    
    pdf.output(f"{month_name.lower()}{year}Habit_Tracker.pdf")
    print("PDF Generated successfully.")

if __name__ == "__main__":
    create_tracker()