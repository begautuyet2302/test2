import csv
import os
import glob
from datetime import datetime, timedelta
from collections import defaultdict

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

class WorkScheduleGenerator:
    def __init__(self):
        self.staff_data = {}
        self.schedule = {}
        self.staff_work_hours = defaultdict(float)
        
    def parse_staff_availability(self, staff_info):
        """Parse staff availability from CSV file without pandas"""
        csv_file_path = None
        priority_files = ["lich.csv", "schedule.csv", "staff.csv", "data.csv"]
        
        for filename in priority_files:
            if os.path.exists(filename):
                csv_file_path = filename
                break
        
        if csv_file_path is None:
            csv_files = glob.glob("*.csv")
            if csv_files:
                csv_file_path = csv_files[0]
            else:
                raise FileNotFoundError("Không tìm thấy file CSV nào!")
        
        if not os.path.exists(csv_file_path):
            raise FileNotFoundError(f"Không tìm thấy file: {csv_file_path}")
        
        self.staff_data = {}
        
        with open(csv_file_path, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            header = next(reader)
            
            for row in reader:
                if len(row) < 7:
                    continue
                    
                staff_name = row[1].strip()
                if not staff_name:
                    continue
                
                staff_schedule = {'default': {}}
                
                weekday_columns = {
                    2: row[2],  # Thứ 2 (Monday)
                    3: row[3],  # Thứ 3 (Tuesday) 
                    4: row[4],  # Thứ 4 (Wednesday)
                    5: row[5],  # Thứ 5 (Thursday)
                    6: row[6]   # Thứ 6 (Friday)
                }
                
                for weekday_num, shifts_str in weekday_columns.items():
                    if not shifts_str or not shifts_str.strip():
                        continue
                        
                    time_ranges = []
                    shifts = [shift.strip() for shift in str(shifts_str).split(';')]
                    
                    for shift in shifts:
                        if 'Ca 9h - 12h' in shift:
                            time_ranges.append((9, 12))
                        elif 'Ca 13h30 - 16h' in shift:
                            time_ranges.append((13.5, 16))
                        elif 'Ca 16h - 18h30' in shift:
                            time_ranges.append((16, 18.5))
                    
                    if time_ranges:
                        staff_schedule['default'][weekday_num] = time_ranges
                
                if staff_schedule['default']:
                    self.staff_data[staff_name] = staff_schedule
    
    def get_staff_availability(self, staff_name, date, shift_type):
        """Get staff availability for a specific date and shift"""
        staff_info = self.staff_data.get(staff_name, {})
        weekday = date.weekday() + 2
        
        if 'excluded_dates' in staff_info:
            if date.strftime('%Y-%m-%d') in staff_info['excluded_dates']:
                return False, 0
        
        if 'periods' in staff_info:
            for (start_date, end_date), schedule in staff_info['periods'].items():
                if start_date <= date.strftime('%Y-%m-%d') <= end_date:
                    if weekday in schedule:
                        return self._check_shift_overlap(schedule[weekday], shift_type)
            return False, 0
        
        if 'default' in staff_info and weekday in staff_info['default']:
            return self._check_shift_overlap(staff_info['default'][weekday], shift_type)
        
        return False, 0
    
    def _check_shift_overlap(self, time_ranges, shift_type):
        """Check if staff can work in the specified shift and return overlap hours"""
        if shift_type == 'morning':
            shift_start, shift_end = 9, 12
        elif shift_type == 'afternoon1':
            shift_start, shift_end = 13.5, 16
        else:  # afternoon2
            shift_start, shift_end = 16, 18.5
        
        total_overlap = 0
        
        for start, end in time_ranges:
            overlap_start = max(start, shift_start)
            overlap_end = min(end, shift_end)
            if overlap_start < overlap_end:
                total_overlap += overlap_end - overlap_start
        
        return total_overlap > 0, total_overlap
    
    def generate_schedule(self, start_date_str, end_date_str):
        """Generate work schedule with better load balancing"""
        self.parse_staff_availability({})
        
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
        
        initial_schedule = self._generate_initial_schedule(start_date, end_date)
        balanced_schedule = self._balance_workload(initial_schedule, start_date, end_date)
        
        return balanced_schedule
    
    def _generate_initial_schedule(self, start_date, end_date):
        """Generate initial schedule"""
        current_date = start_date
        schedule = []
        
        while current_date <= end_date:
            if current_date.weekday() < 5:
                date_str = current_date.strftime('%Y-%m-%d')
                
                morning_staff = self._select_shift_staff(current_date, 'morning')
                afternoon1_staff = self._select_shift_staff(current_date, 'afternoon1')
                afternoon2_staff = self._select_shift_staff(current_date, 'afternoon2')

                schedule.append({
                    'date': current_date,
                    'Ngày': current_date.strftime('%d/%m/%Y'),
                    'Thứ': f"Thứ {current_date.weekday() + 2}",
                    'Sáng (9h-12h)': morning_staff,
                    'Chiều (13h30-16h)': afternoon1_staff,
                    'Chiều (16h-18h30)': afternoon2_staff
                })
            
            current_date += timedelta(days=1)
        
        return schedule
    
    def _balance_workload(self, schedule, start_date, end_date):
        """Balance workload using reassignment"""
        max_iterations = 3
        
        for iteration in range(max_iterations):
            work_hours = defaultdict(float)
            
            for day in schedule:
                for staff in day['Sáng (9h-12h)']:
                    can_work, hours = self.get_staff_availability(staff, day['date'], 'morning')
                    work_hours[staff] += min(hours, 3)
                
                for staff in day['Chiều (13h30-16h)']:
                    can_work, hours = self.get_staff_availability(staff, day['date'], 'afternoon1')
                    work_hours[staff] += min(hours, 2.5)
                
                for staff in day['Chiều (16h-18h30)']:
                    can_work, hours = self.get_staff_availability(staff, day['date'], 'afternoon2')
                    work_hours[staff] += min(hours, 2.5)
            
            min_hours = min(work_hours.values()) if work_hours else 0
            max_hours = max(work_hours.values()) if work_hours else 0
            
            if max_hours - min_hours <= 8:
                break
            
            self._reassign_shifts(schedule, work_hours)
        
        return schedule
    
    def _reassign_shifts(self, schedule, work_hours):
        """Try to reassign shifts to balance workload"""
        avg_hours = sum(work_hours.values()) / len(work_hours) if work_hours else 0
        
        overworked = [(name, hours) for name, hours in work_hours.items() if hours > avg_hours + 5]
        underworked = [(name, hours) for name, hours in work_hours.items() if hours < avg_hours - 5]
        
        if not overworked or not underworked:
            return
        
        for day in schedule:
            for shift_type in ['morning', 'afternoon1', 'afternoon2']:
                if shift_type == 'morning':
                    shift_key = 'Sáng (9h-12h)'
                elif shift_type == 'afternoon1':
                    shift_key = 'Chiều (13h30-16h)'
                else:
                    shift_key = 'Chiều (16h-18h30)'
                current_staff = day[shift_key]
                
                for i, staff_name in enumerate(current_staff):
                    if any(name == staff_name for name, _ in overworked):
                        for replacement_name, _ in underworked:
                            can_work, _ = self.get_staff_availability(replacement_name, day['date'], shift_type)
                            if can_work and replacement_name not in current_staff:
                                current_staff[i] = replacement_name
                                break
    
    def _select_shift_staff(self, date, shift_type):
        """Select 3 staff members for a shift, balancing workload fairly"""
        available_staff = []
        
        for staff_name in self.staff_data.keys():
            can_work, hours = self.get_staff_availability(staff_name, date, shift_type)
            if can_work:
                available_staff.append((staff_name, hours))
        
        if len(available_staff) <= 3:
            selected_staff = [staff[0] for staff in available_staff]
        else:
            def selection_score(staff_tuple):
                name, available_hours = staff_tuple
                current_hours = self.staff_work_hours[name]
                
                balance_score = -current_hours * 0.8
                efficiency_score = available_hours * 0.2
                
                return balance_score + efficiency_score
            
            available_staff.sort(key=selection_score, reverse=True)
            selected_staff = [staff[0] for staff in available_staff[:3]]
        
        if shift_type == 'morning':
            shift_hours = 3
        elif shift_type == 'afternoon1':
            shift_hours = 2.5
        else:
            shift_hours = 2.5
        
        for staff_name in selected_staff:
            staff_hours = next((hours for name, hours in available_staff if name == staff_name), shift_hours)
            actual_hours = min(staff_hours, shift_hours)
            self.staff_work_hours[staff_name] += actual_hours
        
        return selected_staff
    
    def save_to_excel(self, schedule, filename):
        """Save schedule to Excel file with formatting"""
        if not EXCEL_AVAILABLE:
            # Fallback to CSV if openpyxl not available
            csv_filename = filename.replace('.xlsx', '.csv')
            self.save_to_csv(schedule, csv_filename)
            return
            
        wb = Workbook()
        ws = wb.active
        ws.title = "Work Schedule"
        
        # Define styles
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        center_alignment = Alignment(horizontal="center", vertical="center")
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Header
        headers = ['Date', 'Thứ', 'TYPE', 'Lab01', 'Lab02', 'Lab03']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # Data rows
        row_num = 2
        for day_info in schedule:
            date = day_info['date']
            date_str = date.strftime('%d/%m/%Y')
            weekday_str = f"Thứ {day_info['Thứ'].split()[1]}"
            
            # Morning shift
            morning_staff = day_info['Sáng (9h-12h)']
            morning_data = [
                date_str,
                weekday_str,
                'Sáng (09:00~12:00)',
                morning_staff[0] if len(morning_staff) > 0 else '',
                morning_staff[1] if len(morning_staff) > 1 else '',
                morning_staff[2] if len(morning_staff) > 2 else ''
            ]
            
            for col, value in enumerate(morning_data, 1):
                cell = ws.cell(row=row_num, column=col, value=value)
                cell.alignment = center_alignment
                cell.border = thin_border
            row_num += 1
            
            # Afternoon shift 1
            afternoon1_staff = day_info['Chiều (13h30-16h)']
            afternoon1_data = [
                date_str,
                weekday_str,
                'Chiều (13:30 ~ 16:00)',
                afternoon1_staff[0] if len(afternoon1_staff) > 0 else '',
                afternoon1_staff[1] if len(afternoon1_staff) > 1 else '',
                afternoon1_staff[2] if len(afternoon1_staff) > 2 else ''
            ]
            
            for col, value in enumerate(afternoon1_data, 1):
                cell = ws.cell(row=row_num, column=col, value=value)
                cell.alignment = center_alignment
                cell.border = thin_border
            row_num += 1
            
            # Afternoon shift 2
            afternoon2_staff = day_info['Chiều (16h-18h30)']
            afternoon2_data = [
                date_str,
                weekday_str,
                'Chiều (16:00 ~ 18:30)',
                afternoon2_staff[0] if len(afternoon2_staff) > 0 else '',
                afternoon2_staff[1] if len(afternoon2_staff) > 1 else '',
                afternoon2_staff[2] if len(afternoon2_staff) > 2 else ''
            ]
            
            for col, value in enumerate(afternoon2_data, 1):
                cell = ws.cell(row=row_num, column=col, value=value)
                cell.alignment = center_alignment
                cell.border = thin_border
            row_num += 1
        
        # Auto-adjust column widths
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 30)
            ws.column_dimensions[column].width = adjusted_width
        
        wb.save(filename)
    
    def save_to_csv(self, schedule, filename):
        """Save schedule to CSV file as fallback"""
        csv_data = []
        csv_data.append(['Date', 'Thứ', 'TYPE', 'Lab01', 'Lab02', 'Lab03'])
        
        for day_info in schedule:
            date = day_info['date']
            date_str = date.strftime('%d/%m/%Y')
            weekday_str = f"Thứ {day_info['Thứ'].split()[1]}"
            
            morning_staff = day_info['Sáng (9h-12h)']
            morning_row = [
                date_str,
                weekday_str,
                'Sáng (09:00~12:00)',
                morning_staff[0] if len(morning_staff) > 0 else '',
                morning_staff[1] if len(morning_staff) > 1 else '',
                morning_staff[2] if len(morning_staff) > 2 else ''
            ]
            csv_data.append(morning_row)
            
            afternoon1_staff = day_info['Chiều (13h30-16h)']
            afternoon1_row = [
                date_str,
                weekday_str,
                'Chiều (13:30 ~ 16:00)',
                afternoon1_staff[0] if len(afternoon1_staff) > 0 else '',
                afternoon1_staff[1] if len(afternoon1_staff) > 1 else '',
                afternoon1_staff[2] if len(afternoon1_staff) > 2 else ''
            ]
            csv_data.append(afternoon1_row)
            
            afternoon2_staff = day_info['Chiều (16h-18h30)']
            afternoon2_row = [
                date_str,
                weekday_str,
                'Chiều (16:00 ~ 18:30)',
                afternoon2_staff[0] if len(afternoon2_staff) > 0 else '',
                afternoon2_staff[1] if len(afternoon2_staff) > 1 else '',
                afternoon2_staff[2] if len(afternoon2_staff) > 2 else ''
            ]
            csv_data.append(afternoon2_row)
        
        with open(filename, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(csv_data)

def get_auto_date_range():
    """Get automatic date range for schedule"""
    today = datetime.now()
    weekday = today.weekday()
    
    if weekday >= 5:
        days_until_next_monday = (7 - weekday) + 0
        start_date = today + timedelta(days=days_until_next_monday)
        end_date = start_date + timedelta(days=4)
    else:
        start_date = today + timedelta(days=1)
        days_until_friday = 4 - weekday
        if days_until_friday <= 0: 
            days_until_friday += 7
        end_date = today + timedelta(days=days_until_friday)
    
    return start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')

def main():
    """Main function"""
    start_date, end_date = get_auto_date_range()
    scheduler = WorkScheduleGenerator()
    schedule = scheduler.generate_schedule(start_date, end_date)
    scheduler.save_to_excel(schedule, 'work_schedule.xlsx')

if __name__ == "__main__":
    main()