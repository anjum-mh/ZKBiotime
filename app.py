import os
import sqlite3
import subprocess
import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="ZKBioTime Attendance Suite & Fine Management", page_icon="📊", layout="wide")

DB_FILE = "attendance_history.db"

# --- HIERARCHY MAPPING ---
PREFIX_MAP = {
    'MTCM': ('Industrial Machinery and Maintenance Technologies Department', 'Mechanical Technology CMT'),
    'HVAC': ('Industrial Machinery and Maintenance Technologies Department', 'HVACR'),
    'HVACR': ('Industrial Machinery and Maintenance Technologies Department', 'HVACR'),
    'AUTO': ('Industrial Machinery and Maintenance Technologies Department', 'Automotive'),
    'AD': ('Industrial Machinery and Maintenance Technologies Department', 'Automotive'),
    'MECH': ('Industrial Machinery and Maintenance Technologies Department', 'Mechanical Technology'),
    'CIV': ('Construction and Infrastructure Technologies Department', 'Civil Technology'),
    'QS': ('Construction and Infrastructure Technologies Department', 'QS Technology'),
    'ARCH': ('Construction and Infrastructure Technologies Department', 'Architecture Technology'),
    'ICT': ('Computing and Digital Technologies Department', 'Information and Communication Technology'),
    'ELEC': ('Computing and Digital Technologies Department', 'Electrical Technology'),
    'AI': ('Computing and Digital Technologies Department', 'Artificial Intelligence')
}

PARENT_ORDER = [
    'Industrial Machinery and Maintenance Technologies Department',
    'Construction and Infrastructure Technologies Department',
    'Computing and Digital Technologies Department'
]

# --- DATABASE SETUP & MIGRATION ---
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS attendance_logs (
            emp_id TEXT,
            name TEXT,
            record_date TEXT,
            shift_name TEXT,
            parent_dept TEXT,
            sub_dept TEXT,
            year TEXT,
            section TEXT,
            dept_code TEXT,
            clock_in TEXT,
            is_present INTEGER,
            is_leave INTEGER,
            is_absent INTEGER,
            is_late INTEGER,
            excuse_reason TEXT,
            PRIMARY KEY (emp_id, record_date, shift_name)
        )
    ''')
    
    try:
        c.execute("ALTER TABLE attendance_logs ADD COLUMN excuse_reason TEXT")
    except sqlite3.OperationalError:
        pass 
        
    conn.commit()
    conn.close()

init_db()

def save_records_to_db(df):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    for _, r in df.iterrows():
        c.execute('''
            INSERT OR REPLACE INTO attendance_logs
            (emp_id, name, record_date, shift_name, parent_dept, sub_dept, year, section, dept_code, clock_in, is_present, is_leave, is_absent, is_late, excuse_reason)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, COALESCE((SELECT excuse_reason FROM attendance_logs WHERE emp_id=? AND record_date=? AND shift_name=?), ''))
        ''', (
            str(r['EmpID']), str(r['Name']), str(r['Date']), str(r['ShiftName']),
            str(r['ParentDept']), str(r['SubDept']), str(r['Year']), str(r['Section']),
            str(r['DeptCode']), str(r['ClockIn']), int(r['IsPresent']), int(r['IsLeave']),
            int(r['IsAbsent']), int(r['IsLate']), str(r['EmpID']), str(r['Date']), str(r['ShiftName'])
        ))
    conn.commit()
    conn.close()

def convert_absence_to_leave(emp_id, dates_list, shift_name, reason):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    for d_str in dates_list:
        c.execute('''
            UPDATE attendance_logs
            SET is_absent = 0, is_leave = 1, excuse_reason = ?
            WHERE emp_id = ? AND record_date = ? AND shift_name = ?
        ''', (reason, str(emp_id), str(d_str), str(shift_name)))
    conn.commit()
    conn.close()

def parse_dept_header(dept_str):
    clean_str = str(dept_str).replace('Department:', '').strip().upper()
    parts = clean_str.split('-')
    prefix = parts[0].strip() if len(parts) > 0 else 'OTHER'
    year_digit = parts[1].strip() if len(parts) > 1 else '1'
    section_letter = parts[2].strip() if len(parts) > 2 else 'A'

    year_map = {'1': '1st Year', '2': '2nd Year', '3': '3rd Year'}
    year_str = year_map.get(year_digit, f"{year_digit}th Year")
    sec_str = f"Section {section_letter}"

    parent_dept, sub_dept = PREFIX_MAP.get(
        prefix,
        ('Industrial Machinery and Maintenance Technologies Department', 'Other / General')
    )
    return parent_dept, sub_dept, year_str, sec_str, clean_str

def detect_shift_and_threshold(df_records):
    valid_times = [r['ClockIn'] for r in df_records if r['ClockIn']]
    timetable_hints = [r['Timetable'].lower() for r in df_records if r.get('Timetable')]
    
    is_friday = False
    for r in df_records:
        if r.get('Date'):
            try:
                dt = datetime.strptime(r['Date'][:10], '%Y-%m-%d')
                if dt.weekday() == 4:
                    is_friday = True
                    break
            except ValueError:
                pass

    shift_name = 'First Shift'
    if any('evening' in t or '2nd' in t or 'second' in t for t in timetable_hints):
        shift_name = 'Second Shift'
    elif valid_times:
        avg_hour = sum(int(t.split(':')[0]) for t in valid_times if ':' in t) / max(len(valid_times), 1)
        if avg_hour >= 11:
            shift_name = 'Second Shift'

    late_threshold = ('14:05' if is_friday else '12:55') if shift_name == 'Second Shift' else '08:05'
    return shift_name, late_threshold, is_friday

def parse_biotime_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    df_raw = pd.read_excel(uploaded_file, sheet_name=xls.sheet_names[0])

    current_dept_raw = None
    raw_records = []
    col_idx_map = {'id': 0, 'name': 1, 'date': 2, 'timetable': None, 'clockin': 3}
    skip_keywords = {'employee id', 'nan', 'none', '', 'id', 'statistics', 'total', 'summary', 'count', 'average'}

    for _, row in df_raw.iterrows():
        row_vals = [str(val).strip() if pd.notna(val) else '' for val in row]
        row_text = " ".join(row_vals).lower()
        val0 = row_vals[0] if len(row_vals) > 0 else ''

        if 'department:' in row_text:
            for val in row_vals:
                if 'department:' in val.lower():
                    current_dept_raw = val
                    break
            continue

        if 'employee id' in row_text or 'clock in' in row_text:
            for idx, val in enumerate(row_vals):
                v_lower = val.lower()
                if 'employee id' in v_lower or 'emp id' in v_lower:
                    col_idx_map['id'] = idx
                elif 'name' in v_lower:
                    col_idx_map['name'] = idx
                elif 'date' in v_lower:
                    col_idx_map['date'] = idx
                elif 'timetable' in v_lower:
                    col_idx_map['timetable'] = idx
                elif 'clock' in v_lower or 'in' in v_lower:
                    col_idx_map['clockin'] = idx
            continue

        if not current_dept_raw or not val0 or val0.lower() in skip_keywords:
            continue
        if any(kw in row_text for kw in ['statistics', 'summary', 'total record', 'page ', 'printed on']):
            continue

        emp_id = row_vals[col_idx_map['id']] if col_idx_map['id'] < len(row_vals) else val0
        name = row_vals[col_idx_map['name']] if col_idx_map['name'] and col_idx_map['name'] < len(row_vals) else ''
        report_date = row_vals[col_idx_map['date']][:10] if col_idx_map['date'] and col_idx_map['date'] < len(row_vals) else datetime.now().strftime('%Y-%m-%d')
        timetable = row_vals[col_idx_map['timetable']] if col_idx_map['timetable'] and col_idx_map['timetable'] < len(row_vals) else ''

        clock_in = ''
        c_idx = col_idx_map['clockin']
        if c_idx is not None and c_idx < len(row_vals):
            c_val = row_vals[c_idx]
            if c_val.lower() not in ['nan', 'none', '']:
                clock_in = c_val[:5] if len(c_val) >= 5 else c_val

        parent, sub_dept, year, section, dept_code = parse_dept_header(current_dept_raw)
        raw_records.append({
            'EmpID': emp_id, 'Name': name, 'Date': report_date, 'Timetable': timetable,
            'ParentDept': parent, 'SubDept': sub_dept, 'Year': year, 'Section': section,
            'DeptCode': dept_code, 'ClockIn': clock_in
        })

    if not raw_records:
        return pd.DataFrame(), "First Shift", "08:05"

    shift_name, late_threshold, is_friday = detect_shift_and_threshold(raw_records)
    processed_records = []
    full_shift_title = f"{shift_name}{' (Friday)' if is_friday else ''}"
    for r in raw_records:
        clock_in = r['ClockIn']
        r['ShiftName'] = full_shift_title
        r['LateThreshold'] = late_threshold
        r['IsPresent'] = 1 if clock_in else 0
        r['IsLeave'] = 0
        r['IsAbsent'] = 1 if not clock_in else 0
        r['IsLate'] = 1 if (clock_in and clock_in > late_threshold) else 0
        processed_records.append(r)

    return pd.DataFrame(processed_records), full_shift_title, late_threshold

def apply_leave_mapping(df, leave_mapping):
    df = df.copy()
    df['IsLeave'] = 0
    df['IsAbsent'] = 1 - df['IsPresent']
    
    for code, l_val in leave_mapping.items():
        if l_val > 0:
            absent_idx = df[(df['DeptCode'] == code) & (df['IsAbsent'] == 1)].index.tolist()
            leave_targets = absent_idx[:l_val]
            if leave_targets:
                df.loc[leave_targets, 'IsLeave'] = 1
                df.loc[leave_targets, 'IsAbsent'] = 0
    return df

def build_html_report(df, leave_mapping, shift_title, report_date):
    df = apply_leave_mapping(df, leave_mapping)

    html = f"""
    <html>
    <head>
        <style>
            @page {{ size: A4 portrait; margin: 6mm; }}
            body {{ font-family: Calibri, sans-serif; color: black; background: white; margin: 0; padding: 0; }}
            h2 {{ text-align: center; font-size: 16pt; margin: 0 0 2px 0; color: #1E293B; }}
            .date-shift {{ text-align: center; font-size: 16pt; font-weight: bold; margin: 0 0 8px 0; color: #334155; }}
            table {{ width: 100%; border-collapse: collapse; page-break-inside: avoid; }}
            th, td {{ border: 1px solid #94A3B8; padding: 3px 6px; text-align: center; white-space: nowrap; }}
            th {{ background-color: #1E293B !important; color: white !important; font-size: 14pt !important; font-weight: bold !important; -webkit-print-color-adjust: exact; }}
            .parent-header {{ background-color: #E2E8F0 !important; font-size: 14pt !important; font-weight: bold !important; text-align: left !important; color: #1E293B !important; -webkit-print-color-adjust: exact; }}
            .subtotal-row {{ background-color: #F1F5F9 !important; font-size: 14pt !important; font-weight: bold !important; -webkit-print-color-adjust: exact; }}
            .grand-total {{ background-color: #1E293B !important; color: white !important; font-size: 14pt !important; font-weight: bold !important; -webkit-print-color-adjust: exact; }}
            .regular-row {{ font-size: 12pt !important; }}
            td:first-child {{ text-align: left; }}
        </style>
    </head>
    <body>
        <h2>Daily Attendance Report</h2>
        <div class="date-shift">Shift: {shift_title} &nbsp;|&nbsp; Date: {report_date}</div>
        <table>
            <thead>
                <tr>
                    <th>Department / Section</th>
                    <th>Total Strength</th>
                    <th>Present</th>
                    <th>Absent</th>
                    <th>Leave</th>
                    <th>Late</th>
                </tr>
            </thead>
            <tbody>
    """

    grand_tot, grand_pres, grand_abs, grand_leave, grand_late = 0, 0, 0, 0, 0

    for parent in PARENT_ORDER:
        p_data = df[df['ParentDept'] == parent]
        if p_data.empty:
            continue
        
        html += f'<tr><td colspan="6" class="parent-header">{parent}</td></tr>'
        p_tot, p_pres, p_abs, p_leave, p_late = 0, 0, 0, 0, 0

        for dept_code, sec_group in p_data.groupby('DeptCode'):
            tot = len(sec_group)
            pres = sec_group['IsPresent'].sum()
            leave = sec_group['IsLeave'].sum()
            absent = max(0, tot - pres - leave)
            late = sec_group['IsLate'].sum()

            p_tot += tot
            p_pres += pres
            p_abs += absent
            p_leave += leave
            p_late += late

            html += f"""
            <tr class="regular-row">
                <td>{dept_code}</td>
                <td>{tot}</td>
                <td>{pres}</td>
                <td>{absent}</td>
                <td>{leave}</td>
                <td>{late}</td>
            </tr>
            """

        grand_tot += p_tot
        grand_pres += p_pres
        grand_abs += p_abs
        grand_leave += p_leave
        grand_late += p_late

        html += f"""
        <tr class="subtotal-row">
            <td>Subtotal - {parent}</td>
            <td>{p_tot}</td>
            <td>{p_pres}</td>
            <td>{p_abs}</td>
            <td>{p_leave}</td>
            <td>{p_late}</td>
        </tr>
        """

    html += f"""
        <tr class="grand-total">
            <td>GRAND TOTAL</td>
            <td>{grand_tot}</td>
            <td>{grand_pres}</td>
            <td>{grand_abs}</td>
            <td>{grand_leave}</td>
            <td>{grand_late}</td>
        </tr>
    """

    html += "</tbody></table></body></html>"
    return html

def generate_interactive_master_workbook(df, output_path):
    wb = openpyxl.Workbook()
    default_sheet = wb.active

    primary_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    font_title = Font(name="Calibri", size=14, bold=True, color="1E293B")
    font_hdr = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_reg = Font(name="Calibri", size=11)
    font_grand = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    
    btn_blue_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    btn_blue_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    btn_blue_border = Border(left=Side(style='medium', color='1E40AF'), right=Side(style='medium', color='1E40AF'), top=Side(style='medium', color='1E40AF'), bottom=Side(style='medium', color='1E40AF'))

    btn_dark_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    btn_dark_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    btn_dark_border = Border(left=Side(style='medium', color='0F172A'), right=Side(style='medium', color='0F172A'), top=Side(style='medium', color='0F172A'), bottom=Side(style='medium', color='0F172A'))

    btn_amber_fill = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")
    btn_amber_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    btn_amber_border = Border(left=Side(style='medium', color='B45309'), right=Side(style='medium', color='B45309'), top=Side(style='medium', color='B45309'), bottom=Side(style='medium', color='B45309'))

    thin_side = Side(border_style="thin", color="CBD5E1")
    double_bottom = Side(border_style="double", color="1E293B")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_grand = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_bottom)

    used_sheet_names = set()

    def make_unique_sheet_name(raw_name, prefix=""):
        clean_name = raw_name.strip().replace(' ', '_').replace('/', '_').replace('-', '_')
        candidate = f"{prefix}{clean_name}"[:28]
        final_name = candidate
        counter = 1
        while final_name.lower() in [s.lower() for s in used_sheet_names]:
            final_name = f"{candidate[:24]}_{counter}"
            counter += 1
        used_sheet_names.add(final_name)
        return final_name

    parent_sheets = {}
    sub_dept_sheets = {}
    section_sheets = {}
    absent_sheets = {}
    late_sheets = {}

    for parent in PARENT_ORDER:
        p_data = df[df['ParentDept'] == parent]
        if p_data.empty:
            continue
        parent_sheets[parent] = make_unique_sheet_name(parent[:15], "Sub_")

        for sub_dept, sub_group in p_data.groupby('SubDept'):
            sub_dept_sheets[sub_dept] = make_unique_sheet_name(sub_dept[:15], "Sec_")

            for dept_code, sec_group in sub_group.groupby('DeptCode'):
                section_sheets[dept_code] = make_unique_sheet_name(dept_code[:15], "D_")
                absent_sheets[dept_code] = make_unique_sheet_name(dept_code[:15], "Abs_")
                late_sheets[dept_code] = make_unique_sheet_name(dept_code[:15], "Late_")

    ws_master = wb.create_sheet(title="Master Summary")
    ws_master.views.sheetView[0].showGridLines = True
    ws_master.merge_cells("A1:G1")
    ws_master["A1"] = "HIERARCHICAL ATTENDANCE SUMMARY (MASTER)"
    ws_master["A1"].font = font_title
    ws_master["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_master.row_dimensions[1].height = 25

    headers = ["Parent Department", "Total Strength", "Total Present", "Total Absent", "Total Leave", "Total Late", "Action Button"]
    for c_idx, h in enumerate(headers, 1):
        cell = ws_master.cell(row=2, column=c_idx, value=h)
        cell.font = font_hdr
        cell.fill = primary_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_cell
    ws_master.row_dimensions[2].height = 22

    r_idx = 3
    master_start_row = 3
    for parent in PARENT_ORDER:
        p_data = df[df['ParentDept'] == parent]
        if p_data.empty:
            continue
        tot = len(p_data)
        pres = int(p_data['IsPresent'].sum())
        leave = int(p_data['IsLeave'].sum())
        absent = max(0, tot - pres - leave)
        late = int(p_data['IsLate'].sum())

        sub_sheet_name = parent_sheets[parent]
        
        ws_master.cell(row=r_idx, column=1, value=parent).font = font_reg
        ws_master.cell(row=r_idx, column=2, value=tot).font = font_reg
        ws_master.cell(row=r_idx, column=3, value=pres).font = font_reg
        ws_master.cell(row=r_idx, column=4, value=absent).font = font_reg
        ws_master.cell(row=r_idx, column=5, value=leave).font = font_reg
        ws_master.cell(row=r_idx, column=6, value=late).font = font_reg
        
        action_cell = ws_master.cell(row=r_idx, column=7, value=" VIEW CHILD DEPTS ")
        action_cell.font = btn_blue_font
        action_cell.fill = btn_blue_fill
        action_cell.border = btn_blue_border
        action_cell.alignment = Alignment(horizontal="center", vertical="center")
        action_cell.hyperlink = f"#'{sub_sheet_name}'!A1"

        for c in range(1, 7):
            c_cell = ws_master.cell(row=r_idx, column=c)
            c_cell.border = border_cell
            if c > 1:
                c_cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws_master.row_dimensions[r_idx].height = 24
        r_idx += 1

    master_end_row = r_idx - 1

    ws_master.cell(row=r_idx, column=1, value="GRAND TOTAL").font = font_grand
    ws_master.cell(row=r_idx, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for c_idx, col_letter in enumerate(['B', 'C', 'D', 'E', 'F'], 2):
        cell = ws_master.cell(row=r_idx, column=c_idx, value=f"=SUM({col_letter}{master_start_row}:{col_letter}{master_end_row})")
        cell.font = font_grand
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_master.cell(row=r_idx, column=7, value="").fill = primary_fill

    for c in range(1, 8):
        c_cell = ws_master.cell(row=r_idx, column=c)
        c_cell.border = border_grand
        if c < 7:
            c_cell.fill = primary_fill
    ws_master.row_dimensions[r_idx].height = 25

    ws_master.column_dimensions['A'].width = 45
    for col in ['B', 'C', 'D', 'E', 'F']:
        ws_master.column_dimensions[col].width = 18
    ws_master.column_dimensions['G'].width = 26

    for parent in PARENT_ORDER:
        p_data = df[df['ParentDept'] == parent]
        if p_data.empty:
            continue
        sub_sheet_name = parent_sheets[parent]
        ws_sub = wb.create_sheet(title=sub_sheet_name)
        ws_sub.views.sheetView[0].showGridLines = True

        ws_sub.merge_cells("A1:G1")
        ws_sub["A1"] = f"{parent.upper()} - SUB DEPARTMENTS"
        ws_sub["A1"].font = font_title
        ws_sub["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws_sub.row_dimensions[1].height = 25

        back_cell = ws_sub.cell(row=2, column=1, value=" ◀ BACK TO MASTER ")
        back_cell.font = btn_dark_font
        back_cell.fill = btn_dark_fill
        back_cell.border = btn_dark_border
        back_cell.alignment = Alignment(horizontal="center", vertical="center")
        back_cell.hyperlink = "#'Master Summary'!A1"

        sub_headers = ["Child / Sub-Department", "Total Strength", "Total Present", "Total Absent", "Total Leave", "Total Late", "Action Button"]
        for c_idx, h in enumerate(sub_headers, 1):
            cell = ws_sub.cell(row=3, column=c_idx, value=h)
            cell.font = font_hdr
            cell.fill = primary_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_cell
        ws_sub.row_dimensions[3].height = 22

        sub_start_row = 4
        sub_r_idx = sub_start_row
        for sub_dept, sub_group in p_data.groupby('SubDept'):
            stot = len(sub_group)
            spres = int(sub_group['IsPresent'].sum())
            sleave = int(sub_group['IsLeave'].sum())
            sabsent = max(0, stot - spres - sleave)
            slate = int(sub_group['IsLate'].sum())

            sec_sheet_name = sub_dept_sheets[sub_dept]

            ws_sub.cell(row=sub_r_idx, column=1, value=sub_dept).font = font_reg
            ws_sub.cell(row=sub_r_idx, column=2, value=stot).font = font_reg
            ws_sub.cell(row=sub_r_idx, column=3, value=spres).font = font_reg
            ws_sub.cell(row=sub_r_idx, column=4, value=sabsent).font = font_reg
            ws_sub.cell(row=sub_r_idx, column=5, value=sleave).font = font_reg
            ws_sub.cell(row=sub_r_idx, column=6, value=slate).font = font_reg

            act_cell = ws_sub.cell(row=sub_r_idx, column=7, value=" VIEW SECTIONS ")
            act_cell.font = btn_blue_font
            act_cell.fill = btn_blue_fill
            act_cell.border = btn_blue_border
            act_cell.alignment = Alignment(horizontal="center", vertical="center")
            act_cell.hyperlink = f"#'{sec_sheet_name}'!A1"

            for c in range(1, 7):
                c_cell = ws_sub.cell(row=sub_r_idx, column=c)
                c_cell.border = border_cell
                if c > 1:
                    c_cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c_cell.alignment = Alignment(horizontal="left", vertical="center")
            ws_sub.row_dimensions[sub_r_idx].height = 24
            sub_r_idx += 1

        sub_end_row = sub_r_idx - 1

        ws_sub.cell(row=sub_r_idx, column=1, value="TOTAL").font = font_grand
        ws_sub.cell(row=sub_r_idx, column=1).alignment = Alignment(horizontal="center", vertical="center")
        for c_idx, col_letter in enumerate(['B', 'C', 'D', 'E', 'F'], 2):
            cell = ws_sub.cell(row=sub_r_idx, column=c_idx, value=f"=SUM({col_letter}{sub_start_row}:{col_letter}{sub_end_row})")
            cell.font = font_grand
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_sub.cell(row=sub_r_idx, column=7, value="").fill = primary_fill

        for c in range(1, 8):
            c_cell = ws_sub.cell(row=sub_r_idx, column=c)
            c_cell.border = border_grand
            if c < 7:
                c_cell.fill = primary_fill
        ws_sub.row_dimensions[sub_r_idx].height = 25

        ws_sub.column_dimensions['A'].width = 40
        for col in ['B', 'C', 'D', 'E', 'F']:
            ws_sub.column_dimensions[col].width = 18
        ws_sub.column_dimensions['G'].width = 24

        for sub_dept, sub_group in p_data.groupby('SubDept'):
            sec_sheet_name = sub_dept_sheets[sub_dept]
            ws_sec = wb.create_sheet(title=sec_sheet_name)
            ws_sec.views.sheetView[0].showGridLines = True
            ws_sec.merge_cells("A1:G1")
            ws_sec["A1"] = f"{sub_dept.upper()} - SECTIONS SUMMARY"
            ws_sec["A1"].font = font_title
            ws_sec["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws_sec.row_dimensions[1].height = 25

            b_cell = ws_sec.cell(row=2, column=1, value=" ◀ BACK TO SUB-DEPTS ")
            b_cell.font = btn_dark_font
            b_cell.fill = btn_dark_fill
            b_cell.border = btn_dark_border
            b_cell.alignment = Alignment(horizontal="center", vertical="center")
            b_cell.hyperlink = f"#'{sub_sheet_name}'!A1"

            sec_headers = ["Year & Section", "Total Strength", "Present", "Absent", "Leave", "Late", "Action Button"]
            for c_idx, h in enumerate(sec_headers, 1):
                cell = ws_sec.cell(row=3, column=c_idx, value=h)
                cell.font = font_hdr
                cell.fill = primary_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border_cell
            ws_sec.row_dimensions[3].height = 22

            sec_start_row = 4
            sec_r_idx = sec_start_row
            for dept_code, sec_group in sub_group.groupby('DeptCode'):
                dtot = len(sec_group)
                dpres = int(sec_group['IsPresent'].sum())
                dleave = int(sec_group['IsLeave'].sum())
                dabsent = max(0, dtot - dpres - dleave)
                dlate = int(sec_group['IsLate'].sum())

                detail_sheet_name = section_sheets[dept_code]

                ws_sec.cell(row=sec_r_idx, column=1, value=dept_code).font = font_reg
                ws_sec.cell(row=sec_r_idx, column=2, value=dtot).font = font_reg
                ws_sec.cell(row=sec_r_idx, column=3, value=dpres).font = font_reg
                ws_sec.cell(row=sec_r_idx, column=4, value=dabsent).font = font_reg
                ws_sec.cell(row=sec_r_idx, column=5, value=dleave).font = font_reg
                ws_sec.cell(row=sec_r_idx, column=6, value=dlate).font = font_reg

                d_act = ws_sec.cell(row=sec_r_idx, column=7, value=" VIEW DETAILS ")
                d_act.font = btn_blue_font
                d_act.fill = btn_blue_fill
                d_act.border = btn_blue_border
                d_act.alignment = Alignment(horizontal="center", vertical="center")
                d_act.hyperlink = f"#'{detail_sheet_name}'!A1"

                for c in range(1, 7):
                    c_cell = ws_sec.cell(row=sec_r_idx, column=c)
                    c_cell.border = border_cell
                    if c > 1:
                        c_cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        c_cell.alignment = Alignment(horizontal="left", vertical="center")
                ws_sec.row_dimensions[sec_r_idx].height = 24
                sec_r_idx += 1

            sec_end_row = sec_r_idx - 1

            ws_sec.cell(row=sec_r_idx, column=1, value="TOTAL").font = font_grand
            ws_sec.cell(row=sec_r_idx, column=1).alignment = Alignment(horizontal="center", vertical="center")
            for c_idx, col_letter in enumerate(['B', 'C', 'D', 'E', 'F'], 2):
                cell = ws_sec.cell(row=sec_r_idx, column=c_idx, value=f"=SUM({col_letter}{sec_start_row}:{col_letter}{sec_end_row})")
                cell.font = font_grand
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws_sec.cell(row=sec_r_idx, column=7, value="").fill = primary_fill

            for c in range(1, 8):
                c_cell = ws_sec.cell(row=sec_r_idx, column=c)
                c_cell.border = border_grand
                if c < 7:
                    c_cell.fill = primary_fill
            ws_sec.row_dimensions[sec_r_idx].height = 25

            ws_sec.column_dimensions['A'].width = 30
            for col in ['B', 'C', 'D', 'E', 'F']:
                ws_sec.column_dimensions[col].width = 18
            ws_sec.column_dimensions['G'].width = 22

            for dept_code, sec_group in sub_group.groupby('DeptCode'):
                detail_sheet_name = section_sheets[dept_code]
                abs_sheet_name = absent_sheets[dept_code]
                late_sheet_name = late_sheets[dept_code]

                ws_det = wb.create_sheet(title=detail_sheet_name)
                ws_det.views.sheetView[0].showGridLines = True
                ws_det.merge_cells("A1:F1")
                ws_det["A1"] = f"SECTION ROSTER: {sub_dept} ({dept_code})"
                ws_det["A1"].font = font_title
                ws_det["A1"].alignment = Alignment(horizontal="center", vertical="center")
                ws_det.row_dimensions[1].height = 25

                det_back = ws_det.cell(row=2, column=1, value=" ◀ BACK TO SECTIONS ")
                det_back.font = btn_dark_font
                det_back.fill = btn_dark_fill
                det_back.border = btn_dark_border
                det_back.alignment = Alignment(horizontal="center", vertical="center")
                det_back.hyperlink = f"#'{sec_sheet_name}'!A1"

                abs_link = ws_det.cell(row=2, column=3, value=" VIEW ABSENT STUDENTS ")
                abs_link.font = btn_amber_font
                abs_link.fill = btn_amber_fill
                abs_link.border = btn_amber_border
                abs_link.alignment = Alignment(horizontal="center", vertical="center")
                abs_link.hyperlink = f"#'{abs_sheet_name}'!A1"

                late_link = ws_det.cell(row=2, column=5, value=" VIEW LATE STUDENTS ")
                late_link.font = btn_amber_font
                late_link.fill = btn_amber_fill
                late_link.border = btn_amber_border
                late_link.alignment = Alignment(horizontal="center", vertical="center")
                late_link.hyperlink = f"#'{late_sheet_name}'!A1"
                ws_det.row_dimensions[2].height = 26

                det_headers = ["Employee ID", "Employee Name", "Sub-Department", "Year & Section", "Clock-In Time", "Status"]
                for c_idx, h in enumerate(det_headers, 1):
                    cell = ws_det.cell(row=3, column=c_idx, value=h)
                    cell.font = font_hdr
                    cell.fill = primary_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = border_cell
                ws_det.row_dimensions[3].height = 22

                green_fill = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")
                green_font = Font(name="Calibri", size=11, color="0F5132", bold=True)
                amber_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                amber_font = Font(name="Calibri", size=11, color="664D03", bold=True)
                red_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
                red_font = Font(name="Calibri", size=11, color="842029", bold=True)

                d_r_idx = 4
                for _, row in sec_group.iterrows():
                    is_pres = row['IsPresent'] == 1
                    is_leave = row['IsLeave'] == 1
                    is_late = row['IsLate'] == 1

                    if is_late:
                        status_str = f"LATE ({row['ClockIn']})"
                        s_fill, s_font = amber_fill, amber_font
                    elif is_pres:
                        status_str = f"PRESENT ({row['ClockIn']})"
                        s_fill, s_font = green_fill, green_font
                    elif is_leave:
                        status_str = "LEAVE (Excused)"
                        s_fill, s_font = amber_fill, amber_font
                    else:
                        status_str = "ABSENT (--:--)"
                        s_fill, s_font = red_fill, red_font

                    ws_det.cell(row=d_r_idx, column=1, value=str(row['EmpID'])).font = font_reg
                    ws_det.cell(row=d_r_idx, column=2, value=str(row['Name'])).font = font_reg
                    ws_det.cell(row=d_r_idx, column=3, value=str(row['SubDept'])).font = font_reg
                    ws_det.cell(row=d_r_idx, column=4, value=f"{row['Year']} - {row['Section']}").font = font_reg
                    ws_det.cell(row=d_r_idx, column=5, value=str(row['ClockIn']) if row['ClockIn'] else "--:--").font = font_reg
                    
                    status_cell = ws_det.cell(row=d_r_idx, column=6, value=status_str)
                    status_cell.font = s_font
                    status_cell.fill = s_fill

                    for c in range(1, 7):
                        c_cell = ws_det.cell(row=d_r_idx, column=c)
                        c_cell.border = border_cell
                        if c != 6:
                            c_cell.alignment = Alignment(horizontal="center" if c > 2 else "left", vertical="center")
                        else:
                            c_cell.alignment = Alignment(horizontal="center", vertical="center")
                    ws_det.row_dimensions[d_r_idx].height = 20
                    d_r_idx += 1

                for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                    ws_det.column_dimensions[col].width = 22

                # Absent Sub-sheet
                abs_group = sec_group[(sec_group['IsPresent'] == 0) & (sec_group['IsLeave'] == 0)]
                ws_abs = wb.create_sheet(title=abs_sheet_name)
                ws_abs.views.sheetView[0].showGridLines = True
                ws_abs.merge_cells("A1:F1")
                ws_abs["A1"] = f"ABSENT STUDENTS: {dept_code}"
                ws_abs["A1"].font = font_title
                ws_abs["A1"].alignment = Alignment(horizontal="center", vertical="center")
                ws_abs.row_dimensions[1].height = 25

                abs_back = ws_abs.cell(row=2, column=1, value=" ◀ BACK TO ROSTER ")
                abs_back.font = btn_dark_font
                abs_back.fill = btn_dark_fill
                abs_back.border = btn_dark_border
                abs_back.alignment = Alignment(horizontal="center", vertical="center")
                abs_back.hyperlink = f"#'{detail_sheet_name}'!A1"
                ws_abs.row_dimensions[2].height = 24

                for c_idx, h in enumerate(det_headers, 1):
                    cell = ws_abs.cell(row=3, column=c_idx, value=h)
                    cell.font = font_hdr
                    cell.fill = primary_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = border_cell
                ws_abs.row_dimensions[3].height = 22

                a_r_idx = 4
                for _, row in abs_group.iterrows():
                    ws_abs.cell(row=a_r_idx, column=1, value=str(row['EmpID'])).font = font_reg
                    ws_abs.cell(row=a_r_idx, column=2, value=str(row['Name'])).font = font_reg
                    ws_abs.cell(row=a_r_idx, column=3, value=str(row['SubDept'])).font = font_reg
                    ws_abs.cell(row=a_r_idx, column=4, value=f"{row['Year']} - {row['Section']}").font = font_reg
                    ws_abs.cell(row=a_r_idx, column=5, value="--:--").font = font_reg
                    
                    sc = ws_abs.cell(row=a_r_idx, column=6, value="ABSENT")
                    sc.font = red_font
                    sc.fill = red_fill

                    for c in range(1, 7):
                        c_cell = ws_abs.cell(row=a_r_idx, column=c)
                        c_cell.border = border_cell
                        c_cell.alignment = Alignment(horizontal="center" if c > 2 else "left", vertical="center")
                    ws_abs.row_dimensions[a_r_idx].height = 20
                    a_r_idx += 1

                for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                    ws_abs.column_dimensions[col].width = 22

                # Late Sub-sheet
                late_group = sec_group[sec_group['IsLate'] == 1]
                ws_late = wb.create_sheet(title=late_sheet_name)
                ws_late.views.sheetView[0].showGridLines = True
                ws_late.merge_cells("A1:F1")
                ws_late["A1"] = f"LATE STUDENTS: {dept_code}"
                ws_late["A1"].font = font_title
                ws_late["A1"].alignment = Alignment(horizontal="center", vertical="center")
                ws_late.row_dimensions[1].height = 25

                late_back = ws_late.cell(row=2, column=1, value=" ◀ BACK TO ROSTER ")
                late_back.font = btn_dark_font
                late_back.fill = btn_dark_fill
                late_back.border = btn_dark_border
                late_back.alignment = Alignment(horizontal="center", vertical="center")
                late_back.hyperlink = f"#'{detail_sheet_name}'!A1"
                ws_late.row_dimensions[2].height = 24

                for c_idx, h in enumerate(det_headers, 1):
                    cell = ws_late.cell(row=3, column=c_idx, value=h)
                    cell.font = font_hdr
                    cell.fill = primary_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = border_cell
                ws_late.row_dimensions[3].height = 22

                l_r_idx = 4
                for _, row in late_group.iterrows():
                    ws_late.cell(row=l_r_idx, column=1, value=str(row['EmpID'])).font = font_reg
                    ws_late.cell(row=l_r_idx, column=2, value=str(row['Name'])).font = font_reg
                    ws_late.cell(row=l_r_idx, column=3, value=str(row['SubDept'])).font = font_reg
                    ws_late.cell(row=l_r_idx, column=4, value=f"{row['Year']} - {row['Section']}").font = font_reg
                    ws_late.cell(row=l_r_idx, column=5, value=str(row['ClockIn'])).font = font_reg
                    
                    sc = ws_late.cell(row=l_r_idx, column=6, value=f"LATE ({row['ClockIn']})")
                    sc.font = amber_font
                    sc.fill = amber_fill

                    for c in range(1, 7):
                        c_cell = ws_late.cell(row=l_r_idx, column=c)
                        c_cell.border = border_cell
                        c_cell.alignment = Alignment(horizontal="center" if c > 2 else "left", vertical="center")
                    ws_late.row_dimensions[l_r_idx].height = 20
                    l_r_idx += 1

                for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                    ws_late.column_dimensions[col].width = 22

    if default_sheet.title in wb.sheetnames and default_sheet != ws_master:
        wb.remove(default_sheet)

    wb.save(output_path)

def generate_excel_report(df, leave_mapping, output_path):
    df = apply_leave_mapping(df, leave_mapping)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daily Attendance Report"
    ws.views.sheetView[0].showGridLines = True

    primary_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    subtotal_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    font_title = Font(name="Calibri", size=15, bold=True, color="1E293B")
    font_dept_hdr = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_col_hdr = Font(name="Calibri", size=11, bold=True, color="1E293B")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    font_grand = Font(name="Calibri", size=12, bold=True, color="FFFFFF")

    thin_side = Side(border_style="thin", color="CBD5E1")
    double_bottom = Side(border_style="double", color="1E293B")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_grand = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_bottom)

    report_date = df['Date'].iloc[0] if not df.empty else ''
    full_shift_title = df['ShiftName'].iloc[0] if not df.empty else ''

    ws.merge_cells("A1:F1")
    ws["A1"] = f"Daily Attendance Report - {full_shift_title} - {report_date}"
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    current_row = 3
    subtotal_rows = []

    for parent in PARENT_ORDER:
        p_data = df[df['ParentDept'] == parent]
        if p_data.empty:
            continue

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
        cell_p = ws.cell(row=current_row, column=1, value=parent)
        cell_p.font = font_dept_hdr
        cell_p.fill = primary_fill
        cell_p.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[current_row].height = 22
        current_row += 1

        headers = ["", "Total Strength", "Present", "Absent", "Leave", "Late"]
        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=c_idx, value=h)
            cell.font = font_col_hdr
            cell.alignment = Alignment(horizontal="center" if c_idx > 1 else "left", vertical="center")
            cell.border = border_cell
        ws.row_dimensions[current_row].height = 20
        current_row += 1

        sec_start_row = current_row
        for dept_code, sec_group in p_data.groupby('DeptCode'):
            tot = len(sec_group)
            pres = sec_group['IsPresent'].sum()
            leave = sec_group['IsLeave'].sum()
            absent = max(0, tot - pres - leave)
            late = sec_group['IsLate'].sum()

            ws.cell(row=current_row, column=1, value=dept_code).font = font_regular
            ws.cell(row=current_row, column=2, value=tot).font = font_regular
            ws.cell(row=current_row, column=3, value=pres).font = font_regular
            ws.cell(row=current_row, column=4, value=absent).font = font_regular
            ws.cell(row=current_row, column=5, value=leave).font = font_regular
            ws.cell(row=current_row, column=6, value=late).font = font_regular

            for c in range(1, 7):
                c_cell = ws.cell(row=current_row, column=c)
                c_cell.border = border_cell
                if c > 1:
                    c_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[current_row].height = 20
            current_row += 1

        sec_end_row = current_row - 1
        ws.cell(row=current_row, column=1, value="Subtotal").font = font_bold
        ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(row=current_row, column=2, value=f"=SUM(B{sec_start_row}:B{sec_end_row})").font = font_bold
        ws.cell(row=current_row, column=3, value=f"=SUM(C{sec_start_row}:C{sec_end_row})").font = font_bold
        ws.cell(row=current_row, column=4, value=f"=SUM(D{sec_start_row}:D{sec_end_row})").font = font_bold
        ws.cell(row=current_row, column=5, value=f"=SUM(E{sec_start_row}:E{sec_end_row})").font = font_bold
        ws.cell(row=current_row, column=6, value=f"=SUM(F{sec_start_row}:F{sec_end_row})").font = font_bold

        for c in range(2, 7):
            ws.cell(row=current_row, column=c).alignment = Alignment(horizontal="center", vertical="center")
        for c in range(1, 7):
            c_cell = ws.cell(row=current_row, column=c)
            c_cell.border = border_cell
            c_cell.fill = subtotal_fill
        ws.row_dimensions[current_row].height = 22
        subtotal_rows.append(current_row)
        current_row += 1

    ws.cell(row=current_row, column=1, value="GRAND TOTAL").font = font_grand
    ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
    cols = ['B', 'C', 'D', 'E', 'F']
    for c_idx, col_letter in enumerate(cols, 2):
        cell_refs = ",".join([f"{col_letter}{r}" for r in subtotal_rows])
        ws.cell(row=current_row, column=c_idx, value=f"=SUM({cell_refs})").font = font_grand
        ws.cell(row=current_row, column=c_idx).alignment = Alignment(horizontal="center", vertical="center")

    for c in range(1, 7):
        c_cell = ws.cell(row=current_row, column=c)
        c_cell.border = border_grand
        c_cell.fill = primary_fill

    ws.row_dimensions[current_row].height = 24
    ws.column_dimensions['A'].width = 38
    for col in ['B', 'C', 'D', 'E', 'F']:
        ws.column_dimensions[col].width = 16

    wb.save(output_path)

def generate_all_reports_suite(df, leave_mapping):
    df = apply_leave_mapping(df, leave_mapping)
    
    run_timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    shift_tag = df['ShiftName'].iloc[0].replace(' ', '_').replace('(', '').replace(')', '')
    output_dir = os.path.join("Attendance_Records", f"Attendance_Report_{run_timestamp}_{shift_tag}")
    os.makedirs(output_dir, exist_ok=True)

    save_records_to_db(df)
    generate_interactive_master_workbook(df, os.path.join(output_dir, "Master_Hierarchical_Attendance.xlsx"))
    generate_excel_report(df, leave_mapping, os.path.join(output_dir, "Daily_Attendance_Executive_Summary.xlsx"))

    return os.path.abspath(output_dir)

# --- FINE CALCULATION LOGIC ---
def calculate_student_fine_breakdown(emp_df, is_inliving=False):
    absent_records = emp_df[emp_df['is_absent'] == 1].copy()
    late_records = emp_df[emp_df['is_late'] == 1].copy()
    
    absent_count = len(absent_records)
    late_count = len(late_records)
    
    if is_inliving:
        fine_absent = absent_count * 200
        fine_late = late_count * 200
        return absent_count, late_count, fine_absent, fine_late

    absent_dates = []
    for _, r in absent_records.iterrows():
        try:
            absent_dates.append(datetime.strptime(r['record_date'][:10], '%Y-%m-%d'))
        except Exception:
            pass
            
    absent_dates = sorted(list(set(absent_dates)))
    
    fine_absent = 0
    full_weeks_count = 0
    i = 0
    n = len(absent_dates)
    processed = set()
    
    while i < n:
        d = absent_dates[i]
        if d.weekday() == 0 and (i + 4) < n:
            block = [absent_dates[i+k] for k in range(5)]
            if [dt.weekday() for dt in block] == [0, 1, 2, 3, 4] and all((block[k+1] - block[k]).days == 1 for k in range(4)):
                fine_absent += (5 * 85)
                full_weeks_count += 1
                for dt in block:
                    processed.add(dt)
                i += 5
                continue
        i += 1
        
    for d in absent_dates:
        if d not in processed:
            wd = d.weekday()
            if wd in [1, 2, 3]:
                fine_absent += 50
            elif wd in [0, 4]:
                fine_absent += 70
                
    fine_late = late_count * 100
    return absent_count, late_count, fine_absent, fine_late

def generate_fine_excel(report_data, section_code, date_range_str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Monthly Fine Settlement"
    ws.views.sheetView[0].showGridLines = True

    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    font_title = Font(name="Century Gothic", size=13, bold=True, color="1E293B")
    font_sub = Font(name="Century Gothic", size=10, bold=True, color="475569")
    font_hdr = Font(name="Century Gothic", size=9, bold=True, color="FFFFFF")
    font_reg = Font(name="Century Gothic", size=9)
    font_grand = Font(name="Century Gothic", size=10, bold=True, color="FFFFFF")

    primary_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    thin_side = Side(border_style="thin", color="CBD5E1")
    double_bottom = Side(border_style="double", color="1E293B")
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_grand = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_bottom)

    ws.merge_cells("A1:J1")
    ws["A1"] = f"MONTHLY FINE SETTLEMENT SHEET - SECTION: {section_code}"
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:J2")
    ws["A2"] = f"Billing Cycle: {date_range_str}"
    ws["A2"].font = font_sub
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    headers = ["Student ID", "Student Name", "Abs", "Late", "Auto Fine", "Disc. Fine", "Waiver", "Total (Rs.)", "Warnings", "Remarks"]
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=c_idx, value=h)
        cell.font = font_hdr
        cell.fill = primary_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_cell
    ws.row_dimensions[4].height = 22

    r_idx = 5
    for item in report_data:
        ws.cell(row=r_idx, column=1, value=str(item['ID'])).font = font_reg
        ws.cell(row=r_idx, column=2, value=str(item['Name'])).font = font_reg
        ws.cell(row=r_idx, column=3, value=item['Absentees']).font = font_reg
        ws.cell(row=r_idx, column=4, value=item['Late']).font = font_reg
        ws.cell(row=r_idx, column=5, value=item['Auto_Fine']).font = font_reg
        ws.cell(row=r_idx, column=6, value=item['Disc_Fine']).font = font_reg
        ws.cell(row=r_idx, column=7, value=item['Waiver']).font = font_reg
        ws.cell(row=r_idx, column=8, value=f"=MAX(0, E{r_idx}+F{r_idx}-G{r_idx})").font = font_reg
        ws.cell(row=r_idx, column=9, value=item['Warnings']).font = font_reg
        ws.cell(row=r_idx, column=10, value=item['Remarks']).font = font_reg

        for c in range(1, 11):
            c_cell = ws.cell(row=r_idx, column=c)
            c_cell.border = border_cell
            if c in [3, 4, 5, 6, 7, 8, 9]:
                c_cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[r_idx].height = 18
        r_idx += 1

    end_row = r_idx - 1
    ws.cell(row=r_idx, column=1, value="GRAND TOTAL").font = font_grand
    ws.cell(row=r_idx, column=1).alignment = Alignment(horizontal="center", vertical="center")

    for c_idx, col_letter in enumerate(['C', 'D', 'E', 'F', 'G', 'H'], 3):
        cell = ws.cell(row=r_idx, column=c_idx, value=f"=SUM({col_letter}5:{col_letter}{end_row})")
        cell.font = font_grand
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for c in range(1, 11):
        c_cell = ws.cell(row=r_idx, column=c)
        c_cell.border = border_grand
        c_cell.fill = primary_fill
    ws.row_dimensions[r_idx].height = 22

    widths = {'A': 13, 'B': 18, 'C': 6, 'D': 6, 'E': 10, 'F': 10, 'G': 9, 'H': 11, 'I': 14, 'J': 16}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- STREAMLIT UI TABS ---
st.title("📊 Enterprise Attendance Suite & Disciplinary Tracking")

tab1, tab2, tab3, tab4 = st.tabs([
    "🚀 Process Daily Log", 
    "🚨 Warning & Leave Exemption System", 
    "🔍 Student Profile Search", 
    "💰 Monthly Fine Settlement"
])

# --- TAB 1: DAILY LOG ---
with tab1:
    st.markdown("### Upload Raw BioTime Log")
    uploaded_file = st.file_uploader("Upload ZKBioTime Excel File (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            df, shift_name, late_thresh = parse_biotime_excel(uploaded_file)
            if not df.empty:
                st.success(f"Detected **{shift_name}** | Parsed **{len(df)}** student records.")
                
                available_codes = set(df['DeptCode'].unique().tolist())
                ordered_dept_codes = []
                for parent in PARENT_ORDER:
                    p_codes = sorted(df[df['ParentDept'] == parent]['DeptCode'].unique().tolist())
                    ordered_dept_codes.extend(p_codes)
                
                for code in sorted(list(available_codes)):
                    if code not in ordered_dept_codes:
                        ordered_dept_codes.append(code)

                st.markdown("---")
                st.markdown("### 📝 Enter Department Leave Counts")
                leave_mapping = {}
                cols = st.columns(2)
                for idx, code in enumerate(ordered_dept_codes):
                    with cols[idx % 2]:
                        leave_mapping[code] = st.number_input(f"[{code}] Leave Count", min_value=0, value=0, step=1, key=f"tab1_{code}")

                st.markdown("---")
                if st.button("🚀 Process Log & Update History Database", type="primary", use_container_width=True):
                    out_folder = generate_all_reports_suite(df, leave_mapping)
                    st.session_state['processed'] = True
                    st.session_state['out_folder'] = out_folder
                    st.session_state['shift_title'] = df['ShiftName'].iloc[0] if not df.empty else shift_name
                    st.session_state['report_date'] = df['Date'].iloc[0] if not df.empty else datetime.now().strftime('%Y-%m-%d')
                    st.session_state['df_raw_saved'] = df
                    st.session_state['leave_mapping_saved'] = leave_mapping
                    st.success(f"Log saved to database & created folder: `{out_folder}`")

                if st.session_state.get('processed', False):
                    st.markdown("### 🖨️ Daily Summary Preview")
                    report_date = st.session_state.get('report_date', '')
                    shift_title = st.session_state.get('shift_title', '')
                    out_folder = st.session_state.get('out_folder', 'Attendance_Records')
                    saved_df = st.session_state.get('df_raw_saved', df)
                    saved_leaves = st.session_state.get('leave_mapping_saved', leave_mapping)

                    st.info(f"📁 **Active Run Folder:** `{out_folder}`")
                    html_report = build_html_report(saved_df, saved_leaves, shift_title, report_date)
                    st.components.v1.html(html_report, height=450, scrolling=True)

                    col_actions_1, col_actions_2, col_actions_3 = st.columns(3)
                    with col_actions_1:
                        exec_file = os.path.join(out_folder, "Daily_Attendance_Executive_Summary.xlsx")
                        if os.path.exists(exec_file):
                            with open(exec_file, "rb") as f:
                                st.download_button(
                                    label="📥 Download Executive Workbook",
                                    data=f,
                                    file_name=os.path.basename(exec_file),
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )

                    with col_actions_2:
                        if st.button("📂 Open This Run's Folder", use_container_width=True):
                            try:
                                os.startfile(out_folder)
                                st.toast("Opened folder successfully!")
                            except Exception as ex:
                                st.error(f"Could not open folder automatically: {ex}")

                    with col_actions_3:
                        print_script = (
                            "<script>"
                            "function printReport() {"
                            "    var myWindow = window.open('', '_blank', 'width=800,height=600');"
                            f"    myWindow.document.write(`{html_report}`);"
                            "    myWindow.document.close();"
                            "    myWindow.focus();"
                            "    setTimeout(function() { myWindow.print(); myWindow.close(); }, 500);"
                            "}"
                            "</script>"
                            "<button onclick=\"printReport()\" style=\""
                            "    width: 100%; background-color: #8B5CF6; color: white; padding: 10px 15px;"
                            "    border: none; border-radius: 8px; font-size: 15px; font-weight: bold; cursor: pointer;"
                            "\">🖨️ Print One-Page Report</button>"
                        )
                        st.components.v1.html(print_script, height=60)

        except Exception as e:
            st.error(f"Error processing file: {e}")

# --- TAB 2: WARNINGS & LEAVE EXEMPTION ---
with tab2:
    st.markdown("### 🚨 Disciplinary Warning & Leave Exemption Portal")
    conn = sqlite3.connect(DB_FILE)
    df_db = pd.read_sql_query("SELECT * FROM attendance_logs", conn)
    conn.close()

    if df_db.empty:
        st.info("No attendance records in history database yet. Process a daily log first!")
    else:
        abs_summary = df_db.groupby(['emp_id', 'name', 'dept_code', 'sub_dept']).agg(
            Total_Days=('record_date', 'nunique'),
            Total_Absent=('is_absent', 'sum'),
            Total_Late=('is_late', 'sum'),
            Total_Leave=('is_leave', 'sum')
        ).reset_index()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Logged Days", df_db['record_date'].nunique())
        c2.metric("1st Warning (4+ Abs)", len(abs_summary[abs_summary['Total_Absent'] >= 4]))
        c3.metric("2nd Warning (8+ Abs)", len(abs_summary[abs_summary['Total_Absent'] >= 8]))
        c4.metric("Withdrawal Notice (12+ Abs)", len(abs_summary[abs_summary['Total_Absent'] >= 12]))

        st.markdown("---")
        st.markdown("#### 🚩 Personnel Requiring Action")

        warning_filter = st.radio("Filter Stage:", ["All Warnings", "1st Warning (4-7 Absences)", "2nd Warning (8-11 Absences)", "Withdrawal Notice (12+ Absences)"])
        
        if warning_filter == "1st Warning (4-7 Absences)":
            flagged = abs_summary[(abs_summary['Total_Absent'] >= 4) & (abs_summary['Total_Absent'] < 8)]
        elif warning_filter == "2nd Warning (8-11 Absences)":
            flagged = abs_summary[(abs_summary['Total_Absent'] >= 8) & (abs_summary['Total_Absent'] < 12)]
        elif warning_filter == "Withdrawal Notice (12+ Absences)":
            flagged = abs_summary[abs_summary['Total_Absent'] >= 12]
        else:
            flagged = abs_summary[abs_summary['Total_Absent'] >= 4]

        sorted_flagged = flagged.sort_values(by='Total_Absent', ascending=False)
        st.dataframe(sorted_flagged, use_container_width=True)

        if not sorted_flagged.empty:
            export_filename = f"Disciplinary_Warning_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                sorted_flagged.to_excel(writer, sheet_name='Warning_List', index=False)
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Export Filtered Warning List to Excel (For Department Verification)",
                data=processed_data,
                file_name=export_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        st.markdown("---")
        st.markdown("### 🏥 Convert Absence to Approved/Medical Leave")
        st.caption("Select a flagged student to view their absent dates and convert medical proofs into excused leaves.")

        flagged_emp_list = sorted_flagged['emp_id'].tolist() if not sorted_flagged.empty else sorted(df_db[df_db['is_absent'] == 1]['emp_id'].unique().tolist())
        
        if flagged_emp_list:
            selected_exemption_emp = st.selectbox("Select Student / Employee ID for Exemption:", flagged_emp_list, key="exempt_select")

            if selected_exemption_emp:
                emp_absent_records = df_db[(df_db['emp_id'] == selected_exemption_emp) & (df_db['is_absent'] == 1)].sort_values(by='record_date')
                emp_info = df_db[df_db['emp_id'] == selected_exemption_emp].iloc[0]

                st.write(f"**Student Name:** {emp_info['name']} | **Dept:** {emp_info['dept_code']}")

                if emp_absent_records.empty:
                    st.success("This student has no recorded unexcused absences!")
                else:
                    absent_dates_with_shifts = emp_absent_records[['record_date', 'shift_name']].drop_duplicates()
                    date_shift_options = [f"{row['record_date']} ({row['shift_name']})" for _, row in absent_dates_with_shifts.iterrows()]
                    
                    selected_options = st.multiselect("Select Date & Shift to Excuse / Convert to Leave:", date_shift_options)
                    reason_input = st.text_input("Approval Reason / Medical Certificate Details:", value="Medical Certificate Provided")

                    if st.button("✅ Convert Selected Dates to Excused Leave", type="primary"):
                        if not selected_options:
                            st.warning("Please select at least one date/shift to convert.")
                        else:
                            for opt in selected_options:
                                d_str = opt.split(" (")[0]
                                s_name = opt.split(" (")[1].rstrip(")")
                                convert_absence_to_leave(selected_exemption_emp, [d_str], s_name, reason_input)
                            st.success(f"Successfully converted selected record(s) to Approved Leave!")
                            st.rerun()
        else:
            st.info("No students currently have unexcused absences.")

# --- TAB 3: STUDENT SEARCH ---
with tab3:
    st.markdown("### 🔍 Individual Student Attendance Performance Profile")
    conn = sqlite3.connect(DB_FILE)
    df_db = pd.read_sql_query("SELECT * FROM attendance_logs", conn)
    conn.close()

    if not df_db.empty:
        emp_list = sorted(df_db['emp_id'].unique().tolist())
        selected_emp = st.selectbox("Select Employee / Student ID:", emp_list)

        if selected_emp:
            emp_records = df_db[df_db['emp_id'] == selected_emp].sort_values(by=['record_date', 'shift_name'])
            emp_name = emp_records['name'].iloc[0]
            emp_dept = emp_records['dept_code'].iloc[0]

            tot_days = len(emp_records)
            tot_pres = emp_records['is_present'].sum()
            tot_abs = emp_records['is_absent'].sum()
            tot_late = emp_records['is_late'].sum()
            tot_leave = emp_records['is_leave'].sum()

            pct = (tot_pres / max(tot_days, 1)) * 100

            st.subheader(f"Profile: {emp_name} ({selected_emp}) - {emp_dept}")
            
            p1, p2, p3, p4, p5 = st.columns(5)
            p1.metric("Attendance %", f"{pct:.1f}%")
            p2.metric("Total Present", tot_pres)
            p3.metric("Total Absent", tot_abs)
            p4.metric("Total Late", tot_late)
            p5.metric("Total Leave", tot_leave)

            st.markdown("#### Detailed Clock-In History Log")
            if 'excuse_reason' not in emp_records.columns:
                emp_records['excuse_reason'] = ''
            st.dataframe(emp_records[['record_date', 'shift_name', 'clock_in', 'is_present', 'is_absent', 'is_late', 'is_leave', 'excuse_reason']], use_container_width=True)

# --- TAB 4: MONTHLY FINE SETTLEMENT ENGINE ---
with tab4:
    st.markdown("### 💰 Monthly Fine Settlement & Disciplinary Billing Engine")
    st.caption("Generate automated section-wise monthly fine sheets (Default Cycle: 20th of previous month to 19th of current month).")

    conn = sqlite3.connect(DB_FILE)
    df_db = pd.read_sql_query("SELECT * FROM attendance_logs", conn)
    conn.close()

    if df_db.empty:
        st.info("No logs in database. Process daily attendance logs first!")
    else:
        c_sec, c_start, c_end = st.columns([2, 1.5, 1.5])
        
        sections_list = sorted(df_db['dept_code'].unique().tolist())
        with c_sec:
            selected_section = st.selectbox("Select Class / Section:", sections_list, key="fine_sec_select")
            
        today = date.today()
        if today.day >= 20:
            default_start = date(today.year, today.month, 20)
            if today.month == 12:
                default_end = date(today.year + 1, 1, 19)
            else:
                default_end = date(today.year, today.month + 1, 19)
        else:
            if today.month == 1:
                default_start = date(today.year - 1, 12, 20)
            else:
                default_start = date(today.year, today.month - 1, 20)
            default_end = date(today.year, today.month, 19)

        with c_start:
            start_d = st.date_input("Start Date (Cycle)", default_start)
        with c_end:
            end_d = st.date_input("End Date (Cycle)", default_end)

        start_str = start_d.strftime('%Y-%m-%d')
        end_str = end_d.strftime('%Y-%m-%d')
        cycle_label = f"{start_str} to {end_str}"

        st.markdown("---")

        sec_logs = df_db[
            (df_db['dept_code'] == selected_section) & 
            (df_db['record_date'] >= start_str) & 
            (df_db['record_date'] <= end_str)
        ]

        warning_map = {}
        for emp_id, emp_grp in df_db.groupby('emp_id'):
            tot_abs = emp_grp['is_absent'].sum()
            if tot_abs >= 12:
                warning_map[emp_id] = "Withdrawal Notice"
            elif tot_abs >= 8:
                warning_map[emp_id] = "2nd Warning"
            elif tot_abs >= 4:
                warning_map[emp_id] = "1st Warning"
            else:
                warning_map[emp_id] = "Normal"

        unique_students = sec_logs[['emp_id', 'name']].drop_duplicates().sort_values(by='emp_id')

        if unique_students.empty:
            st.warning(f"No attendance logs found for section **{selected_section}** between {start_str} and {end_str}.")
        else:
            st.markdown(f"#### 📋 Fine Settlement Roster — Section `{selected_section}`")
            st.info("💡 Enter your adjustments below, then click **Calculate & Refresh Totals** to update all amounts instantly.")

            with st.form(key=f"fine_form_{selected_section}"):
                grid_hdr = st.columns([1.1, 1.8, 0.6, 0.6, 0.9, 1.1, 1.2, 1.2, 1.3, 1.4, 2.0])
                headers = ["ID", "Name", "Abs", "Late", "In-Living?", "Auto Fine", "Disc.", "Waiver", "Total (Rs.)", "Warnings", "Remarks"]
                for idx, h in enumerate(headers):
                    grid_hdr[idx].markdown(f"**{h}**")

                form_inputs = []

                for idx, (_, st_row) in enumerate(unique_students.iterrows()):
                    e_id = str(st_row['emp_id'])
                    e_name = str(st_row['name'])
                    
                    emp_period_df = sec_logs[sec_logs['emp_id'] == e_id]
                    
                    r_cols = st.columns([1.1, 1.8, 0.6, 0.6, 0.9, 1.1, 1.2, 1.2, 1.3, 1.4, 2.0])
                    
                    r_cols[0].write(e_id)
                    r_cols[1].write(e_name)
                    
                    k_inliving = f"inliving_{selected_section}_{e_id}_{idx}"
                    k_disc = f"disc_{selected_section}_{e_id}_{idx}"
                    k_waiver = f"waiver_{selected_section}_{e_id}_{idx}"
                    k_rem = f"rem_{selected_section}_{e_id}_{idx}"
                    
                    is_inliving = r_cols[4].checkbox("", key=k_inliving)
                    
                    abs_cnt, late_cnt, absent_fine, late_fine = calculate_student_fine_breakdown(emp_period_df, is_inliving)
                    auto_fine = absent_fine + late_fine
                    
                    r_cols[2].write(abs_cnt)
                    r_cols[3].write(late_cnt)
                    r_cols[5].write(f"Rs. {auto_fine}")
                    
                    disc_fine = r_cols[6].number_input("", min_value=0, value=0, step=50, key=k_disc, label_visibility="collapsed")
                    waiver = r_cols[7].number_input("", min_value=0, value=0, step=50, key=k_waiver, label_visibility="collapsed")
                    
                    prov_total = max(0, auto_fine + disc_fine - waiver)
                    r_cols[8].markdown(f"**Rs. {prov_total}**")
                    
                    warn_status = warning_map.get(e_id, "Normal")
                    if warn_status == "Withdrawal Notice":
                        r_cols[9].error("Withdrawal")
                    elif warn_status == "2nd Warning":
                        r_cols[9].warning("2nd Warning")
                    elif warn_status == "1st Warning":
                        r_cols[9].info("1st Warning")
                    else:
                        r_cols[9].write("Normal")

                    remarks = r_cols[10].text_input("", placeholder="Notes...", key=k_rem, label_visibility="collapsed")

                    form_inputs.append({
                        'ID': e_id,
                        'Name': e_name,
                        'Abs': abs_cnt,
                        'Late': late_cnt,
                        'Auto_Fine': auto_fine,
                        'Disc_Fine': disc_fine,
                        'Waiver': waiver,
                        'Warnings': warn_status,
                        'Remarks': remarks
                    })

                st.markdown("---")
                submitted = st.form_submit_button("🧮 Calculate & Refresh Totals", type="primary", use_container_width=True)

            report_items = []
            tot_abs_sum = 0
            tot_late_sum = 0
            tot_auto_fine_sum = 0
            tot_disc_fine_sum = 0
            tot_waiver_sum = 0
            tot_grand_sum = 0

            for item in form_inputs:
                e_id = item['ID']
                auto_fine = item['Auto_Fine']
                disc_fine = item['Disc_Fine']
                waiver = item['Waiver']
                
                total_fine = max(0, auto_fine + disc_fine - waiver)

                tot_abs_sum += item['Abs']
                tot_late_sum += item['Late']
                tot_auto_fine_sum += auto_fine
                tot_disc_fine_sum += disc_fine
                tot_waiver_sum += waiver
                tot_grand_sum += total_fine

                report_items.append({
                    'ID': e_id,
                    'Name': item['Name'],
                    'Absentees': item['Abs'],
                    'Late': item['Late'],
                    'Auto_Fine': auto_fine,
                    'Disc_Fine': disc_fine,
                    'Waiver': waiver,
                    'Total': total_fine,
                    'Warnings': item['Warnings'],
                    'Remarks': item['Remarks']
                })

            st.markdown("### 🧮 Section Grand Totals")
            m1, m2, m3, m4, m5, m6 = st.columns(6)
            m1.metric("Total Abs", tot_abs_sum)
            m2.metric("Total Late", tot_late_sum)
            m3.metric("Auto Fines", f"Rs. {tot_auto_fine_sum:,}")
            m4.metric("Disc. Fines", f"Rs. {tot_disc_fine_sum:,}")
            m5.metric("Total Waivers", f"Rs. {tot_waiver_sum:,}")
            m6.metric("GRAND TOTAL", f"Rs. {tot_grand_sum:,}")

            st.markdown("---")
            
            col_exp1, col_exp2 = st.columns(2)
            
            with col_exp1:
                excel_bytes = generate_fine_excel(report_items, selected_section, cycle_label)
                st.download_button(
                    label=f"📥 Download Printable Excel Fine Sheet ({selected_section})",
                    data=excel_bytes,
                    file_name=f"Fine_Settlement_{selected_section}_{end_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )

            with col_exp2:
                fine_rows_html = ""
                for item in report_items:
                    fine_rows_html += f"""
                    <tr>
                        <td>{item['ID']}</td>
                        <td style="text-align: left;">{item['Name']}</td>
                        <td>{item['Absentees']}</td>
                        <td>{item['Late']}</td>
                        <td>Rs. {item['Auto_Fine']}</td>
                        <td>Rs. {item['Disc_Fine']}</td>
                        <td>Rs. {item['Total']}</td>
                        <td>{item['Warnings']}</td>
                        <td style="text-align: left;">{item['Remarks']}</td>
                    </tr>
                    """

                print_html = f"""
                <html>
                <head>
                    <style>
                        @page {{ size: A4 portrait; margin: 4mm; }}
                        body {{ font-family: 'Century Gothic', sans-serif; color: black; background: white; margin: 0; padding: 0; }}
                        h2 {{ text-align: center; margin: 0 0 2px 0; font-size: 13pt; color: #1E293B; }}
                        .sub {{ text-align: center; font-size: 9pt; font-weight: bold; margin-bottom: 6px; color: #475569; }}
                        table {{ width: 100%; border-collapse: collapse; page-break-inside: auto; }}
                        tr {{ page-break-inside: avoid; page-break-after: auto; }}
                        th, td {{ border: 1px solid #64748B; padding: 3px 3px; text-align: center; font-size: 8pt; white-space: nowrap; }}
                        th {{ background-color: #1E293B !important; color: white !important; font-weight: bold; -webkit-print-color-adjust: exact; }}
                        .grand {{ background-color: #1E293B !important; color: white !important; font-weight: bold; -webkit-print-color-adjust: exact; }}
                    </style>
                </head>
                <body>
                    <h2>MONTHLY FINE SETTLEMENT SHEET — SECTION: {selected_section}</h2>
                    <div class="sub">Billing Cycle: {cycle_label}</div>
                    <table>
                        <thead>
                            <tr>
                                <th>Student ID</th>
                                <th>Student Name</th>
                                <th>Abs</th>
                                <th>Late</th>
                                <th>Auto Fine</th>
                                <th>Disc.</th>
                                <th>Total (Rs.)</th>
                                <th>Warnings</th>
                                <th>Remarks</th>
                            </tr>
                        </thead>
                        <tbody>
                            {fine_rows_html}
                            <tr class="grand">
                                <td colspan="2">GRAND TOTAL</td>
                                <td>{tot_abs_sum}</td>
                                <td>{tot_late_sum}</td>
                                <td>Rs. {tot_auto_fine_sum:,}</td>
                                <td>Rs. {tot_disc_fine_sum:,}</td>
                                <td>Rs. {tot_grand_sum:,}</td>
                                <td colspan="2"></td>
                            </tr>
                        </tbody>
                    </table>
                </body>
                </html>
                """

                print_fine_script = (
                    "<script>"
                    "function printFineReport() {"
                    "    var myWindow = window.open('', '_blank', 'width=950,height=650');"
                    f"    myWindow.document.write(`{print_html}`);"
                    "    myWindow.document.close();"
                    "    myWindow.focus();"
                    "    setTimeout(function() {{ myWindow.print(); myWindow.close(); }}, 500);"
                    "}"
                    "</script>"
                    "<button onclick=\"printFineReport()\" style=\""
                    "    width: 100%; background-color: #8B5CF6; color: white; padding: 10px 15px;"
                    "    border: none; border-radius: 8px; font-size: 15px; font-weight: bold; cursor: pointer;"
                    "\">🖨️ Print Section Fine Sheet (A4 Portrait)</button>"
                )
                st.components.v1.html(print_fine_script, height=60)