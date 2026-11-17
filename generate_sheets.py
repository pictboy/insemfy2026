import os
import re
import pdfplumber
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_marks_workbook():
    base_dir = r"d:\apictnaly"
    output_filename = os.path.join(base_dir, "PICT_Student_Marks_Division_Wise.xlsx")
    pdf_path = os.path.join(base_dir, "V1_RollNo List-AY-2026-27_260916_185335.pdf")

    # 1. Parse PDF for Division Roll Lists
    pattern = re.compile(r'^(\d+)\s+([A-Z0-9]+)\s+(FY\d+)\s+(.+)$')
    pdf_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                m = pattern.match(line.strip())
                if m:
                    roll_no, mis_id, div, name = m.groups()
                    pdf_rows.append({
                        'Roll No': int(roll_no),
                        'MIS ID': mis_id.strip(),
                        'Division': div.strip(),
                        'Student Name': name.strip()
                    })

    pdf_df = pd.DataFrame(pdf_rows)
    print(f"Total students in PDF: {len(pdf_df)}")

    # 2. Parse Subject Excel Files
    excel_files = {
        'lac': ('Linear Algebra and Calculus (2601BS11).xlsx', 'LAC'),
        'cpps': ('C Programming for Problem Solving (2601PC11).xlsx', 'CPPS'),
        'mfr': ('Mechanics for Robotics (2601ES11).xlsx', 'MFR'),
        'qp': ('Quantum Physics (2601BS12).xlsx', 'QP'),
        'ieee': ('Integrated_Electrical_&_Electronics_Engineering_2601ES12.xlsx', 'IEEE'),
        'cst': ('Chemical Science and Technology (2601BS13).xlsx', 'CST'),
        'ese': ('Environment and Sustainable Engineering (2601VE11).xlsx', 'ESE'),
        'eg': ('Engineering Graphics (2601ES13).xlsx', 'EG')
    }

    marks = {code: {} for code in excel_files}
    all_students_master = {}

    for code, (fname, display_name) in excel_files.items():
        fpath = os.path.join(base_dir, fname)
        df = pd.read_excel(fpath)
        for _, row in df.iterrows():
            sc = str(row['Student Code']).strip()
            tot = str(row['Total Marks']).strip()
            val = int(tot.split('/')[0]) if '/' in tot else int(tot)
            marks[code][sc] = val
            if sc not in all_students_master:
                all_students_master[sc] = {
                    'Student Name': str(row['Student Name']).strip(),
                    'Student Code': sc
                }

    print("Loaded all subject marks successfully.")

    # 3. Create Workbook
    wb = openpyxl.Workbook()
    # remove default sheet
    default_sheet = wb.active

    # Styling Palettes
    navy_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    navy_accent_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    table_header_fill_1 = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    table_header_fill_2 = PatternFill(start_color="244062", end_color="244062", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    total_col_fill = PatternFill(start_color="E8F1F5", end_color="E8F1F5", fill_type="solid")
    stats_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    stats_hdr_fill = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")

    font_title = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    font_subtitle = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    font_table_hdr = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Segoe UI", size=10, color="000000")
    font_bold = Font(name="Segoe UI", size=10, bold=True, color="000000")
    font_absent = Font(name="Segoe UI", size=10, bold=True, color="C00000")
    font_stats = Font(name="Segoe UI", size=10, bold=True, color="1B365D")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    thin_border_side = Side(style='thin', color='D9D9D9')
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    thick_bottom = Border(bottom=Side(style='medium', color='1F4E78'))
    double_bottom = Border(top=thin_border_side, bottom=Side(style='double', color='1F4E78'))

    div_1_7_subjs = [('lac', 'LAC'), ('cpps', 'CPPS'), ('mfr', 'MFR'), ('qp', 'QP'), ('ieee', 'IEEE')]
    div_8_13_subjs = [('lac', 'LAC'), ('cpps', 'CPPS'), ('cst', 'CST'), ('ese', 'ESE'), ('eg', 'EG')]

    division_stats = []

    # 4. Generate 13 Division Sheets
    for div_num in range(1, 14):
        div_code = f"FY{div_num}"
        div_students = pdf_df[pdf_df['Division'] == div_code].sort_values(by='Roll No').copy()
        sheet_title = f"Div {div_num} ({div_code})"
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True

        subjs = div_1_7_subjs if div_num <= 7 else div_8_13_subjs

        # Headers
        ws.merge_cells("A1:K1")
        c1 = ws["A1"]
        c1.value = "SCTR's PUNE INSTITUTE OF COMPUTER TECHNOLOGY, PUNE - 411043"
        c1.font = font_title
        c1.fill = navy_header_fill
        c1.alignment = align_center
        ws.row_dimensions[1].height = 28

        ws.merge_cells("A2:K2")
        c2 = ws["A2"]
        subj_str = ", ".join([s[1] for s in subjs])
        c2.value = f"First Year B.Tech (A.Y. 2026-27) | Division {div_num} ({div_code}) Marks Out of 20 | Subjects: {subj_str}"
        c2.font = font_subtitle
        c2.fill = navy_accent_fill
        c2.alignment = align_center
        ws.row_dimensions[2].height = 22

        # Table Column Headers
        headers = [
            "Sr. No.",
            "Roll No.",
            "MIS ID",
            "Student Name",
            f"{subjs[0][1]} (/20)",
            f"{subjs[1][1]} (/20)",
            f"{subjs[2][1]} (/20)",
            f"{subjs[3][1]} (/20)",
            f"{subjs[4][1]} (/20)",
            "Total (/100)",
            "Percentage (%)"
        ]

        ws.row_dimensions[4].height = 25
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cell.font = font_table_hdr
            cell.fill = table_header_fill_1
            cell.alignment = align_center
            cell.border = cell_border

        # Populate Student Rows
        start_row = 5
        cur_row = start_row

        for sr_no, (_, stu) in enumerate(div_students.iterrows(), 1):
            mis = stu['MIS ID']
            roll = stu['Roll No']
            name = stu['Student Name']

            is_zebra = (sr_no % 2 == 0)
            row_fill = zebra_fill if is_zebra else None

            ws.row_dimensions[cur_row].height = 20

            # Sr No
            c = ws.cell(row=cur_row, column=1, value=sr_no)
            c.font = font_data
            c.alignment = align_center
            c.border = cell_border
            if row_fill: c.fill = row_fill

            # Roll No
            c = ws.cell(row=cur_row, column=2, value=roll)
            c.font = font_bold
            c.alignment = align_center
            c.border = cell_border
            if row_fill: c.fill = row_fill

            # MIS ID
            c = ws.cell(row=cur_row, column=3, value=mis)
            c.font = font_data
            c.alignment = align_center
            c.border = cell_border
            if row_fill: c.fill = row_fill

            # Student Name
            c = ws.cell(row=cur_row, column=4, value=name)
            c.font = font_data
            c.alignment = align_left
            c.border = cell_border
            if row_fill: c.fill = row_fill

            # 5 Subject Marks
            for s_idx, (code, _) in enumerate(subjs):
                col_num = 5 + s_idx
                mark_val = marks[code].get(mis, None)
                c = ws.cell(row=cur_row, column=col_num)
                c.border = cell_border
                if row_fill: c.fill = row_fill
                if mark_val is not None:
                    c.value = mark_val
                    c.number_format = '0"/20"'
                    c.font = font_data
                    c.alignment = align_center
                else:
                    c.value = "AB"
                    c.font = font_absent
                    c.alignment = align_center

            # Total Formula: SUM(E:I)
            c_tot = ws.cell(row=cur_row, column=10)
            c_tot.value = f"=SUM(E{cur_row}:I{cur_row})"
            c_tot.number_format = '0"/100"'
            c_tot.font = font_bold
            c_tot.alignment = align_center
            c_tot.border = cell_border
            c_tot.fill = total_col_fill

            # Percentage Formula: J / 100 formatted as percentage
            c_pct = ws.cell(row=cur_row, column=11)
            c_pct.value = f"=J{cur_row}/100"
            c_pct.number_format = '0.0%'
            c_pct.font = font_bold
            c_pct.alignment = align_center
            c_pct.border = cell_border
            if row_fill: c_pct.fill = row_fill

            cur_row += 1

        end_row = cur_row - 1

        # Summary Statistics Rows
        stat_configs = [
            ("Class Average", f"=AVERAGE", '0.0"/20"', '0.0"/100"', '0.0%'),
            ("Highest Marks", f"=MAX", '0"/20"', '0"/100"', '0.0%'),
            ("Lowest Marks (Present)", f"=MIN", '0"/20"', '0"/100"', '0.0%'),
            ("Count Present", f"=COUNT", '0', '0', '0')
        ]

        cur_row += 1
        for label, func, num_fmt_sub, num_fmt_tot, num_fmt_pct in stat_configs:
            ws.row_dimensions[cur_row].height = 20
            # Label
            ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=4)
            lbl_c = ws.cell(row=cur_row, column=1, value=label)
            lbl_c.font = font_stats
            lbl_c.alignment = align_right
            lbl_c.fill = stats_hdr_fill
            for col_i in range(1, 5):
                ws.cell(row=cur_row, column=col_i).border = cell_border
                ws.cell(row=cur_row, column=col_i).fill = stats_hdr_fill

            # Subject Stat formulas
            for s_idx in range(5):
                col_l = get_column_letter(5 + s_idx)
                stat_c = ws.cell(row=cur_row, column=5 + s_idx)
                stat_c.value = f"{func}({col_l}{start_row}:{col_l}{end_row})"
                stat_c.font = font_stats
                stat_c.alignment = align_center
                stat_c.number_format = num_fmt_sub
                stat_c.fill = stats_fill
                stat_c.border = cell_border

            # Total stat
            tot_l = get_column_letter(10)
            tot_c = ws.cell(row=cur_row, column=10)
            tot_c.value = f"{func}({tot_l}{start_row}:{tot_l}{end_row})"
            tot_c.font = font_stats
            tot_c.alignment = align_center
            tot_c.number_format = num_fmt_tot
            tot_c.fill = stats_fill
            tot_c.border = cell_border

            # Pct stat
            pct_l = get_column_letter(11)
            pct_c = ws.cell(row=cur_row, column=11)
            pct_c.value = f"{func}({pct_l}{start_row}:{pct_l}{end_row})"
            pct_c.font = font_stats
            pct_c.alignment = align_center
            pct_c.number_format = num_fmt_pct
            pct_c.fill = stats_fill
            pct_c.border = cell_border

            cur_row += 1

        # Adjust column widths
        col_widths = {
            1: 8,   # Sr No
            2: 12,  # Roll No
            3: 12,  # MIS ID
            4: 34,  # Student Name
            5: 14,  # Sub 1
            6: 14,  # Sub 2
            7: 14,  # Sub 3
            8: 14,  # Sub 4
            9: 14,  # Sub 5
            10: 16, # Total
            11: 16  # Pct
        }
        for col_idx, width in col_widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        # Freeze panes below header
        ws.freeze_panes = "A5"

        # Enable AutoFilter
        ws.auto_filter.ref = f"A4:K{end_row}"

        division_stats.append({
            'Division': f"Div {div_num} ({div_code})",
            'Count': len(div_students),
            'Min Roll': div_students['Roll No'].min(),
            'Max Roll': div_students['Roll No'].max(),
            'Subjects': subj_str,
            'Sheet': sheet_title
        })

    # 5. Create "Consolidated Master" Sheet containing all 13 tables stacked
    ws_master = wb.create_sheet(title="All Divisions (Master)")
    ws_master.views.sheetView[0].showGridLines = True

    ws_master.merge_cells("A1:L1")
    c1 = ws_master["A1"]
    c1.value = "SCTR's PUNE INSTITUTE OF COMPUTER TECHNOLOGY - CONSOLIDATED MARKS (ALL 13 DIVISIONS)"
    c1.font = font_title
    c1.fill = navy_header_fill
    c1.alignment = align_center
    ws_master.row_dimensions[1].height = 28

    master_headers = [
        "Division", "Sr. No.", "Roll No.", "MIS ID", "Student Name",
        "Sub 1 (/20)", "Sub 2 (/20)", "Sub 3 (/20)", "Sub 4 (/20)", "Sub 5 (/20)",
        "Total (/100)", "Percentage (%)"
    ]
    ws_master.row_dimensions[3].height = 25
    for col_i, h in enumerate(master_headers, 1):
        c = ws_master.cell(row=3, column=col_i, value=h)
        c.font = font_table_hdr
        c.fill = table_header_fill_1
        c.alignment = align_center
        c.border = cell_border

    m_row = 4
    for div_num in range(1, 14):
        div_code = f"FY{div_num}"
        div_students = pdf_df[pdf_df['Division'] == div_code].sort_values(by='Roll No')
        subjs = div_1_7_subjs if div_num <= 7 else div_8_13_subjs

        for sr_no, (_, stu) in enumerate(div_students.iterrows(), 1):
            mis = stu['MIS ID']
            roll = stu['Roll No']
            name = stu['Student Name']
            is_zebra = (m_row % 2 == 0)
            row_fill = zebra_fill if is_zebra else None

            ws_master.row_dimensions[m_row].height = 19

            c = ws_master.cell(row=m_row, column=1, value=div_code)
            c.font = font_bold; c.alignment = align_center; c.border = cell_border
            if row_fill: c.fill = row_fill

            c = ws_master.cell(row=m_row, column=2, value=sr_no)
            c.font = font_data; c.alignment = align_center; c.border = cell_border
            if row_fill: c.fill = row_fill

            c = ws_master.cell(row=m_row, column=3, value=roll)
            c.font = font_bold; c.alignment = align_center; c.border = cell_border
            if row_fill: c.fill = row_fill

            c = ws_master.cell(row=m_row, column=4, value=mis)
            c.font = font_data; c.alignment = align_center; c.border = cell_border
            if row_fill: c.fill = row_fill

            c = ws_master.cell(row=m_row, column=5, value=name)
            c.font = font_data; c.alignment = align_left; c.border = cell_border
            if row_fill: c.fill = row_fill

            # 5 marks
            for s_idx, (code, label) in enumerate(subjs):
                val = marks[code].get(mis, None)
                cell_m = ws_master.cell(row=m_row, column=6 + s_idx)
                cell_m.border = cell_border
                if row_fill: cell_m.fill = row_fill
                if val is not None:
                    cell_m.value = val
                    cell_m.number_format = '0"/20"'
                    cell_m.font = font_data
                    cell_m.alignment = align_center
                else:
                    cell_m.value = "AB"
                    cell_m.font = font_absent
                    cell_m.alignment = align_center

            # Total
            tot_c = ws_master.cell(row=m_row, column=11, value=f"=SUM(F{m_row}:J{m_row})")
            tot_c.number_format = '0"/100"'
            tot_c.font = font_bold
            tot_c.alignment = align_center
            tot_c.border = cell_border
            tot_c.fill = total_col_fill

            # Pct
            pct_c = ws_master.cell(row=m_row, column=12, value=f"=K{m_row}/100")
            pct_c.number_format = '0.0%'
            pct_c.font = font_bold
            pct_c.alignment = align_center
            pct_c.border = cell_border
            if row_fill: pct_c.fill = row_fill

            m_row += 1

    ws_master.freeze_panes = "A4"
    ws_master.auto_filter.ref = f"A3:L{m_row - 1}"
    m_widths = {1: 12, 2: 8, 3: 12, 4: 12, 5: 34, 6: 14, 7: 14, 8: 14, 9: 14, 10: 14, 11: 16, 12: 16}
    for col_idx, width in m_widths.items():
        ws_master.column_dimensions[get_column_letter(col_idx)].width = width

    # 6. Create "Unassigned (Late Admissions)" Sheet
    unassigned_students = [sc for sc in all_students_master if sc not in set(pdf_df['MIS ID'])]
    unassigned_students.sort(key=lambda sc: all_students_master[sc]['Student Name'])

    ws_unassigned = wb.create_sheet(title="Unassigned (Late Admissions)")
    ws_unassigned.views.sheetView[0].showGridLines = True

    ws_unassigned.merge_cells("A1:K1")
    c1 = ws_unassigned["A1"]
    c1.value = "STUDENTS PRESENT IN MARKS SHEETS BUT NOT IN ROLL NO LIST (V1_RollNo List-AY-2026-27)"
    c1.font = font_title
    c1.fill = navy_header_fill
    c1.alignment = align_center
    ws_unassigned.row_dimensions[1].height = 28

    ws_unassigned.merge_cells("A2:K2")
    c2 = ws_unassigned["A2"]
    c2.value = "These 70 students were admitted in later CAP rounds/institutional rounds. Group 1: Div 1-7 (MFR/QP/IEEE), Group 2: Div 8-13 (CST/ESE/EG)"
    c2.font = font_subtitle
    c2.fill = navy_accent_fill
    c2.alignment = align_center
    ws_unassigned.row_dimensions[2].height = 22

    un_headers = [
        "Sr. No.", "Student Code (MIS ID)", "Student Name", "Subject Group",
        "LAC (/20)", "CPPS (/20)", "Sub 3 (/20)", "Sub 4 (/20)", "Sub 5 (/20)",
        "Total (/100)", "Percentage (%)"
    ]
    ws_unassigned.row_dimensions[4].height = 25
    for col_i, h in enumerate(un_headers, 1):
        c = ws_unassigned.cell(row=4, column=col_i, value=h)
        c.font = font_table_hdr
        c.fill = table_header_fill_2
        c.alignment = align_center
        c.border = cell_border

    u_row = 5
    for sr_no, sc in enumerate(unassigned_students, 1):
        stu = all_students_master[sc]
        is_zebra = (u_row % 2 == 0)
        row_fill = zebra_fill if is_zebra else None

        # Check if in Div 1-7 or Div 8-13
        is_div_1_7 = (sc in marks['mfr'] or sc in marks['qp'] or sc in marks['ieee'])
        grp_label = "Div 1-7 (LAC, CPPS, MFR, QP, IEEE)" if is_div_1_7 else "Div 8-13 (LAC, CPPS, CST, ESE, EG)"
        subjs = div_1_7_subjs if is_div_1_7 else div_8_13_subjs

        ws_unassigned.row_dimensions[u_row].height = 20

        c = ws_unassigned.cell(row=u_row, column=1, value=sr_no)
        c.font = font_data; c.alignment = align_center; c.border = cell_border
        if row_fill: c.fill = row_fill

        c = ws_unassigned.cell(row=u_row, column=2, value=sc)
        c.font = font_bold; c.alignment = align_center; c.border = cell_border
        if row_fill: c.fill = row_fill

        c = ws_unassigned.cell(row=u_row, column=3, value=stu['Student Name'])
        c.font = font_data; c.alignment = align_left; c.border = cell_border
        if row_fill: c.fill = row_fill

        c = ws_unassigned.cell(row=u_row, column=4, value=grp_label)
        c.font = font_data; c.alignment = align_center; c.border = cell_border
        if row_fill: c.fill = row_fill

        # 5 marks
        for s_idx, (code, _) in enumerate(subjs):
            val = marks[code].get(sc, None)
            c = ws_unassigned.cell(row=u_row, column=5 + s_idx)
            c.border = cell_border
            if row_fill: c.fill = row_fill
            if val is not None:
                c.value = val
                c.number_format = '0"/20"'
                c.font = font_data
                c.alignment = align_center
            else:
                c.value = "AB"
                c.font = font_absent
                c.alignment = align_center

        # Total
        tot_c = ws_unassigned.cell(row=u_row, column=10, value=f"=SUM(E{u_row}:I{u_row})")
        tot_c.number_format = '0"/100"'
        tot_c.font = font_bold
        tot_c.alignment = align_center
        tot_c.border = cell_border
        tot_c.fill = total_col_fill

        # Pct
        pct_c = ws_unassigned.cell(row=u_row, column=11, value=f"=J{u_row}/100")
        pct_c.number_format = '0.0%'
        pct_c.font = font_bold
        pct_c.alignment = align_center
        pct_c.border = cell_border
        if row_fill: pct_c.fill = row_fill

        u_row += 1

    ws_unassigned.freeze_panes = "A5"
    ws_unassigned.auto_filter.ref = f"A4:K{u_row - 1}"
    u_widths = {1: 8, 2: 24, 3: 32, 4: 34, 5: 14, 6: 14, 7: 14, 8: 14, 9: 14, 10: 16, 11: 16}
    for col_idx, width in u_widths.items():
        ws_unassigned.column_dimensions[get_column_letter(col_idx)].width = width

    # 7. Create "Overview & Summary" Sheet
    ws_summary = wb.create_sheet(title="Overview & Summary", index=0)
    ws_summary.views.sheetView[0].showGridLines = True

    ws_summary.merge_cells("A1:G1")
    c1 = ws_summary["A1"]
    c1.value = "PUNE INSTITUTE OF COMPUTER TECHNOLOGY - F.Y. B.TECH MARKS SUMMARY (A.Y. 2026-27)"
    c1.font = font_title
    c1.fill = navy_header_fill
    c1.alignment = align_center
    ws_summary.row_dimensions[1].height = 28

    ws_summary.merge_cells("A2:G2")
    c2 = ws_summary["A2"]
    c2.value = "Continuous Assessment / Unit Test Evaluation Summary | 13 Divisions | 884 Assigned Students + 70 Later Admissions"
    c2.font = font_subtitle
    c2.fill = navy_accent_fill
    c2.alignment = align_center
    ws_summary.row_dimensions[2].height = 22

    sum_headers = [
        "Division", "Division Name", "Total Students", "Roll Number Range",
        "Curriculum Subjects (Total Marks /20)", "Division Sheet Link"
    ]
    ws_summary.row_dimensions[4].height = 25
    for col_i, h in enumerate(sum_headers, 1):
        c = ws_summary.cell(row=4, column=col_i, value=h)
        c.font = font_table_hdr
        c.fill = table_header_fill_1
        c.alignment = align_center
        c.border = cell_border

    s_row = 5
    for stat in division_stats:
        ws_summary.row_dimensions[s_row].height = 20
        is_zebra = (s_row % 2 == 0)
        row_fill = zebra_fill if is_zebra else None

        c = ws_summary.cell(row=s_row, column=1, value=stat['Division'].split(' ')[1])
        c.font = font_bold; c.alignment = align_center; c.border = cell_border
        if row_fill: c.fill = row_fill

        c = ws_summary.cell(row=s_row, column=2, value=stat['Division'])
        c.font = font_data; c.alignment = align_center; c.border = cell_border
        if row_fill: c.fill = row_fill

        c = ws_summary.cell(row=s_row, column=3, value=stat['Count'])
        c.font = font_bold; c.alignment = align_center; c.border = cell_border
        if row_fill: c.fill = row_fill

        c = ws_summary.cell(row=s_row, column=4, value=f"{stat['Min Roll']} - {stat['Max Roll']}")
        c.font = font_data; c.alignment = align_center; c.border = cell_border
        if row_fill: c.fill = row_fill

        c = ws_summary.cell(row=s_row, column=5, value=stat['Subjects'])
        c.font = font_data; c.alignment = align_left; c.border = cell_border
        if row_fill: c.fill = row_fill

        # Hyperlink to sheet
        link_c = ws_summary.cell(row=s_row, column=6, value=f"Open {stat['Sheet']}")
        link_c.font = Font(name="Segoe UI", size=10, bold=True, color="1F4E78", underline="single")
        link_c.alignment = align_center
        link_c.border = cell_border
        link_c.hyperlink = f"#'{stat['Sheet']}'!A1"
        if row_fill: link_c.fill = row_fill

        s_row += 1

    # Total row in summary
    ws_summary.row_dimensions[s_row].height = 22
    ws_summary.merge_cells(f"A{s_row}:B{s_row}")
    c = ws_summary.cell(row=s_row, column=1, value="TOTAL ASSIGNED STUDENTS")
    c.font = font_bold; c.alignment = align_right; c.fill = stats_hdr_fill
    ws_summary.cell(row=s_row, column=2).fill = stats_hdr_fill
    ws_summary.cell(row=s_row, column=1).border = cell_border
    ws_summary.cell(row=s_row, column=2).border = cell_border

    c_tot = ws_summary.cell(row=s_row, column=3, value=f"=SUM(C5:C{s_row - 1})")
    c_tot.font = font_bold; c_tot.alignment = align_center; c_tot.fill = stats_hdr_fill; c_tot.border = cell_border

    for col_i in range(4, 7):
        c = ws_summary.cell(row=s_row, column=col_i, value="")
        c.fill = stats_hdr_fill; c.border = cell_border

    s_row += 2
    ws_summary.merge_cells(f"A{s_row}:F{s_row}")
    c_extra = ws_summary.cell(row=s_row, column=1, value=f"+ 70 Later Admissions documented in 'Unassigned (Late Admissions)' sheet. Total Records Processed = 954.")
    c_extra.font = font_stats; c_extra.alignment = align_left

    s_widths = {1: 12, 2: 18, 3: 16, 4: 22, 5: 38, 6: 22}
    for col_idx, width in s_widths.items():
        ws_summary.column_dimensions[get_column_letter(col_idx)].width = width

    # Remove the default sheet that was created with the workbook
    wb.remove(default_sheet)

    # Save
    wb.save(output_filename)
    print(f"Workbook successfully generated and saved to: {output_filename}")

if __name__ == "__main__":
    build_marks_workbook()
