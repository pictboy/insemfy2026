import os
import re
import pdfplumber
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_complete_workbook():
    base_dir = r"d:\apictnaly"
    output_filename = os.path.join(base_dir, "PICT_Student_Marks_Division_Wise.xlsx")
    pdf_v1_path = os.path.join(base_dir, "V1_RollNo List-AY-2026-27_260916_185335.pdf")
    pdf_nri_path = os.path.join(base_dir, "Roll call for NRI_OCI_CIWGC_260915_223640.pdf")

    # 1. Parse V1 PDF (884 students)
    pattern_v1 = re.compile(r'^(\d+)\s+([A-Z0-9]+)\s+(FY\d+)\s+(.+)$')
    all_students = []

    with pdfplumber.open(pdf_v1_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            for line in text.split('\n'):
                m = pattern_v1.match(line.strip())
                if m:
                    roll, mis, div, name = m.groups()
                    all_students.append({
                        'Roll No': int(roll),
                        'MIS ID': mis.strip(),
                        'Division': div.strip(),
                        'Student Name': name.strip(),
                        'Category': 'General'
                    })

    print(f"Loaded {len(all_students)} regular students from V1 list.")

    # 2. Parse Subject Marks Files
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
    excel_student_info = {}

    for code, (fname, display_name) in excel_files.items():
        fpath = os.path.join(base_dir, fname)
        df = pd.read_excel(fpath)
        for _, row in df.iterrows():
            sc = str(row['Student Code']).strip()
            tot = str(row['Total Marks']).strip()
            val = int(tot.split('/')[0]) if '/' in tot else int(tot)
            marks[code][sc] = val
            if sc not in excel_student_info:
                excel_student_info[sc] = str(row['Student Name']).strip()

    # 3. Verified NRI Student Mapping
    # (Roll No -> (MIS ID, Canonical Name))
    nri_to_mis = {
        10172: ('F260969', 'Sawant Aarush M.'),
        10173: ('F260970', 'Waghela Aarya M.'),
        10174: ('F260971', 'Patil Aayush V.'),
        10268: ('F260972', 'Gangan Aishwarya N.'),
        10269: ('F260973', 'Faizpurkar Adi A.'),
        10270: ('F260975', 'Pujari Ameya J.'),
        10271: ('F260976', 'Inamdar Amey A.'),
        10272: ('F261036', 'Patidar Ashish M.'),
        10273: ('F260978', 'Phad Arnav N.'),
        10274: ('F260980', 'Choughule Adi Y.'),
        10368: ('F260981', 'Pranjale Aryan -.'),
        10369: ('F260974', 'Patil Atharva C.'),
        10370: ('F261037', 'Kadam Devansh A.'),
        10371: ('F260979', 'Shaikh Mohammed Rafaan R.'),
        10372: ('F260982', 'Fiske Shrikant S.'),
        10373: ('F260985', 'Lokhande Samihan S.'),
        10374: ('F260986', 'Raut Gargi K.'),
        10466: ('F260987', 'Gupta Harshal -.'),
        10467: ('F260989', 'Shingate Harsh V.'),
        10468: ('F260991', 'Mahajan Aaryan H.'),
        10469: ('F261040', 'Bholu Hussain F.'),
        10470: ('F260995', 'Jadhav Ishan U.'),
        10471: ('F260997', 'Jagtap Ishan N.'),
        10472: ('F260988', 'Yadwad Rutvi R.'),
        10473: ('F260999', 'Pachori Mokshat -.'),
        10474: ('F261001', 'Jakhotiya Priyanka A.'),
        10568: ('F260998', 'Walchale Drishti V.'),
        10569: ('F261002', 'Pujari Kaushal K.'),
        10570: ('F260983', 'Patidar Kavya D.'),
        10571: ('F261035', 'Kolhe Vidhi H.'),
        10572: ('F261003', 'Kothari Rashi S.'),
        10573: ('F261006', 'Mahamuni Priyal P.'),
        10574: ('F261005', 'Sharma Soumya S.'),
        10669: ('F261007', 'Bhaila Mustafa -.'),
        10670: ('F261009', 'Yadav Nikhil Kumar J.'),
        10671: ('F261011', 'Kasar Om D.'),
        10672: ('F260990', 'Chavan Sharvari U.'),
        10673: ('F261008', 'Baderia Aditya -.'),
        10772: ('F261014', 'Pawar Om K.'),
        10773: ('F261015', 'Pethe Piyush S.'),
        10774: ('F261028', 'Khan Abdul Rahman F.'),
        10868: ('F260994', 'Dhunware Rayyan R.'),
        10869: ('F261034', 'Shaikh Aryaan R.'),
        10870: ('F261017', 'Patil Eva R.'),
        10871: ('F261027', 'Chipolkar Rumana Y.'),
        10872: ('F261022', 'Jamadar Saad I.'),
        10873: ('F261029', 'Malgundkar Sahil S.'),
        10968: ('F261023', 'Karade Eashan S.'),
        10969: ('F260993', 'Raj Sahil -.'),
        10970: ('F260992', 'Yadav Samarth S.'),
        10971: ('F261018', 'Sarnaik Shaunak S.'),
        10972: ('F261019', 'Neog Samudra S.'),
        10973: ('NRI-10973', 'Shetty Parinith'),
        11067: ('F261013', 'Mulay Shreeya N.'),
        11068: ('NRI-11068', 'Patil Shreyas Dhanraj'),
        11069: ('F261021', 'Apte Shubhan R.'),
        11070: ('F261030', 'Siddarth Ammanath'),
        11071: ('F261004', 'Patil Siddhi P.'),
        11072: ('F261026', 'Chavan Srushti S.'),
        11073: ('F261020', 'Rudagi Aadi S.'),
        11169: ('F261000', 'Shilimkar Sonal S.'),
        11170: ('F261016', 'Kunnummal Veedu Devamithra -.'),
        11171: ('F260977', 'Palikan Suwan -.'),
        11172: ('F261033', 'Bhavar Suyash B.'),
        11173: ('NRI-11173', 'Patil Tanushree Nilesh'),
        11271: ('F261025', 'Sharma Torsha -.'),
        11272: ('F260996', 'Cerejo Blesiya A.'),
        11273: ('F261038', 'Verma Vyom V.'),
        11371: ('F261012', 'Raghani Vidhi H.'),
        11372: ('F261031', 'Vikhe Shardool R.'),
        11373: ('F260984', 'Kashyap Palvi V.')
    }

    pattern_nri = re.compile(r'^(\d+)\s+(FY\d+)\s+(.+)$')
    with pdfplumber.open(pdf_nri_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            for line in text.split('\n'):
                m = pattern_nri.match(line.strip())
                if m:
                    roll = int(m.group(1))
                    div = m.group(2).strip()
                    raw_name = m.group(3).strip()
                    mis, canon_name = nri_to_mis.get(roll, (f"NRI-{roll}", raw_name))
                    all_students.append({
                        'Roll No': roll,
                        'MIS ID': mis,
                        'Division': div,
                        'Student Name': canon_name,
                        'Category': 'NRI / OCI / CIWGC'
                    })

    print(f"Total students combined (V1 + NRI): {len(all_students)}")
    students_df = pd.DataFrame(all_students)

    # 4. Initialize Workbook & Styles
    wb = openpyxl.Workbook()
    default_sheet = wb.active

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

    div_1_7_subjs = [('lac', 'LAC'), ('cpps', 'CPPS'), ('mfr', 'MFR'), ('qp', 'QP'), ('ieee', 'IEEE')]
    div_8_13_subjs = [('lac', 'LAC'), ('cpps', 'CPPS'), ('cst', 'CST'), ('ese', 'ESE'), ('eg', 'EG')]

    division_stats = []

    # 5. Build 13 Division Sheets
    for div_num in range(1, 14):
        div_code = f"FY{div_num}"
        div_students = students_df[students_df['Division'] == div_code].sort_values(by='Roll No').copy()
        sheet_title = f"Div {div_num} ({div_code})"
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True

        subjs = div_1_7_subjs if div_num <= 7 else div_8_13_subjs
        subj_str = ", ".join([s[1] for s in subjs])

        # Title Banners
        ws.merge_cells("A1:K1")
        c1 = ws["A1"]
        c1.value = "SCTR's PUNE INSTITUTE OF COMPUTER TECHNOLOGY, PUNE - 411043"
        c1.font = font_title
        c1.fill = navy_header_fill
        c1.alignment = align_center
        ws.row_dimensions[1].height = 28

        ws.merge_cells("A2:K2")
        c2 = ws["A2"]
        c2.value = f"First Year B.Tech (A.Y. 2026-27) | Division {div_num} ({div_code}) Marks Out of 20 | Subjects: {subj_str}"
        c2.font = font_subtitle
        c2.fill = navy_accent_fill
        c2.alignment = align_center
        ws.row_dimensions[2].height = 22

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
            c.font = font_data; c.alignment = align_center; c.border = cell_border
            if row_fill: c.fill = row_fill

            # Roll No
            c = ws.cell(row=cur_row, column=2, value=roll)
            c.font = font_bold; c.alignment = align_center; c.border = cell_border
            if row_fill: c.fill = row_fill

            # MIS ID
            c = ws.cell(row=cur_row, column=3, value=mis)
            c.font = font_data; c.alignment = align_center; c.border = cell_border
            if row_fill: c.fill = row_fill

            # Student Name
            c = ws.cell(row=cur_row, column=4, value=name)
            c.font = font_data; c.alignment = align_left; c.border = cell_border
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

            # Total
            c_tot = ws.cell(row=cur_row, column=10)
            c_tot.value = f"=SUM(E{cur_row}:I{cur_row})"
            c_tot.number_format = '0"/100"'
            c_tot.font = font_bold
            c_tot.alignment = align_center
            c_tot.border = cell_border
            c_tot.fill = total_col_fill

            # Percentage
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
            ("Class Average", "=AVERAGE", '0.0"/20"', '0.0"/100"', '0.0%'),
            ("Highest Marks", "=MAX", '0"/20"', '0"/100"', '0.0%'),
            ("Lowest Marks (Present)", "=MIN", '0"/20"', '0"/100"', '0.0%'),
            ("Count Present", "=COUNT", '0', '0', '0')
        ]

        cur_row += 1
        for label, func, num_fmt_sub, num_fmt_tot, num_fmt_pct in stat_configs:
            ws.row_dimensions[cur_row].height = 20
            ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=4)
            lbl_c = ws.cell(row=cur_row, column=1, value=label)
            lbl_c.font = font_stats; lbl_c.alignment = align_right; lbl_c.fill = stats_hdr_fill
            for col_i in range(1, 5):
                ws.cell(row=cur_row, column=col_i).border = cell_border
                ws.cell(row=cur_row, column=col_i).fill = stats_hdr_fill

            for s_idx in range(5):
                col_l = get_column_letter(5 + s_idx)
                stat_c = ws.cell(row=cur_row, column=5 + s_idx)
                stat_c.value = f"{func}({col_l}{start_row}:{col_l}{end_row})"
                stat_c.font = font_stats; stat_c.alignment = align_center
                stat_c.number_format = num_fmt_sub; stat_c.fill = stats_fill; stat_c.border = cell_border

            tot_l = get_column_letter(10)
            tot_c = ws.cell(row=cur_row, column=10)
            tot_c.value = f"{func}({tot_l}{start_row}:{tot_l}{end_row})"
            tot_c.font = font_stats; tot_c.alignment = align_center
            tot_c.number_format = num_fmt_tot; tot_c.fill = stats_fill; tot_c.border = cell_border

            pct_l = get_column_letter(11)
            pct_c = ws.cell(row=cur_row, column=11)
            pct_c.value = f"{func}({pct_l}{start_row}:{pct_l}{end_row})"
            pct_c.font = font_stats; pct_c.alignment = align_center
            pct_c.number_format = num_fmt_pct; pct_c.fill = stats_fill; pct_c.border = cell_border

            cur_row += 1

        col_widths = {1: 8, 2: 12, 3: 14, 4: 34, 5: 14, 6: 14, 7: 14, 8: 14, 9: 14, 10: 16, 11: 16}
        for col_idx, width in col_widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        ws.freeze_panes = "A5"
        ws.auto_filter.ref = f"A4:K{end_row}"

        nri_count_in_div = len(div_students[div_students['Category'] != 'General'])
        division_stats.append({
            'Division': f"Div {div_num} ({div_code})",
            'Total Count': len(div_students),
            'General Count': len(div_students) - nri_count_in_div,
            'NRI Count': nri_count_in_div,
            'Min Roll': div_students['Roll No'].min(),
            'Max Roll': div_students['Roll No'].max(),
            'Subjects': subj_str,
            'Sheet': sheet_title
        })

    # 6. Consolidated Master Sheet (955 students)
    ws_master = wb.create_sheet(title="All Divisions (Master)")
    ws_master.views.sheetView[0].showGridLines = True

    ws_master.merge_cells("A1:M1")
    c1 = ws_master["A1"]
    c1.value = "SCTR's PUNE INSTITUTE OF COMPUTER TECHNOLOGY - CONSOLIDATED MARKS (ALL 13 DIVISIONS INCL. NRI)"
    c1.font = font_title; c1.fill = navy_header_fill; c1.alignment = align_center
    ws_master.row_dimensions[1].height = 28

    master_headers = [
        "Division", "Sr. No.", "Roll No.", "MIS ID", "Student Name", "Category",
        "Sub 1 (/20)", "Sub 2 (/20)", "Sub 3 (/20)", "Sub 4 (/20)", "Sub 5 (/20)",
        "Total (/100)", "Percentage (%)"
    ]
    ws_master.row_dimensions[3].height = 25
    for col_i, h in enumerate(master_headers, 1):
        c = ws_master.cell(row=3, column=col_i, value=h)
        c.font = font_table_hdr; c.fill = table_header_fill_1; c.alignment = align_center; c.border = cell_border

    m_row = 4
    for div_num in range(1, 14):
        div_code = f"FY{div_num}"
        div_students = students_df[students_df['Division'] == div_code].sort_values(by='Roll No')
        subjs = div_1_7_subjs if div_num <= 7 else div_8_13_subjs

        for sr_no, (_, stu) in enumerate(div_students.iterrows(), 1):
            mis = stu['MIS ID']
            roll = stu['Roll No']
            name = stu['Student Name']
            cat = stu['Category']
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

            c = ws_master.cell(row=m_row, column=6, value=cat)
            c.font = font_data; c.alignment = align_center; c.border = cell_border
            if row_fill: c.fill = row_fill

            for s_idx, (code, _) in enumerate(subjs):
                val = marks[code].get(mis, None)
                cell_m = ws_master.cell(row=m_row, column=7 + s_idx)
                cell_m.border = cell_border
                if row_fill: cell_m.fill = row_fill
                if val is not None:
                    cell_m.value = val; cell_m.number_format = '0"/20"'; cell_m.font = font_data; cell_m.alignment = align_center
                else:
                    cell_m.value = "AB"; cell_m.font = font_absent; cell_m.alignment = align_center

            tot_c = ws_master.cell(row=m_row, column=12, value=f"=SUM(G{m_row}:K{m_row})")
            tot_c.number_format = '0"/100"'; tot_c.font = font_bold; tot_c.alignment = align_center; tot_c.border = cell_border; tot_c.fill = total_col_fill

            pct_c = ws_master.cell(row=m_row, column=13, value=f"=L{m_row}/100")
            pct_c.number_format = '0.0%'; pct_c.font = font_bold; pct_c.alignment = align_center; pct_c.border = cell_border
            if row_fill: pct_c.fill = row_fill

            m_row += 1

    ws_master.freeze_panes = "A4"
    ws_master.auto_filter.ref = f"A3:M{m_row - 1}"
    m_widths = {1: 12, 2: 8, 3: 12, 4: 14, 5: 34, 6: 18, 7: 14, 8: 14, 9: 14, 10: 14, 11: 14, 12: 16, 13: 16}
    for col_idx, width in m_widths.items():
        ws_master.column_dimensions[get_column_letter(col_idx)].width = width

    # 7. Unassigned / Other Sheet (the 2 leftover students: Chimangaonkar and Mulla)
    allocated_mis = set(students_df['MIS ID'])
    leftovers = [sc for sc in excel_student_info if sc not in allocated_mis]
    leftovers.sort()

    ws_unassigned = wb.create_sheet(title="Other Admissions (Unassigned)")
    ws_unassigned.views.sheetView[0].showGridLines = True

    ws_unassigned.merge_cells("A1:K1")
    c1 = ws_unassigned["A1"]
    c1.value = "OTHER ADMISSIONS / IRREGULAR (NOT IN REGULAR OR NRI ROLL CALL)"
    c1.font = font_title; c1.fill = navy_header_fill; c1.alignment = align_center
    ws_unassigned.row_dimensions[1].height = 28

    un_headers = [
        "Sr. No.", "Student Code", "Student Name", "Curriculum Stream",
        "LAC (/20)", "CPPS (/20)", "Sub 3 (/20)", "Sub 4 (/20)", "Sub 5 (/20)",
        "Total (/100)", "Percentage (%)"
    ]
    ws_unassigned.row_dimensions[3].height = 25
    for col_i, h in enumerate(un_headers, 1):
        c = ws_unassigned.cell(row=3, column=col_i, value=h)
        c.font = font_table_hdr; c.fill = table_header_fill_2; c.alignment = align_center; c.border = cell_border

    u_row = 4
    for sr_no, sc in enumerate(leftovers, 1):
        s_name = excel_student_info[sc]
        is_div_1_7 = (sc in marks['mfr'] or sc in marks['qp'] or sc in marks['ieee'])
        grp_label = "Div 1-7 Group (LAC, CPPS, MFR, QP, IEEE)" if is_div_1_7 else "Div 8-13 Group (LAC, CPPS, CST, ESE, EG)"
        subjs = div_1_7_subjs if is_div_1_7 else div_8_13_subjs

        ws_unassigned.row_dimensions[u_row].height = 20

        c = ws_unassigned.cell(row=u_row, column=1, value=sr_no)
        c.font = font_data; c.alignment = align_center; c.border = cell_border

        c = ws_unassigned.cell(row=u_row, column=2, value=sc)
        c.font = font_bold; c.alignment = align_center; c.border = cell_border

        c = ws_unassigned.cell(row=u_row, column=3, value=s_name)
        c.font = font_data; c.alignment = align_left; c.border = cell_border

        c = ws_unassigned.cell(row=u_row, column=4, value=grp_label)
        c.font = font_data; c.alignment = align_center; c.border = cell_border

        for s_idx, (code, _) in enumerate(subjs):
            val = marks[code].get(sc, None)
            c = ws_unassigned.cell(row=u_row, column=5 + s_idx)
            c.border = cell_border
            if val is not None:
                c.value = val; c.number_format = '0"/20"'; c.font = font_data; c.alignment = align_center
            else:
                c.value = "AB"; c.font = font_absent; c.alignment = align_center

        tot_c = ws_unassigned.cell(row=u_row, column=10, value=f"=SUM(E{u_row}:I{u_row})")
        tot_c.number_format = '0"/100"'; tot_c.font = font_bold; tot_c.alignment = align_center; tot_c.border = cell_border; tot_c.fill = total_col_fill

        pct_c = ws_unassigned.cell(row=u_row, column=11, value=f"=J{u_row}/100")
        pct_c.number_format = '0.0%'; pct_c.font = font_bold; pct_c.alignment = align_center; pct_c.border = cell_border

        u_row += 1

    u_widths = {1: 8, 2: 18, 3: 32, 4: 38, 5: 14, 6: 14, 7: 14, 8: 14, 9: 14, 10: 16, 11: 16}
    for col_idx, width in u_widths.items():
        ws_unassigned.column_dimensions[get_column_letter(col_idx)].width = width

    # 8. Overview & Summary Sheet
    ws_summary = wb.create_sheet(title="Overview & Summary", index=0)
    ws_summary.views.sheetView[0].showGridLines = True

    ws_summary.merge_cells("A1:H1")
    c1 = ws_summary["A1"]
    c1.value = "PUNE INSTITUTE OF COMPUTER TECHNOLOGY - F.Y. B.TECH MARKS SUMMARY (A.Y. 2026-27)"
    c1.font = font_title; c1.fill = navy_header_fill; c1.alignment = align_center
    ws_summary.row_dimensions[1].height = 28

    ws_summary.merge_cells("A2:H2")
    c2 = ws_summary["A2"]
    c2.value = "Continuous Assessment / Unit Test Evaluation Summary | 13 Divisions | 955 Total Students (884 Regular + 71 NRI/OCI/CIWGC)"
    c2.font = font_subtitle; c2.fill = navy_accent_fill; c2.alignment = align_center
    ws_summary.row_dimensions[2].height = 22

    sum_headers = [
        "Div No.", "Division Name", "Total Students", "Regular Students", "NRI Students",
        "Roll No. Range", "Curriculum Subjects (/20)", "Division Sheet"
    ]
    ws_summary.row_dimensions[4].height = 25
    for col_i, h in enumerate(sum_headers, 1):
        c = ws_summary.cell(row=4, column=col_i, value=h)
        c.font = font_table_hdr; c.fill = table_header_fill_1; c.alignment = align_center; c.border = cell_border

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

        c = ws_summary.cell(row=s_row, column=3, value=stat['Total Count'])
        c.font = font_bold; c.alignment = align_center; c.border = cell_border
        if row_fill: c.fill = row_fill

        c = ws_summary.cell(row=s_row, column=4, value=stat['General Count'])
        c.font = font_data; c.alignment = align_center; c.border = cell_border
        if row_fill: c.fill = row_fill

        c = ws_summary.cell(row=s_row, column=5, value=stat['NRI Count'])
        c.font = font_bold; c.alignment = align_center; c.border = cell_border
        if row_fill: c.fill = row_fill

        c = ws_summary.cell(row=s_row, column=6, value=f"{stat['Min Roll']} - {stat['Max Roll']}")
        c.font = font_data; c.alignment = align_center; c.border = cell_border
        if row_fill: c.fill = row_fill

        c = ws_summary.cell(row=s_row, column=7, value=stat['Subjects'])
        c.font = font_data; c.alignment = align_left; c.border = cell_border
        if row_fill: c.fill = row_fill

        link_c = ws_summary.cell(row=s_row, column=8, value=f"Open {stat['Sheet']}")
        link_c.font = Font(name="Segoe UI", size=10, bold=True, color="1F4E78", underline="single")
        link_c.alignment = align_center; link_c.border = cell_border; link_c.hyperlink = f"#'{stat['Sheet']}'!A1"
        if row_fill: link_c.fill = row_fill

        s_row += 1

    # Total Row in Summary
    ws_summary.row_dimensions[s_row].height = 22
    ws_summary.merge_cells(f"A{s_row}:B{s_row}")
    c = ws_summary.cell(row=s_row, column=1, value="TOTAL ALLOCATED STUDENTS")
    c.font = font_bold; c.alignment = align_right; c.fill = stats_hdr_fill
    ws_summary.cell(row=s_row, column=2).fill = stats_hdr_fill
    ws_summary.cell(row=s_row, column=1).border = cell_border
    ws_summary.cell(row=s_row, column=2).border = cell_border

    # Sum Total
    c_tot = ws_summary.cell(row=s_row, column=3, value=f"=SUM(C5:C{s_row - 1})")
    c_tot.font = font_bold; c_tot.alignment = align_center; c_tot.fill = stats_hdr_fill; c_tot.border = cell_border

    # Sum General
    c_gen = ws_summary.cell(row=s_row, column=4, value=f"=SUM(D5:D{s_row - 1})")
    c_gen.font = font_bold; c_gen.alignment = align_center; c_gen.fill = stats_hdr_fill; c_gen.border = cell_border

    # Sum NRI
    c_nri = ws_summary.cell(row=s_row, column=5, value=f"=SUM(E5:E{s_row - 1})")
    c_nri.font = font_bold; c_nri.alignment = align_center; c_nri.fill = stats_hdr_fill; c_nri.border = cell_border

    for col_i in range(6, 9):
        c = ws_summary.cell(row=s_row, column=col_i, value="")
        c.fill = stats_hdr_fill; c.border = cell_border

    s_row += 2
    ws_summary.merge_cells(f"A{s_row}:H{s_row}")
    c_extra = ws_summary.cell(row=s_row, column=1, value="+ 2 irregular students recorded in 'Other Admissions (Unassigned)'. Total student records tracked = 957.")
    c_extra.font = font_stats; c_extra.alignment = align_left

    s_widths = {1: 10, 2: 18, 3: 16, 4: 16, 5: 14, 6: 22, 7: 38, 8: 20}
    for col_idx, width in s_widths.items():
        ws_summary.column_dimensions[get_column_letter(col_idx)].width = width

    wb.remove(default_sheet)
    try:
        wb.save(output_filename)
        print(f"Workbook successfully updated and saved to: {output_filename}")
    except PermissionError:
        fallback_filename = os.path.join(base_dir, "PICT_Student_Marks_Division_Wise_with_NRI.xlsx")
        wb.save(fallback_filename)
        print(f"Original file was open in Excel. Saved updated workbook to: {fallback_filename}")

if __name__ == "__main__":
    build_complete_workbook()
