import re

def apply():
    with open("app.py", "r", encoding="utf-8") as f:
        content = f.read()

    # 1. Helper function insertion
    helper_code = """                    else:
                        for col in range(1, 15):
                            sheet_ces.cell(row=r, column=col).value = None

            # Prevent division by zero
            for col in range(4, 14):
                if sheet_ces.cell(row=11, column=col).value is None:
                    sheet_ces.cell(row=11, column=col).value = 0

def extract_lab_marks_from_excel(file_stream, num_marks_cols):
    \"\"\"
    Parses a filled Lab Excel sheet.
    Returns:
        students_data: dict mapping USN -> {
            'name': str,
            'marks': list of floats/ints/None,
            'see_100': float/None,
            'grade': str/None
        }
    \"\"\"
    file_stream.seek(0)
    wb = load_workbook(io.BytesIO(file_stream.read()), data_only=True)
    sheet = None
    for name in wb.sheetnames:
        if 'lab' in name.lower():
            sheet = wb[name]
            break
    if not sheet:
        sheet = wb.active
        
    header_row_idx = None
    usn_col_idx = None
    name_col_idx = None
    
    for r in range(1, 30):
        row_vals = [sheet.cell(row=r, column=c).value for c in range(1, 25)]
        row_vals_str = [str(v).strip().upper() if v is not None else "" for v in row_vals]
        if "USN" in row_vals_str:
            header_row_idx = r
            usn_col_idx = row_vals_str.index("USN") + 1
            for idx, val in enumerate(row_vals_str):
                if "NAME" in val or "STUDENT" in val:
                    name_col_idx = idx + 1
                    break
            if not name_col_idx:
                name_col_idx = usn_col_idx + 1
            break
            
    if not header_row_idx:
        header_row_idx = 8
        usn_col_idx = 2
        name_col_idx = 3
        
    students_data = {}
    max_row = sheet.max_row
    for r in range(header_row_idx + 1, max_row + 1):
        usn = sheet.cell(row=r, column=usn_col_idx).value
        name = sheet.cell(row=r, column=name_col_idx).value
        
        if usn is not None:
            usn_str = str(usn).strip()
            if not usn_str or usn_str.upper() in ("USN", "NONE", "NAN") or len(usn_str) < 5:
                continue
            
            marks = []
            for m_col in range(4, 4 + num_marks_cols):
                val = sheet.cell(row=r, column=m_col).value
                if val is not None and str(val).strip() != "":
                    try:
                        marks.append(float(val))
                    except ValueError:
                        marks.append(None)
                else:
                    marks.append(None)
                    
            see_100 = None
            grade = None
            if num_marks_cols == 13:
                val_see = sheet.cell(row=r, column=19).value
                if val_see is not None and str(val_see).strip() != "":
                    try:
                        see_100 = float(val_see)
                    except ValueError:
                        pass
                val_grade = sheet.cell(row=r, column=20).value
                if val_grade is not None:
                    grade = str(val_grade).strip()
                        
            name_str = str(name).strip() if name is not None else ""
            students_data[usn_str.upper()] = {
                'name': name_str,
                'marks': marks,
                'see_100': see_100,
                'grade': grade
            }
            
    return students_data

# ----------------- CORE PROCESSING PIPELINE -----------------"""

    target_helper = """                    else:
                        for col in range(1, 15):
                            sheet_ces.cell(row=r, column=col).value = None

            # Prevent division by zero
            for col in range(4, 14):
                if sheet_ces.cell(row=11, column=col).value is None:
                    sheet_ces.cell(row=11, column=col).value = 0

# ----------------- CORE PROCESSING PIPELINE -----------------"""

    if target_helper not in content:
        print("Error: helper target not found!")
        return False
    content = content.replace(target_helper, helper_code)

    # 2. Modify process_stand_alone_lab function
    # Let's replace the whole process_stand_alone_lab up to its end.
    # Its start is:
    # def process_stand_alone_lab(template_file, rubrics_file, api_key, co_vals, po_vals, override_course_code=None, ces_file=None):
    # Its end is:
    #     for i in range(num_students):
    #         usn = sorted_usns[i]
    #         marks = extracted_marks_map[usn]
    #         r_lab = 9 + i
    #         for idx in range(12):
    #             col = 4 + idx
    #             if idx < len(marks):
    #                 val = marks[idx]
    #                 if val is not None and val != "":
    #                     try:
    #                         sheet.cell(row=r_lab, column=col).value = float(val)
    #                     except ValueError:
    #                         pass

    # Let's locate the entire body and replace it.
    pattern_lab = r"def process_stand_alone_lab\(template_file, rubrics_file, api_key, co_vals, po_vals, override_course_code=None, ces_file=None\):.*?# 8\. Populate and Map Course Exit Survey \(CES\)"
    
    new_lab_code = """def process_stand_alone_lab(template_file, lab_marks_file, co_vals, po_vals, override_course_code=None, ces_file=None):
    \"\"\"Fills the Stand Alone Lab sheet cleanly and handles dynamic rosters.\"\"\"
    import io
    
    wb = load_workbook(io.BytesIO(template_file.read()))
    sheet = wb['Lab'] if 'Lab' in wb.sheetnames else wb.active
    
    course_code = override_course_code
    if not course_code:
        # Try to read course code from sheet_theory cell row 5, col 1
        if 'Theory' in wb.sheetnames:
            cell_val = wb['Theory'].cell(row=5, column=1).value
            if cell_val and ":" in str(cell_val):
                course_code = str(cell_val).split(":")[-1].strip()
        
    # 1. Configure Course Type Cells in Theory sheet
    if 'Theory' in wb.sheetnames:
        sheet_theory = wb['Theory']
        sheet_theory.cell(row=7, column=4).value = "NO"  # STAND ALONE THEORY
        sheet_theory.cell(row=8, column=4).value = "YES" # STAND ALONE LAB
        sheet_theory.cell(row=9, column=4).value = "NO"  # IPCC
        if course_code:
            sheet_theory.cell(row=5, column=1).value = f"COURSE  CODE : {course_code}"
        
    # 2. Extract marks from filled Lab Excel file
    try:
        extracted_marks_map = extract_lab_marks_from_excel(lab_marks_file, num_marks_cols=13)
    except Exception as e:
        raise ValueError(f"Failed to read lab marks Excel sheet: {e}")

    # 3. Build Student Roster
    master_students = {}
    for usn, sdata in extracted_marks_map.items():
        master_students[usn] = sdata['name']
            
    sorted_usns = sorted(master_students.keys())
    num_students = len(sorted_usns)
    
    if num_students == 0:
        raise ValueError("No students extracted from the uploaded Lab Marks sheet.")

    # 4. Populate and Synchronize Roster to Lab sheet (Rows 9+) and Theory sheet (Rows 17 to 85)
    for i in range(num_students):
        usn = sorted_usns[i]
        name = master_students[usn]
        
        # Write to Lab sheet
        r_lab = 9 + i
        sheet.cell(row=r_lab, column=1).value = float(i + 1)
        sheet.cell(row=r_lab, column=2).value = usn
        sheet.cell(row=r_lab, column=3).value = name
        
        # Write to Theory sheet
        if 'Theory' in wb.sheetnames:
            sheet_theory = wb['Theory']
            r_theo = 17 + i
            sheet_theory.cell(row=r_theo, column=1).value = float(i + 1)
            sheet_theory.cell(row=r_theo, column=2).value = usn
            sheet_theory.cell(row=r_theo, column=3).value = name

    # Clear unused roster rows in Theory sheet (Rows 17 + N to 85)
    if 'Theory' in wb.sheetnames:
        sheet_theory = wb['Theory']
        for i in range(num_students, 85 - 17 + 1):
            r_theo = 17 + i
            for c in range(1, 123):
                sheet_theory.cell(row=r_theo, column=c).value = None

    # Clear student marks columns in Theory (columns D to CW, rows 17 to 17 + N - 1)
    if 'Theory' in wb.sheetnames:
        sheet_theory = wb['Theory']
        for r in range(17, 17 + num_students):
            for c in range(4, 102):
                sheet_theory.cell(row=r, column=c).value = None

    # 5. Fill COs and POs mapped in Lab sheet
    # Row 4 is POs Mapped, Row 5 is COs mapped. Columns D to O (4 to 15)
    for idx in range(12):
        col = 4 + idx
        if idx < len(po_vals) and po_vals[idx]:
            sheet.cell(row=4, column=col).value = po_vals[idx]
        if idx < len(co_vals) and co_vals[idx] is not None:
            sheet.cell(row=5, column=col).value = co_vals[idx]

    # Clear marks columns in Lab sheet (columns D to O, rows 9 to 9 + N - 1) before filling
    for r in range(9, 9 + num_students):
        for c in range(4, 16):
            sheet.cell(row=r, column=c).value = None

    # 6. Write Marks and Formulas to Lab Sheet
    for i in range(num_students):
        usn = sorted_usns[i]
        sdata = extracted_marks_map[usn]
        marks = sdata['marks']
        r_lab = 9 + i
        
        # Columns D to O (4 to 15) -> Lab 1 to 12
        for idx in range(12):
            col = 4 + idx
            if idx < len(marks):
                val = marks[idx]
                if val is not None and val != "":
                    try:
                        sheet.cell(row=r_lab, column=col).value = float(val)
                    except ValueError:
                        pass
                        
        # Column P (16) -> Component-2 (Test)
        if len(marks) > 12:
            val_p = marks[12]
            if val_p is not None and val_p != "":
                try:
                    sheet.cell(row=r_lab, column=16).value = float(val_p)
                except ValueError:
                    pass
                    
        # Write formulas for CIE (Q/17) and SEE-50 (R/18)
        sheet.cell(row=r_lab, column=17).value = f"=SUM(AVERAGE(D{r_lab}:O{r_lab}),P{r_lab})"
        sheet.cell(row=r_lab, column=18).value = f"=ROUND(S{r_lab}/2,0)"
        
        # Write S (19) and T (20) if present
        if sdata['see_100'] is not None:
            sheet.cell(row=r_lab, column=19).value = sdata['see_100']
        if sdata['grade'] is not None:
            sheet.cell(row=r_lab, column=20).value = sdata['grade']

    # 8. Populate and Map Course Exit Survey (CES)"""

    content, count = re.subn(pattern_lab, new_lab_code, content, flags=re.DOTALL)
    print(f"Substituted process_stand_alone_lab matches: {count}")
    if count == 0:
        print("Error: process_stand_alone_lab pattern not matched!")
        return False

    # 3. Modify process_ipcc_course function
    # Target from def process_ipcc_course up to sorted_usns/ValueError:
    pattern_ipcc_start = r"def process_ipcc_course\(template_file, qp_files, marks_files, quiz_file, aat_file, rubrics_file, api_key, co_vals, po_vals, override_course_code=None, ces_file=None\):.*?raise ValueError\(\"No students found in Theory files or Lab VLM extraction\.\"\)"
    new_ipcc_start = """def process_ipcc_course(template_file, qp_files, marks_files, quiz_file, aat_file, lab_marks_file, co_vals, po_vals, override_course_code=None, ces_file=None):
    \"\"\"
    Consolidates both Theory assessments and Lab Marks from filled Excel
    into the IPCC template workbook (containing both Theory and Lab sheets).
    \"\"\"
    import io

    wb = load_workbook(io.BytesIO(template_file.read()))
    sheet_theory = wb['Theory'] if 'Theory' in wb.sheetnames else wb.active
    sheet_lab = wb['Lab'] if 'Lab' in wb.sheetnames else wb.active
    
    # 1. Configure Course Type Cells in Theory sheet
    sheet_theory.cell(row=7, column=4).value = "NO"  # STAND ALONE THEORY
    sheet_theory.cell(row=8, column=4).value = "NO"  # STAND ALONE LAB
    sheet_theory.cell(row=9, column=4).value = "YES" # IPCC
    
    # 2. Ingest Lab Marks from filled Excel if provided (6 labs)
    extracted_marks_map = {}
    course_code = override_course_code
    
    if lab_marks_file:
        try:
            extracted_marks_map = extract_lab_marks_from_excel(lab_marks_file, num_marks_cols=6)
        except Exception as e:
            raise ValueError(f"Failed to read lab marks Excel sheet: {e}")

    # 3. Build Roster from both Theory & Lab Marks Ingestion
    master_students = {}
    
    # A. Load students from Theory marks files
    all_marks_files = []
    if marks_files:
        all_marks_files.extend([f for f in marks_files.values() if f])
    if quiz_file:
        all_marks_files.append(quiz_file)
    if aat_file:
        all_marks_files.append(aat_file)
        
    for mf in all_marks_files:
        mf.seek(0)
        df = pd.read_excel(io.BytesIO(mf.read()), engine='xlrd', header=None)
        mf.seek(0)
        header_row_idx = None
        usn_col = None
        for idx, row in df.iterrows():
            row_vals = [str(x).strip() for x in row.values]
            if 'USN' in row_vals:
                header_row_idx = idx
                usn_col = row_vals.index('USN')
                break
        if header_row_idx is not None and usn_col is not None:
            row_vals = [str(x).strip() for x in df.iloc[header_row_idx].values]
            name_col = None
            for c_idx, val in enumerate(row_vals):
                if 'student' in val.lower() or 'name' in val.lower():
                    if 'staff' not in val.lower() and 'division' not in val.lower():
                        name_col = c_idx
                        break
            if name_col is None:
                name_col = usn_col + 1
            for idx in range(header_row_idx + 1, len(df)):
                row = df.iloc[idx]
                u = str(row.iloc[usn_col]).strip()
                n = str(row.iloc[name_col]).strip() if name_col < len(row) else ""
                if u and u.lower() not in ('nan', 'none', '') and u.isalnum():
                    master_students[u] = n

    # B. Load students from Lab Excel file
    for usn, sdata in extracted_marks_map.items():
        if usn not in master_students or not master_students[usn]:
            master_students[usn] = sdata['name']

    sorted_usns = sorted(master_students.keys())
    num_students = len(sorted_usns)
    
    if num_students == 0:
        raise ValueError("No students found in Theory files or Lab Marks sheet.")"""

    content, count = re.subn(pattern_ipcc_start, new_ipcc_start, content, flags=re.DOTALL)
    print(f"Substituted process_ipcc_course start matches: {count}")
    if count == 0:
        print("Error: process_ipcc_course start pattern not matched!")
        return False

    # 4. Modify process_ipcc_course lab marks writing section
    target_ipcc_marks = """    # Write VLM extracted marks
    for i in range(num_students):
        usn = sorted_usns[i]
        r_lab = 9 + i
        if usn in extracted_marks_map:
            marks = extracted_marks_map[usn]
            for idx in range(6):
                col = 4 + idx
                if idx < len(marks):
                    val = marks[idx]
                    if val is not None and val != "":
                        try:
                            sheet_lab.cell(row=r_lab, column=col).value = float(val)
                        except ValueError:
                            pass"""

    new_ipcc_marks = """    # Write lab marks
    for i in range(num_students):
        usn = sorted_usns[i]
        r_lab = 9 + i
        if usn in extracted_marks_map:
            marks = extracted_marks_map[usn]['marks']
            for idx in range(6):
                col = 4 + idx
                if idx < len(marks):
                    val = marks[idx]
                    if val is not None and val != "":
                        try:
                            sheet_lab.cell(row=r_lab, column=col).value = float(val)
                        except ValueError:
                            pass"""

    if target_ipcc_marks not in content:
        print("Error: target_ipcc_marks not found in app.py!")
        return False
    content = content.replace(target_ipcc_marks, new_ipcc_marks)

    # 5. UI: tab layout for IPCC Course lab uploader and remove API key block
    target_ipcc_ui = """    # Lab defaults (IPCC)
    co_input = "1, 1, 1, 2, 2, 2"
    po_input = "1,5; 1,5; 1,5; 2,5; 2,5; 2,5"
    rubrics_file = None
    user_api_key = ""

    with tabs[0]:
        col1, col2 = st.columns(2)
        with col1:
            qp_files[1] = st.file_uploader("CIA-1 Question Paper (.docx)", type=["docx"], key="qp1")
        with col2:
            marks_files[1] = st.file_uploader("CIA-1 IA Marks (.xls)", type=["xls"], key="m1")

    with tabs[1]:
        col1, col2 = st.columns(2)
        with col1:
            qp_files[2] = st.file_uploader("CIA-2 Question Paper (.docx)", type=["docx"], key="qp2")
        with col2:
            marks_files[2] = st.file_uploader("CIA-2 IA Marks (.xls)", type=["xls"], key="m2")

    with tabs[2]:
        col1, col2 = st.columns(2)
        with col1:
            qp_files[3] = st.file_uploader("CIA-3 Question Paper (.docx)", type=["docx"], key="qp3")
        with col2:
            marks_files[3] = st.file_uploader("CIA-3 IA Marks (.xls)", type=["xls"], key="m3")

    with tabs[3]:
        col1, col2 = st.columns(2)
        with col1:
            quiz_file = st.file_uploader("Quiz Marks (.xls)", type=["xls"], key="quiz")
        with col2:
            aat_file = st.file_uploader("AAT Marks (.xls)", type=["xls"], key="aat")
            
    if course_type == "IPCC Course":
        with tabs[4]:
            st.subheader("Configure Lab Mapping Parameters")
            col_co, col_po = st.columns(2)
            with col_co:
                co_input = st.text_input(
                    "Lab CO Mappings (Row 5)", 
                    value="1, 1, 1, 2, 2, 2",
                    help="Enter exactly 6 CO values separated by commas for columns D to I."
                )
            with col_po:
                po_input = st.text_input(
                    "Lab PO Mappings (Row 4)", 
                    value="1,5; 1,5; 1,5; 2,5; 2,5; 2,5",
                    help="Enter PO mappings for each lab. Separate labs with a semicolon (;) and PO numbers with commas (,)."
                )
                
            st.subheader("Upload Rubrics Scanned PDF")
            rubrics_file = st.file_uploader(
                "Upload Scanned Lab Rubrics PDF (.pdf)",
                type=["pdf"],
                help="Scanned PDF showing student marks for all 6 labs."
            )
            
            st.subheader("Gemini Vision API Configuration")
            env_key = os.environ.get("GEMINI_API_KEY") or ""
            user_api_key = st.text_input(
                "Gemini API Key Override (Optional)",
                value="",
                type="password",
                help="If not set in .env, paste your key here. Leave blank to use the key configured in the .env file."
            )"""

    new_ipcc_ui = """    # Lab defaults (IPCC)
    co_input = "1, 1, 1, 2, 2, 2"
    po_input = "1,5; 1,5; 1,5; 2,5; 2,5; 2,5"
    lab_marks_file = None

    with tabs[0]:
        col1, col2 = st.columns(2)
        with col1:
            qp_files[1] = st.file_uploader("CIA-1 Question Paper (.docx)", type=["docx"], key="qp1")
        with col2:
            marks_files[1] = st.file_uploader("CIA-1 IA Marks (.xls)", type=["xls"], key="m1")

    with tabs[1]:
        col1, col2 = st.columns(2)
        with col1:
            qp_files[2] = st.file_uploader("CIA-2 Question Paper (.docx)", type=["docx"], key="qp2")
        with col2:
            marks_files[2] = st.file_uploader("CIA-2 IA Marks (.xls)", type=["xls"], key="m2")

    with tabs[2]:
        col1, col2 = st.columns(2)
        with col1:
            qp_files[3] = st.file_uploader("CIA-3 Question Paper (.docx)", type=["docx"], key="qp3")
        with col2:
            marks_files[3] = st.file_uploader("CIA-3 IA Marks (.xls)", type=["xls"], key="m3")

    with tabs[3]:
        col1, col2 = st.columns(2)
        with col1:
            quiz_file = st.file_uploader("Quiz Marks (.xls)", type=["xls"], key="quiz")
        with col2:
            aat_file = st.file_uploader("AAT Marks (.xls)", type=["xls"], key="aat")
            
    if course_type == "IPCC Course":
        with tabs[4]:
            st.subheader("Configure Lab Mapping Parameters")
            col_co, col_po = st.columns(2)
            with col_co:
                co_input = st.text_input(
                    "Lab CO Mappings (Row 5)", 
                    value="1, 1, 1, 2, 2, 2",
                    help="Enter exactly 6 CO values separated by commas for columns D to I."
                )
            with col_po:
                po_input = st.text_input(
                    "Lab PO Mappings (Row 4)", 
                    value="1,5; 1,5; 1,5; 2,5; 2,5; 2,5",
                    help="Enter PO mappings for each lab. Separate labs with a semicolon (;) and PO numbers with commas (,)."
                )
                
            st.subheader("Upload Filled Lab Marks Sheet")
            lab_marks_file = st.file_uploader(
                "Upload Filled Lab Marks Sheet (.xlsx)",
                type=["xlsx"],
                help="Provide the filled IPCC Lab Marks Excel sheet containing student marks."
            )"""

    if target_ipcc_ui not in content:
        print("Error: target_ipcc_ui not found in app.py!")
        return False
    content = content.replace(target_ipcc_ui, new_ipcc_ui)

    # 6. UI: uploader for Lab Course and remove API key block
    target_lab_ui = """    st.subheader("Upload Rubrics Scanned PDF")
    rubrics_file = st.file_uploader(
        "Upload Scanned Lab Rubrics PDF (.pdf)",
        type=["pdf"],
        help="Scanned PDF showing student marks for all 12 labs."
    )
    
    st.subheader("Upload Course Exit Survey Data")
    ces_file = st.file_uploader(
        "Upload Course Exit Survey (.xlsx)", 
        type=["xlsx"], 
        key="lab_ces",
        help="Optional Course Exit Survey Excel file."
    )
    
    st.subheader("Gemini Vision API Configuration")
    env_key = os.environ.get("GEMINI_API_KEY") or ""
    user_api_key = st.text_input(
        "Gemini API Key Override (Optional)",
        value="",
        type="password",
        help="If not set in .env, paste your key here. Leave blank to use the key configured in the .env file."
    )"""

    new_lab_ui = """    st.subheader("Upload Filled Lab Marks Sheet")
    lab_marks_file = st.file_uploader(
        "Upload Filled Lab Marks Sheet (.xlsx)",
        type=["xlsx"],
        help="Provide the filled Lab Marks Excel sheet containing student roster and marks."
    )
    
    st.subheader("Upload Course Exit Survey Data")
    ces_file = st.file_uploader(
        "Upload Course Exit Survey (.xlsx)", 
        type=["xlsx"], 
        key="lab_ces",
        help="Optional Course Exit Survey Excel file."
    )"""

    if target_lab_ui not in content:
        print("Error: target_lab_ui not found in app.py!")
        return False
    content = content.replace(target_lab_ui, new_lab_ui)

    # 7. Button validation logic
    target_button_validation = """    elif course_type == "Lab Course" and not rubrics_file:
        st.error("⚠️ Please upload the scanned rubrics PDF containing handwritten marks.")
    elif course_type == "IPCC Course" and not any(marks_files.values()) and not quiz_file and not aat_file and not rubrics_file:
        st.error("⚠️ Please upload at least one assessment data file (Theory or Lab Rubrics) to process.")"""

    new_button_validation = """    elif course_type == "Lab Course" and not lab_marks_file:
        st.error("⚠️ Please upload the filled Lab Marks sheet.")
    elif course_type == "IPCC Course" and not any(marks_files.values()) and not quiz_file and not aat_file and not lab_marks_file:
        st.error("⚠️ Please upload at least one assessment data file (Theory or Lab Marks) to process.")"""

    if target_button_validation not in content:
        print("Error: target_button_validation not found in app.py!")
        return False
    content = content.replace(target_button_validation, new_button_validation)

    # 8. Button click handler execution blocks
    target_handler_exec = """                    api_key = user_api_key if user_api_key else env_key
                    if not api_key:
                        raise ValueError("No Gemini API key found. Please define GEMINI_API_KEY in your .env file or input it above.")

                    out_bytes = process_stand_alone_lab(
                        template_file,
                        rubrics_file,
                        api_key,
                        co_vals,
                        po_vals,
                        override_course_code=manual_course_code if manual_course_code else None,
                        ces_file=ces_file
                    )
                    filename_suggest = "Consolidated_Lab_Scheme.xlsx"
                else:
                    # IPCC Course
                    # Parse CO values (6 values)
                    co_vals = []
                    for val in co_input.split(","):
                        val_str = val.strip()
                        if val_str:
                            try:
                                co_vals.append(float(val_str))
                            except ValueError:
                                co_vals.append(val_str)
                    if len(co_vals) != 6:
                        st.warning(f"Note: You entered {len(co_vals)} CO values. Padded with defaults to make 6.")
                        co_vals = (co_vals + [1.0]*6)[:6]

                    # Parse PO values (6 groups)
                    raw_po_groups = po_input.split(";")
                    po_vals = []
                    for grp in raw_po_groups:
                        grp = grp.strip()
                        if grp:
                            digits = [d.strip() for d in grp.split(",") if d.strip().isdigit()]
                            if digits:
                                po_vals.append("0," + ",".join(digits) + ",0")
                            else:
                                po_vals.append("")
                        else:
                            po_vals.append("")
                    if len(po_vals) != 6:
                        st.warning(f"Note: You entered {len(po_vals)} PO groups. Padded with defaults to make 6.")
                        po_vals = (po_vals + [""]*6)[:6]

                    api_key = user_api_key if user_api_key else env_key
                    if rubrics_file and not api_key:
                        raise ValueError("No Gemini API key found. Please define GEMINI_API_KEY in your .env file or input it above to process the Lab PDF.")

                    out_bytes = process_ipcc_course(
                        template_file,
                        qp_files,
                        marks_files,
                        quiz_file,
                        aat_file,
                        rubrics_file,
                        api_key,
                        co_vals,
                        po_vals,
                        override_course_code=manual_course_code if manual_course_code else None,
                        ces_file=ces_file
                    )
                    filename_suggest = "Consolidated_IPCC_Scheme.xlsx"

                st.success("🎉 Successfully compiled course data and generated spreadsheet!")
                st.download_button(
                    label="📥 Download Consolidated Scheme (.xlsx)",
                    data=out_bytes,
                    file_name=filename_suggest,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ An error occurred during mapping: {str(e)}")
                st.info("Check that all files match the expected formats and that your Gemini API key is active.")"""

    new_handler_exec = """                    out_bytes = process_stand_alone_lab(
                        template_file,
                        lab_marks_file,
                        co_vals,
                        po_vals,
                        override_course_code=manual_course_code if manual_course_code else None,
                        ces_file=ces_file
                    )
                    filename_suggest = "Consolidated_Lab_Scheme.xlsx"
                else:
                    # IPCC Course
                    # Parse CO values (6 values)
                    co_vals = []
                    for val in co_input.split(","):
                        val_str = val.strip()
                        if val_str:
                            try:
                                co_vals.append(float(val_str))
                            except ValueError:
                                co_vals.append(val_str)
                    if len(co_vals) != 6:
                        st.warning(f"Note: You entered {len(co_vals)} CO values. Padded with defaults to make 6.")
                        co_vals = (co_vals + [1.0]*6)[:6]

                    # Parse PO values (6 groups)
                    raw_po_groups = po_input.split(";")
                    po_vals = []
                    for grp in raw_po_groups:
                        grp = grp.strip()
                        if grp:
                            digits = [d.strip() for d in grp.split(",") if d.strip().isdigit()]
                            if digits:
                                po_vals.append("0," + ",".join(digits) + ",0")
                            else:
                                po_vals.append("")
                        else:
                            po_vals.append("")
                    if len(po_vals) != 6:
                        st.warning(f"Note: You entered {len(po_vals)} PO groups. Padded with defaults to make 6.")
                        po_vals = (po_vals + [""]*6)[:6]

                    out_bytes = process_ipcc_course(
                        template_file,
                        qp_files,
                        marks_files,
                        quiz_file,
                        aat_file,
                        lab_marks_file,
                        co_vals,
                        po_vals,
                        override_course_code=manual_course_code if manual_course_code else None,
                        ces_file=ces_file
                    )
                    filename_suggest = "Consolidated_IPCC_Scheme.xlsx"

                st.success("🎉 Successfully compiled course data and generated spreadsheet!")
                st.download_button(
                    label="📥 Download Consolidated Scheme (.xlsx)",
                    data=out_bytes,
                    file_name=filename_suggest,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ An error occurred during mapping: {str(e)}")
                st.info("Check that all files match the expected formats.")"""

    if target_handler_exec not in content:
        print("Error: target_handler_exec not found in app.py!")
        return False
    content = content.replace(target_handler_exec, new_handler_exec)

    with open("app.py", "w", encoding="utf-8") as f:
        f.write(content)
        
    print("Success: All edits applied to app.py!")
    return True

if __name__ == "__main__":
    apply()
