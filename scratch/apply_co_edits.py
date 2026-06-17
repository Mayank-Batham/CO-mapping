import re

def apply():
    with open("app.py", "r", encoding="utf-8") as f:
        content = f.read()

    # 1. Modify extract_lab_marks_from_excel to return (co_vals, po_vals, students_data)
    target_extract = """def extract_lab_marks_from_excel(file_stream, num_marks_cols):
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
            
    return students_data"""

    new_extract = """def extract_lab_marks_from_excel(file_stream, num_marks_cols):
    \"\"\"
    Parses a filled Lab Excel sheet.
    Returns:
        co_vals: list of CO values
        po_vals: list of PO values
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
        
    # Read COs (Row 5) and POs (Row 4) for columns 4 to 4 + num_marks_cols - 1
    co_vals = []
    po_vals = []
    for col in range(4, 4 + num_marks_cols):
        co_vals.append(sheet.cell(row=5, column=col).value)
        po_vals.append(sheet.cell(row=4, column=col).value)
        
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
            
    return co_vals, po_vals, students_data"""

    if target_extract not in content:
        print("Error: extract target helper not found!")
        return False
    content = content.replace(target_extract, target_extract.replace(target_extract, new_extract))

    # 2. Modify process_stand_alone_lab function signature and extraction call
    target_lab_sig = """def process_stand_alone_lab(template_file, lab_marks_file, co_vals, po_vals, override_course_code=None, ces_file=None):
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
        raise ValueError(f"Failed to read lab marks Excel sheet: {e}")"""

    new_lab_sig = """def process_stand_alone_lab(template_file, lab_marks_file, override_course_code=None, ces_file=None):
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
        co_vals, po_vals, extracted_marks_map = extract_lab_marks_from_excel(lab_marks_file, num_marks_cols=13)
    except Exception as e:
        raise ValueError(f"Failed to read lab marks Excel sheet: {e}")"""

    if target_lab_sig not in content:
        print("Error: target_lab_sig not found!")
        return False
    content = content.replace(target_lab_sig, new_lab_sig)

    # 3. Modify process_ipcc_course function signature and extraction call
    target_ipcc_sig = """def process_ipcc_course(template_file, qp_files, marks_files, quiz_file, aat_file, lab_marks_file, co_vals, po_vals, override_course_code=None, ces_file=None):
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
            raise ValueError(f"Failed to read lab marks Excel sheet: {e}")"""

    new_ipcc_sig = """def process_ipcc_course(template_file, qp_files, marks_files, quiz_file, aat_file, lab_marks_file, override_course_code=None, ces_file=None):
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
    co_vals = []
    po_vals = []
    course_code = override_course_code
    
    if lab_marks_file:
        try:
            co_vals, po_vals, extracted_marks_map = extract_lab_marks_from_excel(lab_marks_file, num_marks_cols=6)
        except Exception as e:
            raise ValueError(f"Failed to read lab marks Excel sheet: {e}")"""

    if target_ipcc_sig not in content:
        print("Error: target_ipcc_sig not found!")
        return False
    content = content.replace(target_ipcc_sig, new_ipcc_sig)

    # 4. Modify process_ipcc_course CO-PO writing block to use extracted values
    # In process_ipcc_course:
    # # 6. Write Lab Assessment Marks & CO/PO Mappings (6 labs)
    # # Fill COs and POs mapped in Lab sheet (Row 4 is POs, Row 5 is COs, Cols D to I)
    # for idx in range(6):
    #     col = 4 + idx
    #     if idx < len(po_vals) and po_vals[idx]:
    #         sheet_lab.cell(row=4, column=col).value = po_vals[idx]
    #     if idx < len(co_vals) and co_vals[idx] is not None:
    #         sheet_lab.cell(row=5, column=col).value = co_vals[idx]
    target_ipcc_co_write = """    # 6. Write Lab Assessment Marks & CO/PO Mappings (6 labs)
    # Fill COs and POs mapped in Lab sheet (Row 4 is POs, Row 5 is COs, Cols D to I)
    for idx in range(6):
        col = 4 + idx
        if idx < len(po_vals) and po_vals[idx]:
            sheet_lab.cell(row=4, column=col).value = po_vals[idx]
        if idx < len(co_vals) and co_vals[idx] is not None:
            sheet_lab.cell(row=5, column=col).value = co_vals[idx]"""

    new_ipcc_co_write = """    # 6. Write Lab Assessment Marks & CO/PO Mappings (6 labs)
    # Fill COs and POs mapped in Lab sheet (Row 4 is POs, Row 5 is COs, Cols D to I)
    for idx in range(6):
        col = 4 + idx
        if idx < len(po_vals) and po_vals[idx] is not None:
            sheet_lab.cell(row=4, column=col).value = po_vals[idx]
        if idx < len(co_vals) and co_vals[idx] is not None:
            sheet_lab.cell(row=5, column=col).value = co_vals[idx]"""

    if target_ipcc_co_write not in content:
        print("Error: target_ipcc_co_write not found!")
        return False
    content = content.replace(target_ipcc_co_write, new_ipcc_co_write)

    # 5. UI: Remove CO/PO defaults and config for IPCC Course
    target_ipcc_co_po_ui = """    # Lab defaults (IPCC)
    co_input = "1, 1, 1, 2, 2, 2"
    po_input = "1,5; 1,5; 1,5; 2,5; 2,5; 2,5"
    lab_marks_file = None"""

    new_ipcc_co_po_ui = """    # Lab defaults (IPCC)
    lab_marks_file = None"""

    if target_ipcc_co_po_ui not in content:
        print("Error: target_ipcc_co_po_ui not found!")
        return False
    content = content.replace(target_ipcc_co_po_ui, new_ipcc_co_po_ui)

    target_ipcc_tab4 = """    if course_type == "IPCC Course":
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

    new_ipcc_tab4 = """    if course_type == "IPCC Course":
        with tabs[4]:
            st.subheader("Upload Filled Lab Marks Sheet")
            lab_marks_file = st.file_uploader(
                "Upload Filled Lab Marks Sheet (.xlsx)",
                type=["xlsx"],
                help="Provide the filled IPCC Lab Marks Excel sheet containing student marks."
            )"""

    if target_ipcc_tab4 not in content:
        print("Error: target_ipcc_tab4 not found!")
        return False
    content = content.replace(target_ipcc_tab4, new_ipcc_tab4)

    # 6. UI: Remove CO/PO config for Lab Course
    target_lab_ui_co_po = """else:
    # Lab Course Upload Section
    st.subheader("Configure Lab Mapping Parameters")
    col_co, col_po = st.columns(2)
    with col_co:
        co_input = st.text_input(
            "Lab CO Mappings (Row 5)", 
            value="1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4",
            help="Enter exactly 12 CO values separated by commas for columns D to O."
        )
    with col_po:
        po_input = st.text_input(
            "Lab PO Mappings (Row 4)", 
            value="1,5; 1,5; 1,5; 2,5; 2,5; 2,5; 3,5; 3,5; 3,5; 4,5; 4,5; 4,5",
            help="Enter PO mappings for each lab. Separate labs with a semicolon (;) and PO numbers with commas (,). We format it automatically."
        )
        
    st.subheader("Upload Filled Lab Marks Sheet")
    lab_marks_file = st.file_uploader(
        "Upload Filled Lab Marks Sheet (.xlsx)",
        type=["xlsx"],
        help="Provide the filled Lab Marks Excel sheet containing student roster and marks."
    )"""

    new_lab_ui_co_po = """else:
    # Lab Course Upload Section
    st.subheader("Upload Filled Lab Marks Sheet")
    lab_marks_file = st.file_uploader(
        "Upload Filled Lab Marks Sheet (.xlsx)",
        type=["xlsx"],
        help="Provide the filled Lab Marks Excel sheet containing student roster and marks."
    )"""

    if target_lab_ui_co_po not in content:
        print("Error: target_lab_ui_co_po not found!")
        return False
    content = content.replace(target_lab_ui_co_po, new_lab_ui_co_po)

    # 7. Button click handler: remove co_vals/po_vals parsing and update calls
    # For Lab Course:
    target_lab_handler = """                elif course_type == "Lab Course":
                    # Parse CO values
                    co_vals = []
                    for val in co_input.split(","):
                        val_str = val.strip()
                        if val_str:
                            try:
                                co_vals.append(float(val_str))
                            except ValueError:
                                co_vals.append(val_str)
                    
                    if len(co_vals) != 12:
                        st.warning(f"Note: You entered {len(co_vals)} CO values. Padded with defaults to make 12.")
                        co_vals = (co_vals + [1.0]*12)[:12]

                    # Parse PO values
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
                    
                    if len(po_vals) != 12:
                        st.warning(f"Note: You entered {len(po_vals)} PO groups. Padded with defaults to make 12.")
                        po_vals = (po_vals + [""]*12)[:12]

                    out_bytes = process_stand_alone_lab(
                        template_file,
                        lab_marks_file,
                        co_vals,
                        po_vals,
                        override_course_code=manual_course_code if manual_course_code else None,
                        ces_file=ces_file
                    )"""

    new_lab_handler = """                elif course_type == "Lab Course":
                    out_bytes = process_stand_alone_lab(
                        template_file,
                        lab_marks_file,
                        override_course_code=manual_course_code if manual_course_code else None,
                        ces_file=ces_file
                    )"""

    if target_lab_handler not in content:
        print("Error: target_lab_handler not found!")
        return False
    content = content.replace(target_lab_handler, new_lab_handler)

    # For IPCC Course:
    target_ipcc_handler = """                else:
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
                    )"""

    new_ipcc_handler = """                else:
                    # IPCC Course
                    out_bytes = process_ipcc_course(
                        template_file,
                        qp_files,
                        marks_files,
                        quiz_file,
                        aat_file,
                        lab_marks_file,
                        override_course_code=manual_course_code if manual_course_code else None,
                        ces_file=ces_file
                    )"""

    if target_ipcc_handler not in content:
        print("Error: target_ipcc_handler not found!")
        return False
    content = content.replace(target_ipcc_handler, new_ipcc_handler)

    with open("app.py", "w", encoding="utf-8") as f:
        f.write(content)

    print("Success: All CO/PO edits applied to app.py!")
    return True

if __name__ == "__main__":
    apply()
