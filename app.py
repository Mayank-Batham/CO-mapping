import streamlit as st
import pandas as pd
import docx
from openpyxl import load_workbook
import io
import re

st.set_page_config(page_title="CIA Mark Mapper", layout="wide")

def extract_cos_from_docx(file_stream):
    """Parses a DOCX file and extracts Course Outcome mappings."""
    document = docx.Document(file_stream)
    co_mapping = {} # Q -> CO
    for table in document.tables:
        q_row_idx = None
        co_row_idx = None
        for i, row in enumerate(table.rows):
            cells = [cell.text.strip() for cell in row.cells]
            if not cells: continue
            if 'Question No.' in cells[0]:
                q_row_idx = i
            if 'Course Outcome' in cells[0] or 'Course outcome' in cells[0]:
                co_row_idx = i
        if q_row_idx is not None and co_row_idx is not None:
            q_cells = [c.text.strip() for c in table.rows[q_row_idx].cells]
            co_cells = [c.text.strip() for c in table.rows[co_row_idx].cells]
            for q, co in zip(q_cells[1:], co_cells[1:]):
                if q and q.isdigit():
                    co_mapping[int(q)] = co.replace("CO", "").strip()
    return co_mapping

def extract_marks_from_xls(file_stream):
    """Parses IA Marks XLS generated and extracts marks per USN."""
    # Reading via xlrd for .xls support
    file_bytes = file_stream.read()
    file_stream.seek(0)
    df = pd.read_excel(io.BytesIO(file_bytes), engine='xlrd', header=None)
    header_row_idx = None
    usn_col = None
    for idx, row in df.iterrows():
        row_vals = [str(x).strip() for x in row.values]
        if 'USN' in row_vals:
            header_row_idx = idx
            usn_col = row_vals.index('USN')
            break
            
    if header_row_idx is None: 
        return {}
    
    headers = [str(x).strip().replace('.0', '') for x in df.iloc[header_row_idx].values]
    
    q_cols = {} # mapped question int -> col id
    for col_idx, h in enumerate(headers):
        if h.isdigit():
            q_cols[int(h)] = col_idx
            
    marks_data = {} # USN -> {q_num: mark}
    for idx in range(header_row_idx + 1, len(df)):
        row = df.iloc[idx]
        usn = str(row.iloc[usn_col]).strip()
        if not usn or usn.lower() in ('nan', 'none', '') or not usn.isalnum():
            continue
        
        st_marks = {}
        for q_num, col_idx in q_cols.items():
            mark = row.iloc[col_idx]
            if pd.isna(mark) or str(mark).strip() in ['-', '', 'nan', 'A', 'None']:
                pass
            else:
                try:
                    st_marks[q_num] = float(mark)
                except ValueError:
                    pass
        marks_data[usn] = st_marks
        
    return marks_data

def process_scheme(template_file, qp_files, marks_files):
    """Reads template, applies marks & COs, returns bytes of new workbook."""
    wb = load_workbook(io.BytesIO(template_file.read()))
    sheet = wb['Theory'] if 'Theory' in wb.sheetnames else wb.active
    
    # Locate key rows and offset columns
    header_row = -1
    usn_col = -1
    co_row = -1
    
    for row in range(1, 30):
        # Look for CO row
        c_val = sheet.cell(row=row, column=3).value
        c_val_str = str(c_val).strip() if c_val else ""
        if 'COs mapped' in c_val_str or 'CO mapped' in c_val_str:
            co_row = row
        
        # Look for Header Row
        for col in range(1, 10):
            val = sheet.cell(row=row, column=col).value
            val_str = str(val).strip() if val else ""
            if 'USN' == val_str:
                header_row = row
                usn_col = col
                break
        if header_row != -1 and co_row != -1:
            break
            
    if header_row == -1:
        raise ValueError("Could not find the 'USN' header in the Theory sheet.")
        
    # Map Test Columns
    tests_col_map = {1: {}, 2: {}, 3: {}} 
    first_q1a_col = -1
    for col in range(1, 200):
        val = sheet.cell(row=header_row, column=col).value
        val_str = str(val).strip() if val else ""
        if val_str == 'Q1A':
            first_q1a_col = col
            break

    if first_q1a_col != -1:
        for test_idx in [1, 2, 3]:
            # Each test group has 32 columns (8 questions * 4 parts)
            start_col = first_q1a_col + (test_idx - 1) * 32
            for q_num in range(1, 9):
                # QnA column is offset by 4 for each subsequent question
                tests_col_map[test_idx][q_num] = start_col + (q_num - 1) * 4

    # Process tests
    for test_idx in [1, 2, 3]:
        qp = qp_files.get(test_idx)
        marks = marks_files.get(test_idx)
        
        if qp and marks:
            co_map = extract_cos_from_docx(qp)
            st_marks = extract_marks_from_xls(marks)
            
            col_mappings = tests_col_map[test_idx]
            
            # Write CO mapped values
            if co_row != -1:
                # Clear existing values for this test, then write new ones to be safe
                for col in col_mappings.values():
                    for offset in range(4):
                        sheet.cell(row=co_row, column=col + offset).value = ""
                
                for q_num, co_val in co_map.items():
                    if q_num in col_mappings:
                        col = col_mappings[q_num]
                        sheet.cell(row=co_row, column=col).value = f"0,{co_val}"
            
            # Write Student Marks
            for row in range(header_row + 1, sheet.max_row + 1):
                usn_val = sheet.cell(row=row, column=usn_col).value
                usn_val = str(usn_val).strip() if usn_val else ""
                
                if usn_val in st_marks:
                    # Clear out existing marks for the test's columns (A, B, C, D)
                    for col_idx in col_mappings.values():
                        for offset in range(4):
                            sheet.cell(row=row, column=col_idx + offset).value = ""
                    
                    # Write the extracted marks
                    for q_num, mark in st_marks[usn_val].items():
                        if q_num in col_mappings:
                            col = col_mappings[q_num]
                            sheet.cell(row=row, column=col).value = mark
                            
    out_stream = io.BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# Streamlit App Layout
st.title("CIA Marks & CO Mapper")
st.markdown("Upload the required files to merge the CIA scores and Course Outcomes into the Scheme template.")

with st.sidebar:
    st.header("Excel Template")
    template_file = st.file_uploader("Upload 2022_SCHEME Template (.xlsx)", type=["xlsx"])

st.header("Upload Data")
col1, col2, col3 = st.columns(3)

qp_files = {}
marks_files = {}

with col1:
    st.subheader("CIA-1")
    qp_files[1] = st.file_uploader("CIA-1 Question Paper (.docx)", type=["docx"], key="qp1")
    marks_files[1] = st.file_uploader("CIA-1 IA Marks (.xls)", type=["xls"], key="m1")

with col2:
    st.subheader("CIA-2")
    qp_files[2] = st.file_uploader("CIA-2 Question Paper (.docx)", type=["docx"], key="qp2")
    marks_files[2] = st.file_uploader("CIA-2 IA Marks (.xls)", type=["xls"], key="m2")

with col3:
    st.subheader("CIA-3")
    qp_files[3] = st.file_uploader("CIA-3 Question Paper (.docx)", type=["docx"], key="qp3")
    marks_files[3] = st.file_uploader("CIA-3 IA Marks (.xls)", type=["xls"], key="m3")

if st.button("Process & Generate Excel", type="primary"):
    if not template_file:
        st.error("Please upload the template Excel file.")
    else:
        with st.spinner("Processing..."):
            try:
                out_bytes = process_scheme(template_file, qp_files, marks_files)
                st.success("Successfully generated the data mapped scheme!")
                
                st.download_button(
                    label="Download Mapped Scheme (.xlsx)",
                    data=out_bytes,
                    file_name="Mapped_Scheme.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
