import streamlit as st
import pandas as pd
import docx
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import io
import re

st.set_page_config(
    page_title="CIA Marks & CO Mapper",
    page_icon="📊",
    layout="wide"
)

# Custom Premium Styling
st.markdown("""
<style>
    /* Premium font styling and background enhancements */
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    .main-title {
        background: linear-gradient(135deg, #6C63FF 0%, #3F3D56 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 3rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
    }
    
    .subtitle {
        color: #707070;
        font-size: 1.2rem;
        margin-bottom: 2rem;
    }
    
    .card {
        background-color: #f8fafc;
        border-radius: 12px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        border: 1px solid #e2e8f0;
    }
    
    .card h3 {
        color: #0f172a !important;
        margin-top: 0px !important;
        margin-bottom: 8px !important;
        font-weight: 700 !important;
        font-size: 1.25rem !important;
    }
    
    .card p {
        color: #475569 !important;
        font-size: 0.95rem !important;
        line-height: 1.5 !important;
        margin-bottom: 0px !important;
    }
    
    .tag-coming-soon {
        background: linear-gradient(135deg, #FF6B6B 0%, #FF8E53 100%);
        color: white;
        padding: 3px 8px;
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: bold;
        display: inline-block;
        margin-left: 10px;
    }
    
    .tag-active {
        background: linear-gradient(135deg, #10B981 0%, #059669 100%);
        color: white;
        padding: 3px 8px;
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: bold;
        display: inline-block;
        margin-left: 10px;
    }
    
    /* Sleek buttons */
    .stButton>button {
        background: linear-gradient(135deg, #6C63FF 0%, #514DFD 100%);
        color: white;
        border-radius: 8px;
        padding: 0.6rem 2rem;
        font-weight: 600;
        border: none;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(108, 99, 255, 0.4);
    }
    
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(108, 99, 255, 0.6);
        background: linear-gradient(135deg, #514DFD 0%, #6C63FF 100%);
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- PARSING HELPER FUNCTIONS -----------------

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

def extract_single_column_marks(file_stream):
    """Parses Quiz/AAT XLS and extracts single column marks (Theory Mark) per USN."""
    file_bytes = file_stream.read()
    file_stream.seek(0)
    df = pd.read_excel(io.BytesIO(file_bytes), engine='xlrd', header=None)
    header_row_idx = None
    usn_col = None
    marks_col = None
    
    for idx, row in df.iterrows():
        row_vals = [str(x).strip() for x in row.values]
        if 'USN' in row_vals:
            header_row_idx = idx
            usn_col = row_vals.index('USN')
            # Look for 'Theory Mark' or similar
            for col_idx, val in enumerate(row_vals):
                if 'theory mark' in val.lower() or 'marks' in val.lower() or 'mark' in val.lower():
                    if 'practical' not in val.lower():
                        marks_col = col_idx
                        break
            break
            
    if header_row_idx is None or usn_col is None:
        return {}
    if marks_col is None:
        marks_col = usn_col + 2 # fallback if not found directly
        
    marks_data = {}
    for idx in range(header_row_idx + 1, len(df)):
        row = df.iloc[idx]
        usn = str(row.iloc[usn_col]).strip()
        if not usn or usn.lower() in ('nan', 'none', '') or not usn.isalnum():
            continue
        
        mark = row.iloc[marks_col]
        if pd.isna(mark) or str(mark).strip() in ['-', '', 'nan', 'A', 'None']:
            pass
        else:
            try:
                marks_data[usn] = float(mark)
            except ValueError:
                pass
                
    return marks_data

def extract_metadata(file_stream):
    """Extracts course metadata (Course Code, Faculty, etc.) from IA Marks file."""
    file_bytes = file_stream.read()
    file_stream.seek(0)
    df = pd.read_excel(io.BytesIO(file_bytes), engine='xlrd', header=None)
    metadata = {
        'faculty': '',
        'course_code': '',
        'course_title': '',
        'semester': '',
        'academic_year': '',
        'batch': ''
    }
    
    for idx, row in df.iterrows():
        row_str = [str(x) for x in row.values]
        for val in row_str:
            val_strip = val.strip()
            if ' - ' in val_strip and ('CIA' in val_strip or 'AIML' in val_strip):
                parts = val_strip.split(' - ')
                if len(parts) >= 3:
                    metadata['course_code'] = parts[1].strip()
                    metadata['course_title'] = parts[2].strip()
                    p0_parts = parts[0].split('-')
                    if len(p0_parts) >= 4:
                        metadata['semester'] = p0_parts[1].strip()
                        metadata['academic_year'] = f"{p0_parts[3].strip()}-{p0_parts[4].strip()}" if len(p0_parts) > 4 else p0_parts[3].strip()
            
            if 'staff name' in val_strip.lower() or 'faculty' in val_strip.lower():
                m = re.search(r'(?:staff name|faculty)\s*[:|-]+\s*(.*)', val_strip, re.IGNORECASE)
                if m:
                    metadata['faculty'] = m.group(1).strip()
                    
    if metadata['semester'] and metadata['academic_year']:
        try:
            sem = int(metadata['semester'])
            ay_start = int(metadata['academic_year'].split('-')[0])
            start_year = ay_start - ((sem + 1) // 2 - 1)
            end_year = start_year + 4
            metadata['batch'] = f"{start_year}-{str(end_year)[2:]}"
        except Exception as e:
            pass
            
    return metadata

# ----------------- CORE PROCESSING PIPELINE -----------------

def process_stand_alone_theory(template_file, qp_files, marks_files, quiz_file, aat_file):
    """Fills the Stand Alone Theory sheet cleanly and handles dynamic rosters."""
    wb = load_workbook(io.BytesIO(template_file.read()))
    sheet = wb['Theory'] if 'Theory' in wb.sheetnames else wb.active
    
    # 1. Configure Course Type Cells
    sheet.cell(row=7, column=4).value = "YES" # STAND ALONE THEORY
    sheet.cell(row=8, column=4).value = "NO"  # STAND ALONE LAB
    sheet.cell(row=9, column=4).value = "NO"  # IPCC
    
    # 2. Build Master Student Roster
    master_students = {}
    all_marks_files = []
    
    # Extract list of files
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
                    
    sorted_usns = sorted(master_students.keys())
    num_students = len(sorted_usns)
    
    if num_students == 0:
        raise ValueError("No students found in the uploaded marks files. Ensure Excel files are correct.")
        
    # 3. Populate Roster in Theory sheet (Rows 17 to 85)
    for i in range(85 - 17 + 1):
        r = 17 + i
        if i < num_students:
            usn = sorted_usns[i]
            sheet.cell(row=r, column=1).value = float(i + 1)
            sheet.cell(row=r, column=2).value = usn
            sheet.cell(row=r, column=3).value = master_students[usn]
        else:
            # Clear unused row columns A to DH (1 to 122)
            for c in range(1, 123):
                sheet.cell(row=r, column=c).value = None
                
    # 4. Clear all marks cells in student rows (Theory Cols D to CW, Rows 17 to 17 + N - 1)
    for r in range(17, 17 + num_students):
        for c in range(4, 102): # Col 4 (D) to 101 (CW)
            sheet.cell(row=r, column=c).value = None
            
    # 5. Clear unused student rows in other sheets
    if 'Lab' in wb.sheetnames:
        sheet_lab = wb['Lab']
        for r in range(9 + num_students, 78):
            for c in range(1, sheet_lab.max_column + 1):
                sheet_lab.cell(row=r, column=c).value = None
                
    if 'CIA MARKS' in wb.sheetnames:
        sheet_cia = wb['CIA MARKS']
        for r in range(14 + num_students, 84):
            for c in range(1, sheet_cia.max_column + 1):
                sheet_cia.cell(row=r, column=c).value = None
                
    # 6. Extract and write metadata from IA1 Marks if available
    metadata = {}
    for test_idx in [1, 2, 3]:
        if marks_files.get(test_idx):
            marks_files[test_idx].seek(0)
            metadata = extract_metadata(marks_files[test_idx])
            marks_files[test_idx].seek(0)
            if metadata['course_code']:
                break
                
    if metadata:
        if metadata['academic_year']:
            sheet.cell(row=4, column=1).value = f"Academic Year : {metadata['academic_year']}"
        if metadata['batch']:
            sheet.cell(row=4, column=4).value = f"Batch :{metadata['batch']}"
        if metadata['semester']:
            sheet.cell(row=4, column=17).value = f"Semester  : {metadata['semester']}"
        if metadata['course_code']:
            sheet.cell(row=5, column=1).value = f"COURSE  CODE : {metadata['course_code']}"
        if metadata['course_title']:
            sheet.cell(row=6, column=1).value = f"COURSE TITLE : {metadata['course_title']}"
        if metadata['faculty']:
            sheet.cell(row=10, column=1).value = f"Faculty : {metadata['faculty']}"
            
    # 7. Map Test Columns
    tests_col_map = {1: {}, 2: {}, 3: {}} 
    first_q1a_col = -1
    for col in range(1, 200):
        val = sheet.cell(row=16, column=col).value
        val_str = str(val).strip() if val else ""
        if val_str == 'Q1A':
            first_q1a_col = col
            break

    if first_q1a_col != -1:
        for test_idx in [1, 2, 3]:
            start_col = first_q1a_col + (test_idx - 1) * 32
            for q_num in range(1, 9):
                tests_col_map[test_idx][q_num] = start_col + (q_num - 1) * 4
                
    # 8. Write Test COs and Marks
    for test_idx in [1, 2, 3]:
        qp = qp_files.get(test_idx)
        marks = marks_files.get(test_idx)
        col_mappings = tests_col_map[test_idx]
        
        # CO Mappings
        if qp:
            qp.seek(0)
            co_map = extract_cos_from_docx(qp)
            qp.seek(0)
            for col in col_mappings.values():
                for offset in range(4):
                    sheet.cell(row=12, column=col + offset).value = None
            
            for q_num, co_val in co_map.items():
                if q_num in col_mappings:
                    col = col_mappings[q_num]
                    sheet.cell(row=12, column=col).value = f"0,{co_val}"
                    
        # Marks Mappings
        if marks:
            marks.seek(0)
            st_marks = extract_marks_from_xls(marks)
            marks.seek(0)
            
            for i in range(num_students):
                r = 17 + i
                usn_val = sheet.cell(row=r, column=2).value
                if usn_val in st_marks:
                    for q_num, mark in st_marks[usn_val].items():
                        if q_num in col_mappings:
                            col = col_mappings[q_num]
                            sheet.cell(row=r, column=col).value = mark

        # 8.5 Set maximum marks to 10 for A part of every question (Q1A to Q8A)
        # to eliminate division by zero errors in summaries!
        for q_num, col in col_mappings.items():
            sheet.cell(row=14, column=col).value = 10.0
            col_letter = get_column_letter(col)
            sheet.cell(row=15, column=col).value = f"=PRODUCT({col_letter}14,0.65)"
                            
    # 9. Map Quiz Marks (Column 101)
    if quiz_file:
        quiz_file.seek(0)
        quiz_marks = extract_single_column_marks(quiz_file)
        quiz_file.seek(0)
        for i in range(num_students):
            r = 17 + i
            usn_val = sheet.cell(row=r, column=2).value
            if usn_val in quiz_marks:
                sheet.cell(row=r, column=101).value = quiz_marks[usn_val]
                
    # 10. Map AAT Marks (Column 100)
    if aat_file:
        aat_file.seek(0)
        aat_marks = extract_single_column_marks(aat_file)
        aat_file.seek(0)
        for i in range(num_students):
            r = 17 + i
            usn_val = sheet.cell(row=r, column=2).value
            if usn_val in aat_marks:
                sheet.cell(row=r, column=100).value = aat_marks[usn_val]

    # 11. Prevent division by zero in Final SEE Grade (Col 108/DD) by filling active student rows with "NA"
    for i in range(num_students):
        r = 17 + i
        sheet.cell(row=r, column=108).value = "NA"

    # 12. Prevent division by zero in CES survey sheet by initializing row 11 with 0
    if 'CES' in wb.sheetnames:
        sheet_ces = wb['CES']
        for col in range(4, 14): # D to M
            if sheet_ces.cell(row=11, column=col).value is None:
                sheet_ces.cell(row=11, column=col).value = 0

    # 13. Prevent division by zero and fill CO-PO-PSO mapping matrix based on Course Code
    if 'CO_PO_PSO_MAPPING' in wb.sheetnames:
        sheet_cppm = wb['CO_PO_PSO_MAPPING']
        
        # Default initialization: fill all cells in CO1-CO10 (D4:R13) with 0 to prevent division by zero
        for r in range(4, 14):
            for c in range(4, 19): # Columns D to R
                sheet_cppm.cell(row=r, column=c).value = 0

        # Attempt to automatically load mapping based on course code
        if metadata and metadata.get('course_code'):
            target_code = re.sub(r'[^a-zA-Z0-9]', '', str(metadata['course_code'])).upper()
            
            import os
            src_mapping_file = None
            for f_name in os.listdir('.'):
                if 'CO-PO' in f_name and 'MAPPING' in f_name and f_name.endswith('.xlsx'):
                    src_mapping_file = f_name
                    break
            if not src_mapping_file and os.path.exists('2021-CO-PO -PSO-MAPPING.xlsx'):
                src_mapping_file = '2021-CO-PO -PSO-MAPPING.xlsx'
                
            if src_mapping_file:
                try:
                    src_df = pd.read_excel(src_mapping_file, engine='openpyxl')
                    
                    # Normalize Course Code column values to search for target_code
                    match_idx = None
                    for idx, val in enumerate(src_df['Course Code'].values):
                        if pd.notna(val):
                            norm_val = re.sub(r'[^a-zA-Z0-9]', '', str(val)).upper()
                            if norm_val == target_code:
                                match_idx = idx
                                break
                                
                    if match_idx is not None:
                        # Collect sequential rows belonging to this course
                        co_rows = []
                        for idx in range(match_idx, len(src_df)):
                            if idx > match_idx and pd.notna(src_df.iloc[idx]['Course Code']):
                                break
                            co_rows.append(src_df.iloc[idx])
                            
                        # Extract header columns mapping (e.g. 'PO1' -> Col 4)
                        header_cols = {}
                        for col in range(4, 19): # D to R
                            header_val = sheet_cppm.cell(row=3, column=col).value
                            if header_val:
                                header_cols[str(header_val).strip()] = col
                                
                        # Write the mapped values into the template
                        for co_idx, co_row in enumerate(co_rows):
                            target_row = 4 + co_idx # CO1 -> Row 4, CO2 -> Row 5...
                            if target_row > 13: # Limit to CO10
                                break
                                
                            for col_name, col_idx in header_cols.items():
                                if col_name in co_row:
                                    val = co_row[col_name]
                                    if pd.notna(val) and str(val).strip() != '':
                                        try:
                                            sheet_cppm.cell(row=target_row, column=col_idx).value = float(val)
                                        except ValueError:
                                            sheet_cppm.cell(row=target_row, column=col_idx).value = str(val)
                                    else:
                                        # Leave unmapped cells in active rows as empty/None or 0
                                        sheet_cppm.cell(row=target_row, column=col_idx).value = None
                except Exception as e:
                    # Log mapping warnings gracefully to Streamlit
                    st.warning(f"Note: Could not automatically map CO-PO weights from '{src_mapping_file}': {e}")
                
    out_stream = io.BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# ----------------- STREAMLIT INTERFACE -----------------

st.markdown('<div class="main-title">📊 CIA Marks & CO Mapper</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Seamlessly automate Course Outcome mapping and marks consolidation.</div>', unsafe_allow_html=True)

# Configuration Panel
st.header("1. Choose Course Configuration")

col_type1, col_type2, col_type3 = st.columns(3)

with col_type1:
    st.markdown("""
    <div class="card">
        <h3>Theory Course <span class="tag-active">Active</span></h3>
        <p>Map CIA Question Papers, Student Marks, AAT, and Quiz into the Stand Alone Theory template.</p>
    </div>
    """, unsafe_allow_html=True)
    
with col_type2:
    st.markdown("""
    <div class="card" style="opacity: 0.6;">
        <h3>Lab Course <span class="tag-coming-soon">Soon</span></h3>
        <p>Consolidate laboratory component evaluations and mappings cleanly.</p>
    </div>
    """, unsafe_allow_html=True)

with col_type3:
    st.markdown("""
    <div class="card" style="opacity: 0.6;">
        <h3>IPCC Course <span class="tag-coming-soon">Soon</span></h3>
        <p>Integrated Professional Core Course mapping combining both Theory and Practical slots.</p>
    </div>
    """, unsafe_allow_html=True)

# Sidebar - Template Upload
with st.sidebar:
    st.markdown('<h2 style="font-weight: 800;">⚙️ Setup Template</h2>', unsafe_allow_html=True)
    template_file = st.file_uploader(
        "Upload Stand Alone Theory Template (.xlsx)",
        type=["xlsx"],
        help="Provide the empty STAND ALONE THEORY EMPTY TEMPLATE.xlsx workbook."
    )
    st.markdown("---")
    st.markdown("<small>Designed for department course mapping operations.</small>", unsafe_allow_html=True)

st.header("2. Upload Assessment Data")

# Tabbed Layout or Grouped Layout
tab_cia1, tab_cia2, tab_cia3, tab_other = st.tabs(["📌 CIA-1", "📌 CIA-2", "📌 CIA-3", "📌 Quiz & AAT"])

qp_files = {}
marks_files = {}
quiz_file = None
aat_file = None

with tab_cia1:
    col1, col2 = st.columns(2)
    with col1:
        qp_files[1] = st.file_uploader("CIA-1 Question Paper (.docx)", type=["docx"], key="qp1")
    with col2:
        marks_files[1] = st.file_uploader("CIA-1 IA Marks (.xls)", type=["xls"], key="m1")

with tab_cia2:
    col1, col2 = st.columns(2)
    with col1:
        qp_files[2] = st.file_uploader("CIA-2 Question Paper (.docx)", type=["docx"], key="qp2")
    with col2:
        marks_files[2] = st.file_uploader("CIA-2 IA Marks (.xls)", type=["xls"], key="m2")

with tab_cia3:
    col1, col2 = st.columns(2)
    with col1:
        qp_files[3] = st.file_uploader("CIA-3 Question Paper (.docx)", type=["docx"], key="qp3")
    with col2:
        marks_files[3] = st.file_uploader("CIA-3 IA Marks (.xls)", type=["xls"], key="m3")

with tab_other:
    col1, col2 = st.columns(2)
    with col1:
        quiz_file = st.file_uploader("Quiz Marks (.xls)", type=["xls"], key="quiz")
    with col2:
        aat_file = st.file_uploader("AAT Marks (.xls)", type=["xls"], key="aat")

st.markdown("<br>", unsafe_allow_html=True)

if st.button("Generate Consolidated Excel", type="primary"):
    if not template_file:
        st.error("⚠️ Please configure the Excel Template in the sidebar before processing.")
    elif not any(marks_files.values()) and not quiz_file and not aat_file:
        st.error("⚠️ Please upload at least one marks file (CIA, Quiz, or AAT) to map scores.")
    else:
        with st.spinner("Processing templates and compiling rosters..."):
            try:
                out_bytes = process_stand_alone_theory(
                    template_file,
                    qp_files,
                    marks_files,
                    quiz_file,
                    aat_file
                )
                st.success("🎉 Successfully compiled course data and generated spreadsheet!")
                
                # Fetch course details if available to suggest name
                filename_suggest = "Consolidated_Theory_Scheme.xlsx"
                st.download_button(
                    label="📥 Download Consolidated Scheme (.xlsx)",
                    data=out_bytes,
                    file_name=filename_suggest,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ An error occurred during mapping: {str(e)}")
                st.info("Ensure all uploaded XLS files are standard course marks sheets and DOCX files match the question paper layout.")
