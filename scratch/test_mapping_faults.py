import os
import re
import pandas as pd
import openpyxl

def extract_metadata(file_path):
    print(f"\n--- Checking metadata extraction for {os.path.basename(file_path)} ---")
    try:
        df = pd.read_excel(file_path, engine='xlrd', header=None)
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return None
        
    metadata = {
        'faculty': '',
        'course_code': '',
        'course_title': '',
        'semester': '',
        'academic_year': '',
        'batch': ''
    }
    
    found_any = False
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
                    found_any = True
            
            if 'staff name' in val_strip.lower() or 'faculty' in val_strip.lower():
                m = re.search(r'(?:staff name|faculty)\s*[:|-]+\s*(.*)', val_strip, re.IGNORECASE)
                if m:
                    metadata['faculty'] = m.group(1).strip()
                    found_any = True
                    
    if metadata['semester'] and metadata['academic_year']:
        try:
            sem = int(metadata['semester'])
            ay_start = int(metadata['academic_year'].split('-')[0])
            start_year = ay_start - ((sem + 1) // 2 - 1)
            end_year = start_year + 4
            metadata['batch'] = f"{start_year}-{str(end_year)[2:]}"
        except Exception as e:
            pass
            
    print(f"Extracted Metadata: {metadata}")
    return metadata

def test_matching(metadata, src_mapping_file):
    print(f"\n--- Testing course code matching against {os.path.basename(src_mapping_file)} ---")
    if not metadata or not metadata.get('course_code'):
        print("FAIL: No course_code found in metadata! Skipping mapping.")
        return
        
    target_code = re.sub(r'[^a-zA-Z0-9]', '', str(metadata['course_code'])).upper()
    
    # Extract core code by removing leading scheme indicators (e.g., "22", "21")
    m_target = re.match(r'^\d+(.*)', target_code)
    target_core = m_target.group(1) if m_target else target_code
    print(f"Target Course Code: {target_code}")
    print(f"Target Core Code (scheme-stripped): {target_core}")
    
    try:
        engine = 'openpyxl' if src_mapping_file.lower().endswith('.xlsx') else 'xlrd'
        raw_df = pd.read_excel(src_mapping_file, engine=engine, header=None)
        
        # Scan first 15 rows to find the header row containing 'Course Code'
        header_row_idx = 0
        course_code_col_idx = 0
        found_header = False
        
        for r_idx in range(min(15, len(raw_df))):
            row_vals = [str(x).strip().upper() for x in raw_df.iloc[r_idx].values]
            for c_idx, val in enumerate(row_vals):
                if ('COURSE' in val and 'CODE' in val) or ('SUBJ' in val and 'CODE' in val):
                    header_row_idx = r_idx
                    course_code_col_idx = c_idx
                    found_header = True
                    break
            if found_header:
                break
                
        if found_header:
            print(f"Found Course Code header at Row {header_row_idx}, Column {course_code_col_idx}")
            cols = [str(x).strip() for x in raw_df.iloc[header_row_idx].values]
            seen = {}
            new_cols = []
            for col in cols:
                if col in seen:
                    seen[col] += 1
                    new_cols.append(f"{col}_{seen[col]}")
                else:
                    seen[col] = 0
                    new_cols.append(col)
                    
            src_df = raw_df.iloc[header_row_idx+1:].copy()
            src_df.columns = new_cols
            course_code_col = new_cols[course_code_col_idx]
        else:
            print("WARNING: 'Course Code' or 'Subj Code' header not found. Falling back to default parser.")
            src_df = pd.read_excel(src_mapping_file, engine=engine)
            course_code_col = None
            for col in src_df.columns:
                col_norm = re.sub(r'[^a-zA-Z0-9]', '', str(col)).upper()
                if 'COURSECODE' in col_norm or 'SUBJCODE' in col_norm:
                    course_code_col = col
                    break
            if course_code_col is None:
                course_code_col = src_df.columns[0]
                print(f"Falling back to first column as course code column: {course_code_col}")
        
        print(f"Using Course Code Column: '{course_code_col}'")
        
        # Normalize Course Code column values to search for target_code
        match_idx = None
        for idx, val in enumerate(src_df[course_code_col].values):
            if pd.notna(val):
                norm_val = re.sub(r'[^a-zA-Z0-9]', '', str(val)).upper()
                # Try exact match first
                if norm_val == target_code:
                    match_idx = idx
                    print(f"Found EXACT match: row index {idx}, value in cell: '{val}'")
                    break
                # Fallback: Compare core code (scheme-agnostic, e.g. "AI51" vs "AI51")
                m_candidate = re.match(r'^\d+(.*)', norm_val)
                candidate_core = m_candidate.group(1) if m_candidate else norm_val
                if target_core and len(target_core) >= 3:
                    if target_core in candidate_core or candidate_core in target_core:
                        match_idx = idx
                        print(f"Found CORE match (using scheme fallback): row index {idx}, value in cell: '{val}' (candidate core: '{candidate_core}')")
                        break
        
        if match_idx is None:
            print("FAIL: No matching course found in mapping file!")
            print("Unique codes in that column first 10 rows:")
            print(src_df[course_code_col].dropna().unique()[:15])
            return
            
        # Collect sequential rows belonging to this course
        co_rows = []
        for idx in range(match_idx, len(src_df)):
            val_code = src_df.iloc[idx][course_code_col]
            if idx > match_idx and pd.notna(val_code) and str(val_code).strip() != '':
                break
            co_rows.append(src_df.iloc[idx])
            
        print(f"Found {len(co_rows)} mapping rows (COs) for the matched course.")
        for r_num, row_data in enumerate(co_rows):
            print(f"  CO Row {r_num+1} values: {dict(row_data.dropna())}")

    except Exception as e:
        print(f"Error during matching: {e}")

if __name__ == '__main__':
    marks_files = [
        'AML_IA1_Marks.xls',
        'AML_IA2_Marks.xls',
        'AML_IA3_Marks.xls',
        'AML_IA_Marks (Quiz).xls',
        'AML_IA_Marks(AAT).xls',
        'IA_Marks cie 1.xls',
        'IA_Marks (3).pdf',
        'IA_Marks (4).xls',
    ]
    
    mapping_file = '2021-CO-PO -PSO-MAPPING.xlsx'
    
    for f in marks_files:
        if os.path.exists(f) and f.endswith('.xls'):
            meta = extract_metadata(f)
            if meta and meta.get('course_code'):
                test_matching(meta, mapping_file)
