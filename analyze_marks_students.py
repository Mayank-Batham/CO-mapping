import pandas as pd

files = [
    'AML_IA1_Marks.xls',
    'AML_IA2_Marks.xls',
    'AML_IA3_Marks.xls',
    'AML_IA_Marks (Quiz).xls',
    'AML_IA_Marks(AAT).xls'
]

all_students = {} # USN -> Name

for f_path in files:
    df = pd.read_excel(f_path, engine='xlrd', header=None)
    header_row_idx = None
    usn_col = None
    for idx, row in df.iterrows():
        row_vals = [str(x).strip() for x in row.values]
        if 'USN' in row_vals:
            header_row_idx = idx
            usn_col = row_vals.index('USN')
            break
    
    if header_row_idx is not None:
        # Find Student Name or Student column
        row_vals = [str(x).strip() for x in df.iloc[header_row_idx].values]
        name_col = None
        for c_idx, val in enumerate(row_vals):
            if 'student' in val.lower() or 'name' in val.lower():
                if 'staff' not in val.lower() and 'division' not in val.lower():
                    name_col = c_idx
                    break
        
        if name_col is None:
            name_col = usn_col + 1 # fallback
            
        file_students = {}
        for idx in range(header_row_idx + 1, len(df)):
            row = df.iloc[idx]
            u = str(row.iloc[usn_col]).strip()
            n = str(row.iloc[name_col]).strip()
            if u and u.lower() not in ('nan', 'none', '') and u.isalnum():
                file_students[u] = n
                all_students[u] = n
        print(f"File {f_path}: found {len(file_students)} students")

print(f"Total unique students across all files: {len(all_students)}")
# Sort the student USNs to see them
sorted_usns = sorted(all_students.keys())
print("First 5 sorted USNs:", sorted_usns[:5])
print("Last 5 sorted USNs:", sorted_usns[-5:])
