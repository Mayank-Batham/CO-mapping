import pandas as pd
from openpyxl import load_workbook

print("--- EXPLORING 2022_SCHEME EXCEL ---")
scheme_path = '2022_SCHEME_3 SEM_22AI34_Data Structures with Applications_Calculation_Reshma S-Print (1).xlsx'
try:
    xl = pd.ExcelFile(scheme_path)
    df = xl.parse('Theory', nrows=15, header=None)
    print("First 15 rows of Theory sheet:")
    for idx, row in df.iterrows():
        print(f"Row {idx}:", [str(x)[:15] if not pd.isna(x) else '' for x in row.values])
except Exception as e:
    print("Error reading scheme:", e)

print("\n--- EXPLORING IA_MARKS XLS ---")
xls_path = 'IA_Marks (4).xls'
try:
    try:
        # It is likely HTML inside an xls file
        with open(xls_path, 'r', encoding='utf-8') as f:
            content = f.read(1000)
            if '<html' in content.lower() or '<table' in content.lower():
                print("Confirmed HTML content.")
                dfs = pd.read_html(xls_path)
                print("HTML Tables found:", len(dfs))
                for i, df in enumerate(dfs):
                    print(f"Table {i} head:")
                    print(df.head())
            else:
                print("Not HTML content, first characters:", content[:100])
    except Exception as e:
        print("HTML read failed:", e)
        
    xl = pd.ExcelFile(xls_path, engine='xlrd')
    df = xl.parse(xl.sheet_names[0], nrows=15, header=None)
    print("First 15 rows of xls via xlrd:")
    for idx, row in df.iterrows():
         print(f"Row {idx}:", [str(x)[:20] if not pd.isna(x) else '' for x in row.values])
except Exception as e:
    print("End Error reading xls:", e)
