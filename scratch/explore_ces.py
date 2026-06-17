import pandas as pd
import openpyxl

path = 'course exit survey.xlsx'
try:
    xl = pd.ExcelFile(path)
    print("Sheets in course exit survey.xlsx:", xl.sheet_names)
    for name in xl.sheet_names:
        df = xl.parse(name, nrows=20, header=None)
        print(f"\n--- Sheet: {name} (shape: {df.shape}) ---")
        for idx, row in df.iterrows():
            print(f"Row {idx}:", [str(x)[:30] if not pd.isna(x) else '' for x in row.values])
except Exception as e:
    print("Error:", e)
