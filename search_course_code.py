import pandas as pd

df = pd.read_excel('AML_IA1_Marks.xls', engine='xlrd', header=None)
print("--- Searching for course code in AML_IA1_Marks.xls ---")
for r in range(len(df)):
    for c in range(len(df.columns)):
        val = str(df.iloc[r, c]).strip()
        if val and val != 'nan':
            # Check if it looks like a course code
            if '22AI' in val or '52' in val or 'AML' in val or '|' in val:
                print(f"Cell ({r}, {c}): '{val}'")
