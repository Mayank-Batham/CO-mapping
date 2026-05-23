import pandas as pd

xls_path = 'AML_IA1_Marks.xls'
df = pd.read_excel(xls_path, engine='xlrd', header=None)
print("Rows 10 to 30:")
for idx, row in df.iloc[10:30].iterrows():
    print(f"Row {idx}:", [str(x) if not pd.isna(x) else '' for x in row.values])
