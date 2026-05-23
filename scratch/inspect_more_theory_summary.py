import openpyxl

wb = openpyxl.load_workbook("STAND ALONE THEORY EMPTY TEMPLATE.xlsx", data_only=False)
sheet = wb["Theory"]
for r in range(86, 110):
    val = sheet.cell(row=r, column=4).value # Column D (4)
    if val:
        print(f"Row {r}: {val}")
