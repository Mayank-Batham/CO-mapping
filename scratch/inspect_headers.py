import openpyxl

wb = openpyxl.load_workbook("STAND ALONE THEORY EMPTY TEMPLATE.xlsx", data_only=False)

print("--- THEORY SHEET COL DD ---")
sheet = wb["Theory"]
print("DD14:", sheet["DD14"].value)
print("DD15:", sheet["DD15"].value)
print("DD16:", sheet["DD16"].value)
print("DD17:", sheet["DD17"].value)
print("DD88 (Formula):", sheet["DD88"].value)

print("\n--- LAB SHEET J6 & J83 ---")
sheet_lab = wb["Lab"]
print("J5:", sheet_lab["J5"].value)
print("J6:", sheet_lab["J6"].value)
print("J7:", sheet_lab["J7"].value)
print("J81 (Formula):", sheet_lab["J81"].value)
print("J82 (Formula):", sheet_lab["J82"].value)
print("J83 (Formula):", sheet_lab["J83"].value)

print("\n--- CO_PO_PSO_MAPPING D4:R4 ---")
sheet_map = wb["CO_PO_PSO_MAPPING"]
print("D4:", sheet_map["D4"].value)
print("R4:", sheet_map["R4"].value)
print("S4 (Formula):", sheet_map["S4"].value)

print("\n--- CO_PO_ASSESSMENT K40 ---")
sheet_assess = wb["CO_PO_ASSESSMENT"]
print("K40 (Formula):", sheet_assess["K40"].value)
