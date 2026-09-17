import openpyxl

wb = openpyxl.load_workbook(r"D:\Sandbox\pm_sosanhbom\Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx", data_only=False)
ws = wb.active
print(f"Active sheet: {ws.title}, max_row={ws.max_row}, max_col={ws.max_column}")

# Print row 1 to 5 headers and values/formulas
for r in range(1, 10):
    row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
    # print non-none
    non_empty = [(c, ws.cell(r, c).coordinate, val) for c, val in enumerate(row_vals, 1) if val is not None]
    print(f"Row {r}: {len(non_empty)} cells")
    for c, coord, val in non_empty[:15]:
        print(f"   Col {c} ({coord}): {val}")

# Print row 1 specifically (column names)
headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
print("\nHeaders (Row 1):")
for idx, h in enumerate(headers, 1):
    if h is not None:
        print(f"  Col {idx}: {h}")

# Check sample data rows (e.g. rows 2, 3, 4, 10, 100)
print("\nSample row 2:")
for c in range(1, ws.max_column + 1):
    val = ws.cell(2, c).value
    if val is not None:
        print(f"  Col {c} ({headers[c-1]}): {val}")

print("\nSample row 10:")
for c in range(1, ws.max_column + 1):
    val = ws.cell(10, c).value
    if val is not None:
        print(f"  Col {c} ({headers[c-1]}): {val}")

wb.close()
