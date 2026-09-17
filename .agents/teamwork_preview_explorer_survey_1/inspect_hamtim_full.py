import openpyxl

wb = openpyxl.load_workbook(r"D:\Sandbox\pm_sosanhbom\Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx", data_only=False)
ws = wb.active

with open(r"D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1\hamtim_analysis.txt", "w", encoding="utf-8") as out:
    out.write(f"Active sheet: {ws.title}, max_row={ws.max_row}, max_col={ws.max_column}\n\n")

    # Inspect first 15 rows
    for r in range(1, 16):
        out.write(f"=== ROW {r} ===\n")
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if val is not None:
                coord = ws.cell(r, c).coordinate
                out.write(f"  Col {c} ({coord}): {val}\n")

    # Also check what columns exist across the sheet:
    out.write("\n=== DISTINCT FORMULAS IN KEY COLUMNS ===\n")
    # check columns 1 to 35 on rows 2, 3, 4, 5
    for c in range(1, ws.max_column + 1):
        sample_vals = [str(ws.cell(r, c).value) for r in range(1, 10)]
        out.write(f"Col {c}: sample={sample_vals[:4]}\n")

wb.close()
print("hamtim_analysis.txt written successfully")
