import openpyxl

wb = openpyxl.load_workbook(r"D:\Sandbox\pm_sosanhbom\form_ssbom.xlsm", data_only=False)

with open(r"D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1\form_ssbom_formulas.txt", "w", encoding="utf-8") as out:
    for sname in wb.sheetnames:
        ws = wb[sname]
        out.write(f"\n=========================================\n")
        out.write(f"SHEET: {sname} (max_row={ws.max_row}, max_col={ws.max_column})\n")
        out.write(f"=========================================\n")
        
        # Check rows 1 to 5
        for r in range(1, min(6, ws.max_row + 1)):
            row_items = []
            for c in range(1, min(30, ws.max_column + 1)):
                val = ws.cell(r, c).value
                if val is not None:
                    row_items.append(f"Col {c} ({ws.cell(r, c).coordinate}): {val}")
            if row_items:
                out.write(f"Row {r}:\n  " + "\n  ".join(row_items) + "\n")
        
        # Check distinct formula columns in row 3 or row 2
        out.write("\nDistinct formulas across columns (checking rows 2, 3, 4):\n")
        for c in range(1, min(35, ws.max_column + 1)):
            c_letter = openpyxl.utils.get_column_letter(c)
            formulas = []
            for r in range(1, min(10, ws.max_row + 1)):
                v = str(ws.cell(r, c).value)
                if v.startswith("="):
                    formulas.append(f"R{r}:{v}")
            if formulas:
                out.write(f"  Col {c_letter} ({c}): {', '.join(formulas[:3])}\n")

wb.close()
print("form_ssbom_formulas.txt written successfully")
