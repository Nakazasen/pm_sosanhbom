import os
import openpyxl
from pptx import Presentation

work_dir = r"D:\Sandbox\pm_sosanhbom"
out_dir = r"D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_1"

# Inspect PPTX
pptx_path = os.path.join(work_dir, "Chương trình so sánh BOM tự động.pptx")
if os.path.exists(pptx_path):
    prs = Presentation(pptx_path)
    with open(os.path.join(out_dir, "pptx_summary.txt"), "w", encoding="utf-8") as f:
        for idx, slide in enumerate(prs.slides):
            f.write(f"\n--- SLIDE {idx+1} ---\n")
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        text = paragraph.text.strip()
                        if text:
                            f.write(f"{text}\n")
    print("Extracted pptx_summary.txt")

# Inspect Excel workbooks
excel_files = [
    "Hamtimlinhkienthuoc_UNIT_naotren_BOM.xlsx",
    "form_ssbom.xlsm",
    "formnguoidung.xlsm",
    "tonghop_new12052026_ma1.xlsm"
]

with open(os.path.join(out_dir, "excel_sheets_summary.txt"), "w", encoding="utf-8") as out:
    for fname in excel_files:
        fpath = os.path.join(work_dir, fname)
        if not os.path.exists(fpath):
            continue
        out.write(f"\n=========================================\n")
        out.write(f"FILE: {fname}\n")
        out.write(f"=========================================\n")
        try:
            wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
            out.write(f"Sheet names: {wb.sheetnames}\n")
            for sname in wb.sheetnames:
                ws = wb[sname]
                out.write(f"\n  --- Sheet: {sname} (max_row={ws.max_row}, max_column={ws.max_column}) ---\n")
                # print first 5 rows
                rows_count = 0
                for row in ws.iter_rows(values_only=True):
                    # filter out None
                    vals = [str(c) if c is not None else "" for c in row[:25]]
                    if any(vals):
                        out.write(f"    Row {rows_count+1}: {' | '.join(vals)}\n")
                        rows_count += 1
                    if rows_count >= 6:
                        break
            wb.close()
        except Exception as e:
            out.write(f"Error inspecting {fname}: {e}\n")

print("Excel inspection complete!")
