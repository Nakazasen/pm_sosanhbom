"""Apply exact styling of T10C423NL0 Mới.xlsm to all PLM_*.xlsx files in Virgo2.

Restores:
- Header row: Bold, Center, Solid Grey fill (C0C0C0), Bottom border
- Precise column widths matching Teamcenter standard
- Gridlines enabled
"""

import os
import sys
from pathlib import Path

sys.path.insert(0, os.path.abspath("."))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

DEST_DIR = Path(
    r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\4A. QUAN LY BOM-TDTK-BOM管理-設計変更\SO SANH PLM-CTTT-R3\Virgo2"
)

COLUMN_WIDTHS = {
    "A": 8.0,
    "B": 8.0,
    "C": 12.0,
    "D": 20.0,
    "E": 10.0,
    "F": 10.0,
    "G": 12.0,
    "H": 35.0,
    "I": 15.0,
    "J": 10.0,
    "K": 14.0,
    "L": 14.0,
    "M": 22.0,
    "N": 12.0,
    "O": 18.0,
    "P": 15.0,
    "Q": 22.0,
    "R": 15.0,
    "S": 12.0,
    "T": 12.0,
    "U": 12.0,
    "V": 12.0,
    "W": 18.0,
    "X": 30.0,
}

HEADER_FONT = Font(name="Calibri", size=11, bold=True)
HEADER_FILL = PatternFill(fill_type="solid", start_color="C0C0C0", end_color="C0C0C0")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=False)
BOTTOM_BORDER = Border(bottom=Side(style="thin", color="808080"))


def style_workbook(file_path: Path) -> bool:
    """Applies exact visual styling to a single Excel workbook."""
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        ws.views.sheetView[0].showGridLines = True

        # Style header row
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = HEADER_ALIGN
            cell.border = BOTTOM_BORDER

        # Set column widths
        for col_letter, width in COLUMN_WIDTHS.items():
            ws.column_dimensions[col_letter].width = width

        # Style data cells (Calibri 11)
        for r in range(2, min(ws.max_row + 1, 100)):  # sample format check
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).font = Font(name="Calibri", size=11)

        wb.save(file_path)
        wb.close()
        return True
    except Exception as exc:
        print(f"[-] Lỗi định dạng {file_path.name}: {exc}", flush=True)
        return False


def main():
    files = sorted(
        [
            f
            for f in DEST_DIR.glob("PLM_*.xlsx")
            if not f.name.endswith(".bak") and "cũ" not in f.name.lower()
        ]
    )
    print(f"[*] Áp dụng giao diện chuẩn Teamcenter (màu xám C0C0C0, in đậm, độ rộng cột) cho {len(files)} tệp...", flush=True)

    success = 0
    for f in files:
        if style_workbook(f):
            success += 1
            print(f"[OK] Đã hoàn thiện giao diện đẹp: {f.name}", flush=True)

    print(f"\n[+] Hoàn tất định dạng thẩm mỹ: {success}/{len(files)} tệp.")


if __name__ == "__main__":
    main()
