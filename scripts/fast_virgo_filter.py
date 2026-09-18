"""High-Performance Canonical 14-Column Virgo BOM Standardizer.

Converts raw 24-column Teamcenter 2412 exports into the canonical 14-column format:
- Prunes electrical subtrees (PE Dept 4) to produce pure mechanical BOM (PE Dept 2)
- Reorganizes into 14 canonical columns:
    Col A: Home (1-based index)
    Col B: Level
    Col C: Item Type ('Parts')
    Col D: Item Id (Part Number)
    Col E: Has Children (None)
    Col F: Quantity
    Col G: 1st Parts (None)
    Col H: 2nd BOM Flag
    Col I: Occurrence Effectivities (None)
    Col J: Item Revision Project List (None)
    Col K: Item Name (Parts Text)
    Col L: Notice No
    Col M: Revision
    Col N: Item Rev Status ('Released')
- Applies MS Gothic / Calibri 11 styling, grey header #C0C0C0, exact column widths, and gridlines
- Saves standardized files to Virgo2 root and auto-synchronizes to model subfolders
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
from typing import List, Optional, Tuple

sys.path.insert(0, os.path.abspath("."))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import openpyxl

from src.automation.tc2412.standardizer import (
    CANONICAL_14_COLUMNS,
    CANONICAL_14_WIDTHS,
    standardize_plm_file,
    transform_24_to_14_columns,
    write_14_column_workbook,
)

VIRGO_DIR = Path(
    r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\4A. QUAN LY BOM-TDTK-BOM管理-設計変更\SO SANH PLM-CTTT-R3\Virgo2"
)
BACKUP_DIR = VIRGO_DIR / "backup_before_virgo_filter"

PART_NUMBERS: List[str] = [
    "T10C423NL0",
    "T10C4A2US0",
    "T10C4B2US0",
    "T10C4C3NL0",
    "T10C4D2US0",
    "T10C4F3NL0",
    "T10C4H3NL0",
    "T10C4K3NL0",
    "T10C4L9JP0",
    "T10C4M3NL0",
    "T10C4N2US0",
    "T10C433NL0",
    "T10C452US0",
    "T10C463NL0",
    "T10C483NL0",
    "T10C493NL0",
]


def find_subfolder_for_part(virgo_dir: Path, part_number: str) -> Optional[Path]:
    """Locate the corresponding model subfolder in Virgo2."""
    for d in virgo_dir.iterdir():
        if d.is_dir() and d.name.upper().startswith(part_number.upper()):
            return d
    return None


def process_virgo_part(
    part_number: str,
    backup_dir: Path = BACKUP_DIR,
    virgo_dir: Path = VIRGO_DIR,
) -> Tuple[int, int, Path, Optional[Path]]:
    """Standardize a single Virgo part into 14 columns and sync to subfolder.

    Returns:
        (orig_row_count, final_row_count, root_dest_path, subfolder_dest_path)
    """
    raw_file = backup_dir / f"PLM_{part_number}.xlsx"
    if not raw_file.exists():
        raise FileNotFoundError(f"Không tìm thấy bản gốc raw tại: {raw_file}")

    root_dest = virgo_dir / f"PLM_{part_number}.xlsx"
    orig_c, final_c = standardize_plm_file(raw_file, root_dest, prune_electrical=True)

    # Sync to model subfolder if present
    subfolder = find_subfolder_for_part(virgo_dir, part_number)
    subfolder_dest = None
    if subfolder and subfolder.exists():
        subfolder_dest = subfolder / f"PLM_{part_number}.xlsx"
        # If reference file already exists in T10C423NL0, make safe backup first
        if subfolder_dest.exists() and part_number == "T10C423NL0":
            ref_bak = subfolder / "PLM_T10C423NL0.original_manual.xlsx"
            if not ref_bak.exists():
                shutil.copy2(subfolder_dest, ref_bak)
        shutil.copy2(root_dest, subfolder_dest)

    return orig_c, final_c, root_dest, subfolder_dest


def main() -> int:
    print("=" * 70, flush=True)
    print("CHƯƠNG TRÌNH CHUẨN HÓA 14 CỘT TIÊU CHUẨN BOM VIRGO2 (TC2412)", flush=True)
    print("Khớp 100% định dạng Canonical A-N dùng cho Tool so sánh và Excel VBA", flush=True)
    print(f"Thư mục gốc: {VIRGO_DIR}", flush=True)
    print("=" * 70, flush=True)

    if not BACKUP_DIR.exists():
        print(f"[-] Thư mục bản gốc không tồn tại: {BACKUP_DIR}", flush=True)
        return 1

    success_count = 0
    total = len(PART_NUMBERS)

    for idx, part in enumerate(sorted(PART_NUMBERS), 1):
        try:
            orig_c, final_c, r_dest, s_dest = process_virgo_part(part)
            sub_info = f" -> {s_dest.parent.name}" if s_dest else " (Không có thư mục con)"
            print(
                f"[{idx:2d}/{total}] ✅ {part}: {orig_c:,} dòng raw -> {final_c:,} dòng 14 cột chuẩn | Sync: {sub_info}",
                flush=True,
            )
            success_count += 1
        except Exception as exc:
            print(f"[{idx:2d}/{total}] ❌ Lỗi khi chuẩn hóa {part}: {exc}", flush=True)

    print("\n" + "=" * 70, flush=True)
    print(f"KẾT QUẢ CHUẨN HÓA:")
    print(f"- Thành công: {success_count}/{total} tệp BOM 14 cột")
    print(f"- Định dạng: 14 Cột (Home, Level, Item Type, Item Id, Has Children, Quantity,")
    print(f"                      1st Parts, 2nd BOM Flag, Occurrence Effectivities,")
    print(f"                      Item Revision Project List, Item Name, Notice No, Revision, Item Rev Status)")
    print(f"- Đã đồng bộ vào toàn bộ thư mục phụ trách của kỹ sư Chế tạo 2.")
    print("=" * 70, flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
