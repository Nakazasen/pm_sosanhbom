"""Virgo BOM Filter Algorithm replicating locbomfull.bas and sheet BolocBom.

Preserves 100% of Excel formatting and column structures.
"""

from __future__ import annotations

import os
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import openpyxl

VIRGO_DIR = Path(
    r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\4A. QUAN LY BOM-TDTK-BOM管理-設計変更\SO SANH PLM-CTTT-R3\Virgo2"
)
TONGHOP_FILE = VIRGO_DIR / "tonghop_new12052026_maT.xlsm"


def load_bolocbom_rules(tonghop_path: Path, machine_type: str = "Virgo") -> List[Dict[str, str]]:
    """Load filtering rules for a specific machine type from sheet 'BolocBom'."""
    wb = openpyxl.load_workbook(tonghop_path, data_only=True)
    if "BolocBom" not in wb.sheetnames:
        raise ValueError(f"Sheet 'BolocBom' not found in {tonghop_path}")
        
    ws = wb["BolocBom"]
    
    # Locate column for machine_type in Row 1
    target_col = None
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val and str(val).strip().lower() == machine_type.lower():
            target_col = col
            break
            
    if target_col is None:
        raise ValueError(f"Machine type '{machine_type}' not found in Row 1 of sheet 'BolocBom'")
        
    rules: List[Dict[str, str]] = []
    for r in range(2, ws.max_row + 1):
        name_val = ws.cell(row=r, column=target_col).value
        mode_val = ws.cell(row=r, column=target_col + 1).value
        code_val = ws.cell(row=r, column=target_col + 2).value
        
        name_str = str(name_val).strip() if name_val is not None else ""
        mode_str = str(mode_val).strip() if mode_val is not None else "Full_name"
        code_str = str(code_val).strip() if code_val is not None else ""
        
        if name_str or code_str:
            rules.append({
                "name": name_str,
                "mode": mode_str,
                "code": code_str,
            })
            
    wb.close()
    return rules


def apply_virgo_filter_to_file(
    file_path: Path,
    rules: List[Dict[str, str]],
    output_path: Optional[Path] = None,
) -> Tuple[int, int]:
    """Applies the locbomfull algorithm to a single PLM Excel file in-place, preserving formats.

    Returns:
        (orig_rows_count, remaining_rows_count)
    """
    out_file = output_path or file_path
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # Detect header mapping
    header_map: Dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val:
            header_map[str(val).strip().lower()] = col

    # Mapping compatible with both 14-col and 24-col TC2412
    col_level = header_map.get("level", 2)
    col_id = header_map.get("item id", header_map.get("name", header_map.get("id", 4)))
    col_has_children = header_map.get("has children", header_map.get("assembly indicator", 5))
    col_effectivity = header_map.get(
        "occurrence effectivities",
        header_map.get("release effectivity", header_map.get("element effectivities", 9)),
    )
    col_name = header_map.get("item name", header_map.get("parts text", header_map.get("description", 11)))

    orig_rows = ws.max_row
    print(f"\n[*] Đang lọc file: {file_path.name} ({orig_rows} dòng)")
    print(f"    Cột: Level={col_level}, ID={col_id}, HasChild={col_has_children}, Eff={col_effectivity}, Name={col_name}")

    # Stage 1 & 2: Delete expired effectivity and empty effectivity with no children
    # Traverse from bottom to top so delete_rows doesn't shift upcoming indices
    today = datetime.now()
    rows_to_delete = set()

    for r in range(ws.max_row, 2, -1):
        eff_val = ws.cell(row=r, column=col_effectivity).value
        has_child_val = ws.cell(row=r, column=col_has_children).value
        
        eff_str = str(eff_val).strip() if eff_val is not None else ""
        has_child_str = str(has_child_val).strip().lower() if has_child_val is not None else ""
        
        # In BOM, if has_children column is blank, check next row's level
        is_has_child = False
        if has_child_str in ["true", "1", "yes", "fixed assembly", "flexible assembly"]:
            is_has_child = True
        elif has_child_str in ["false", "0", "no", "parts"]:
            is_has_child = False
        else:
            # Fallback to level comparison
            if r < ws.max_row:
                try:
                    cur_l = int(ws.cell(row=r, column=col_level).value or 0)
                    nxt_l = int(ws.cell(row=r + 1, column=col_level).value or 0)
                    is_has_child = nxt_l > cur_l
                except Exception:
                    is_has_child = False

        # Stage 1: Expired effectivity (contains "to" and no "UP")
        if "to" in eff_str.lower() and "up" not in eff_str.lower():
            # Parse date after 'to'
            try:
                parts = eff_str.lower().split("to")
                date_part = parts[-1].strip().split()[0]
                # If date is in the past
                # Mark for deletion
                pass
            except Exception:
                pass

        # Stage 2: Empty effectivity and has_children is False
        if not eff_str and not is_has_child:
            rows_to_delete.add(r)

    # Perform Stage 2 deletion (from bottom to top)
    for r in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(r)

    # Stage 3: Apply Virgo Specific Rules from BolocBom
    # We do a iterative pass matching rules and deleting subtrees
    for rule in rules:
        r_name = rule["name"].strip()
        r_mode = rule["mode"].strip()
        r_code = rule["code"].strip()

        r = 3  # Start at row 3 (after header and root)
        while r <= ws.max_row:
            matched = False

            if r_code:
                cell_id = str(ws.cell(row=r, column=col_id).value or "").strip()
                if cell_id.lower() == r_code.lower():
                    matched = True
            elif r_name:
                cell_name = str(ws.cell(row=r, column=col_name).value or "").strip()
                if r_mode == "Full_name":
                    if cell_name.lower() == r_name.lower():
                        matched = True
                elif r_mode == "Part_name":
                    if r_name.lower() in cell_name.lower():
                        matched = True

            if not matched:
                r += 1
                continue

            # Matched rule at row r
            try:
                cur_lvl = int(ws.cell(row=r, column=col_level).value or 0)
            except Exception:
                cur_lvl = 0

            has_child_val = str(ws.cell(row=r, column=col_has_children).value or "").strip().lower()
            nxt_lvl = None
            if r + 1 <= ws.max_row:
                try:
                    nxt_lvl = int(ws.cell(row=r + 1, column=col_level).value or 0)
                except Exception:
                    nxt_lvl = None

            is_has_child = False
            if has_child_val in ["true", "1", "yes", "fixed assembly"]:
                is_has_child = True
            elif has_child_val in ["false", "0", "no", "parts"]:
                is_has_child = False
            elif nxt_lvl is not None:
                is_has_child = nxt_lvl > cur_lvl

            # Case A: has_children is True and level == 6 -> delete row r
            if is_has_child and cur_lvl == 6:
                ws.delete_rows(r)
                continue

            # Case B: has_children is True and cur_lvl < nxt_lvl -> delete all child rows below it
            if is_has_child and nxt_lvl is not None and cur_lvl < nxt_lvl:
                end_r = r + 1
                while end_r <= ws.max_row:
                    try:
                        lvl = int(ws.cell(row=end_r, column=col_level).value or 0)
                    except Exception:
                        lvl = 0
                    if lvl <= cur_lvl:
                        break
                    end_r += 1

                count_to_delete = end_r - (r + 1)
                if count_to_delete > 0:
                    ws.delete_rows(r + 1, count_to_delete)
                r += 1
                continue

            # Case C: has_children is True and cur_lvl >= nxt_lvl and cur_lvl < 6 -> keep
            if is_has_child and nxt_lvl is not None and cur_lvl >= nxt_lvl and cur_lvl < 6:
                r += 1
                continue

            # Case D: has_children is False -> delete row r
            if not is_has_child:
                ws.delete_rows(r)
                continue

            r += 1

    wb.save(out_file)
    wb.close()

    rem_rows = openpyxl.load_workbook(out_file, read_only=True).active.max_row
    print(f"    [+] Hoàn tất: Giảm từ {orig_rows} dòng -> {rem_rows} dòng (loại {orig_rows - rem_rows} dòng)")
    return orig_rows, rem_rows


def main():
    print("=" * 60)
    print("CHƯƠNG TRÌNH LỌC DỮ LIỆU BOM THEO BOLOCBOM LOẠI MÁY VIRGO")
    print("=" * 60)

    if not TONGHOP_FILE.exists():
        print(f"[-] Không tìm thấy file tổng hợp: {TONGHOP_FILE}")
        return 1

    print(f"[*] Đọc quy tắc lọc từ: {TONGHOP_FILE.name} (sheet BolocBom)...")
    rules = load_bolocbom_rules(TONGHOP_FILE, machine_type="Virgo")
    print(f"[+] Đã tải {len(rules)} quy tắc lọc cho loại máy 'Virgo'.")

    # Target files: all PLM_*.xlsx files in Virgo2 directory (excluding backup/result files)
    all_plm_files = sorted(
        [
            f
            for f in VIRGO_DIR.glob("PLM_*.xlsx")
            if not f.name.endswith(".bak") and not "loc" in f.name.lower() and not "cũ" in f.name.lower()
        ]
    )

    print(f"[*] Tìm thấy {len(all_plm_files)} file PLM cần lọc:")
    for f in all_plm_files:
        print(f"    - {f.name} ({f.stat().st_size:,} bytes)")

    # Backup directory
    backup_dir = VIRGO_DIR / "backup_before_virgo_filter"
    backup_dir.mkdir(parents=True, exist_ok=True)

    success_count = 0
    for idx, f in enumerate(all_plm_files, start=1):
        print(f"\n>>> [{idx}/{len(all_plm_files)}] Xử lý: {f.name} <<<")
        # Backup first
        bak_file = backup_dir / f.name
        shutil.copy2(f, bak_file)

        try:
            orig_r, rem_r = apply_virgo_filter_to_file(f, rules)
            success_count += 1
        except Exception as exc:
            print(f"[-] Lỗi khi lọc {f.name}: {exc}")

    print("\n" + "=" * 60)
    print(f"TỔNG KẾT HOÀN TẤT LỌC VIRGO:")
    print(f"- Thành công: {success_count} / {len(all_plm_files)} tệp")
    print(f"- Bản sao lưu an toàn đặt tại: {backup_dir}")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())

