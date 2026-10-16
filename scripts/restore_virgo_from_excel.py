import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import shutil
import datetime
import sqlite3
import openpyxl

# Set console output encoding to utf-8
if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

from src.core.bom_filter_manager import BOMFilterManager

def restore_virgo():
    print("==================================================")
    print("STEP 1: READING VIRGO RULES FROM EXCEL FILE")
    print("==================================================")
    excel_path = Path(r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\15. FORM（BIEU MAU）-形式\Form_VBA\Form_VBA_SS_BOM\tonghop_new12052026_ma1.xlsm")
    
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
        
    wb = openpyxl.load_workbook(str(excel_path), data_only=True)
    ws = wb["BolocBom"]
    
    virgo_rules = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx == 1:
            # Header row: ('Virgo', 'Trạng thái lọc', 'Mã Unit/Linh kiện', ...)
            continue
        item_name = row[0]
        match_mode = row[1]
        part_code = row[2]
        
        if item_name is not None and str(item_name).strip():
            clean_item = str(item_name).strip()
            clean_mode = "Part_name" if match_mode and "part" in str(match_mode).lower() else "Full_name"
            clean_code = str(part_code).strip() if part_code is not None and str(part_code).strip() else None
            virgo_rules.append({
                "item_name": clean_item,
                "match_mode": clean_mode,
                "part_code": clean_code
            })
            
    wb.close()
    print(f"-> Parsed {len(virgo_rules)} rules for Virgo from Excel file.")

    print("\n==================================================")
    print("STEP 2: BACKING UP DATABASE BEFORE RESTORATION")
    print("==================================================")
    mgr = BOMFilterManager()
    remote_path, is_remote = mgr.get_active_db_path()
    print(f"Target DB Path: {remote_path} (is_remote={is_remote})")
    
    if remote_path.exists():
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = remote_path.parent / f"ssbom_master.db.bak_{timestamp}"
        print(f"Creating safety backup at: {backup_file}")
        shutil.copy2(remote_path, backup_file)
        print("-> Safety backup created successfully.")

    print("\n==================================================")
    print("STEP 3: RESTORING VIRGO RULES VIA DUAL-WRITE")
    print("==================================================")
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    note_text = "Phục hồi từ tonghop_new12052026_ma1.xlsm"
    
    records = [
        (
            "Virgo",
            r["item_name"],
            r["match_mode"],
            r["part_code"],
            note_text,
            1,
            now_str,
            now_str,
        )
        for r in virgo_rules
    ]
    
    def _do_restore_virgo(conn: sqlite3.Connection):
        # 1. Delete all existing Virgo rules
        del_cur = conn.execute("DELETE FROM bolocbom_rules WHERE LOWER(model_name) = 'virgo';")
        deleted_count = del_cur.rowcount
        # 2. Insert the fresh 50 rules
        conn.executemany(
            """
            INSERT INTO bolocbom_rules (model_name, item_name, match_mode, part_code, notes, is_active, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?);
            """,
            records,
        )
        return deleted_count, len(records)
        
    ok, result, used_fallback = mgr._execute_dual_write(_do_restore_virgo, operation_name="phục hồi bộ lọc Virgo")
    if not ok:
        raise RuntimeError(f"Dual-write failed: {result}")
        
    del_count, ins_count = result
    print(f"-> Successfully deleted {del_count} old rules and inserted {ins_count} restored rules.")
    print(f"-> Used fallback: {used_fallback}")

    print("\n==================================================")
    print("STEP 4: VERIFYING RESTORATION RESULTS")
    print("==================================================")
    restored_rules = mgr.get_rules_for_model("Virgo", active_only=True)
    print(f"Total active Virgo rules returned by BOMFilterManager: {len(restored_rules)}")
    assert len(restored_rules) == len(virgo_rules), f"Mismatch: expected {len(virgo_rules)}, got {len(restored_rules)}"
    
    # Check all rules match
    for idx, (expected, actual) in enumerate(zip(virgo_rules, restored_rules), start=1):
        assert actual.item_name == expected["item_name"], f"Row {idx} item mismatch: {actual.item_name} != {expected['item_name']}"
        assert actual.match_mode == expected["match_mode"], f"Row {idx} mode mismatch: {actual.match_mode} != {expected['match_mode']}"
        assert actual.part_code == expected["part_code"], f"Row {idx} code mismatch: {actual.part_code} != {expected['part_code']}"
    
    print("-> All 50 rules verified 100% identical to Excel!")
    
    # Check other models count
    with mgr._get_connection(mgr.get_read_db_path(force_refresh=True)) as conn:
        cur = conn.execute("SELECT model_name, COUNT(*) FROM bolocbom_rules GROUP BY model_name ORDER BY model_name;")
        print("\nCurrent summary in bolocbom_rules:")
        for r in cur.fetchall():
            print(f"  - {r[0]}: {r[1]} rules")
            
    print("\n[OK] VIRGO FILTER RESTORATION COMPLETE!")

if __name__ == "__main__":
    restore_virgo()
