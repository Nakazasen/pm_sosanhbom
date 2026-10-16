import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import shutil
import datetime
import sqlite3

from src.core.bom_filter_manager import BOMFilterManager

VIRGO_RULES_62 = [
    ("BOTTLE WASTE", "Full_name", None),
    ("FILL UP CONTAINER ASSY", "Part_name", None),
    ("PWB PANEL MAIN ASSY WITH SOFTWARE", "Full_name", None),
    ("PWB KEY-R ASSY", "Full_name", None),
    ("PWB KEY-L ASSY", "Full_name", None),
    ("PWB LCDCONT ASSY WITH SOFTWARE", "Full_name", None),
    ("HOLDER BELT TENSION R", "Full_name", None),
    ("HOLDER HEATER", "Full_name", None),
    ("HOLDER THERMAL-CUTOUT", "Full_name", None),
    ("HOLDER BELT TENSION L", "Full_name", None),
    ("HOLDER THERMISTOR", "Full_name", None),
    ("CONN.CORD ASSY S04301 HEATER CON-FUSER B", "Full_name", None),
    ("BUSH EXIT", "Full_name", None),
    ("PWB FUSER ASSY", "Full_name", None),
    ("PULLEY FEED P ASSY", "Full_name", None),
    ("ONE-WAY CLUTCH 6", "Full_name", None),
    ("GEAR Z23R ONEWAY", "Full_name", None),
    ("PWB ASSY TX", "Full_name", None),
    ("PWB ASSY RX WITH SOFTWARE", "Full_name", None),
    ("PULLEY PICKUP ASSY", "Full_name", None),
    ("PWB CIS CONNECT ASSY", "Full_name", None),
    ("CASE DLP", "Full_name", None),
    ("P2200 DEVE /D26", "Part_name", None),
    ("PWB DLP ASSY", "Full_name", None),
    ("PWB TPC ASSY", "Full_name", None),
    ("PWB MOTOR ASSY", "Full_name", None),
    ("FRAME CONVEYING", "Full_name", None),
    ("PWB ID SENSOR ASSY", "Full_name", None),
    ("PWB SUB ASSY", "Full_name", None),
    ("PWB THERMISTOR ASSY", "Full_name", None),
    ("PWB DRUM CONNECT ASSY", "Full_name", None),
    ("PWB SWITCH ASSY", "Full_name", None),
    ("PWB MAIN ASSY", "Part_name", None),
    ("PWB EXIT ASSY", "Full_name", None),
    ("PWB CONTAINER CONNECT ASSY", "Full_name", None),
    ("PWB THERMISTOR2 ASSY", "Full_name", None),
    ("PWB ENGINE CONNECT ASSY", "Full_name", None),
    ("PWB ERASER ASSY", "Full_name", None),
    ("PWB DRUM ASSY", "Full_name", None),
    ("FRAME DRUM", "Full_name", None),
    ("DRUM ASSY WITH PACKING", "Full_name", None),
    ("PWB APC A ASSY", "Full_name", None),
    ("HOUSING LSU", "Full_name", None),
    ("PWB APC B ASSY", "Full_name", None),
    ("P.W.BOARD ASSY PD B", "Full_name", None),
    ("P.W.BOARD ASSY PD A", "Full_name", None),
    ("PULLEY PICKUP P ASSY", "Full_name", None),
    ("FRAME TOP", "Full_name", None),
    ("ROLLER RETARD P ASSY", "Full_name", None),
    ("ROLLER M/P ASSY", "Full_name", None),
    ("PWB APC A ASSY", "Part_name", None),
    ("PWB APC B ASSY", "Part_name", None),
    ("APC ASSY B", "Part_name", None),
    ("PWB TPC ASSY", "Part_name", None),
    ("PWB DLP ASSY", "Part_name", None),
    ("PWB TPC ASSY", "Part_name", None),
    ("CASE DLP", "Part_name", None),
    ("PWB ERASER ASSY", "Part_name", None),
    ("P.W.BOARD ASSY PD B", "Part_name", None),
    ("HOUSING LSU", "Part_name", None),
    ("P.W.BOARD ASSY PD A", "Part_name", None),
    ("PWB APC A ASSY", "Part_name", None),
]

def update_virgo():
    print(f"Preparing to update Virgo with {len(VIRGO_RULES_62)} rules...")
    mgr = BOMFilterManager()
    remote_path, is_remote = mgr.get_active_db_path()
    print(f"Target DB Path: {remote_path} (is_remote={is_remote})")
    
    if remote_path.exists():
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = remote_path.parent / f"ssbom_master.db.bak_{timestamp}"
        print(f"Creating safety backup: {backup_file}")
        shutil.copy2(remote_path, backup_file)
        print("Safety backup created.")

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    note_text = "Cập nhật 62 quy tắc theo yêu cầu người dùng"
    
    records = [
        (
            "Virgo",
            item,
            mode,
            code,
            note_text,
            1,
            now_str,
            now_str,
        )
        for item, mode, code in VIRGO_RULES_62
    ]
    
    def _do_update(conn: sqlite3.Connection):
        del_cur = conn.execute("DELETE FROM bolocbom_rules WHERE LOWER(model_name) = 'virgo';")
        del_count = del_cur.rowcount
        conn.executemany(
            """
            INSERT INTO bolocbom_rules (model_name, item_name, match_mode, part_code, notes, is_active, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?);
            """,
            records,
        )
        return del_count, len(records)
        
    ok, result, used_fallback = mgr._execute_dual_write(_do_update, operation_name="cập nhật 62 quy tắc Virgo")
    if not ok:
        raise RuntimeError(f"Dual-write failed: {result}")
        
    del_count, ins_count = result
    print(f"Deleted {del_count} previous rules and inserted {ins_count} updated rules.")
    print(f"Used fallback: {used_fallback}")

    # Verification
    rules = mgr.get_rules_for_model("Virgo", active_only=True)
    print(f"Verified: {len(rules)} active Virgo rules retrieved from database.")
    assert len(rules) == len(VIRGO_RULES_62)
    for idx, (exp, act) in enumerate(zip(VIRGO_RULES_62, rules), start=1):
        assert act.item_name == exp[0], f"Row {idx} item mismatch: {act.item_name} != {exp[0]}"
        assert act.match_mode == exp[1], f"Row {idx} mode mismatch: {act.match_mode} != {exp[1]}"
        assert act.part_code == exp[2], f"Row {idx} code mismatch: {act.part_code} != {exp[2]}"
        
    print("[SUCCESS] All 62 Virgo rules successfully updated and verified!")

if __name__ == "__main__":
    update_virgo()
