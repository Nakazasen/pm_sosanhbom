import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import shutil
import datetime
import sqlite3

from src.core.bom_filter_manager import BOMFilterManager

def sync_spica():
    print("Starting Spica rules synchronization into SQLite...")
    mgr = BOMFilterManager()
    remote_path, is_remote = mgr.get_active_db_path()
    print(f"Active DB path: {remote_path} (is_remote={is_remote})")
    print(f"Cache DB path: {mgr.local_cache_path}")
    
    # 1. Create backups
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    if remote_path.exists():
        backup_remote = remote_path.parent / f"ssbom_master.db.bak_{timestamp}"
        print(f"Creating remote backup: {backup_remote}")
        try:
            shutil.copy2(remote_path, backup_remote)
            print("Remote backup created successfully.")
        except Exception as e:
            print(f"Warning creating remote backup: {e}")

    if mgr.local_cache_path.exists():
        backup_cache = mgr.local_cache_path.parent / f"ssbom_master_cache.db.bak_{timestamp}"
        print(f"Creating cache backup: {backup_cache}")
        shutil.copy2(mgr.local_cache_path, backup_cache)
        print("Cache backup created successfully.")

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _do_update(conn: sqlite3.Connection):
        # Update 3V2YJ01090
        c1 = conn.execute(
            """
            UPDATE bolocbom_rules
            SET notes = 'Tạo tự động từ Cây BOM Mô phỏng (Loại bỏ cả cụm)',
                updated_at = ?
            WHERE LOWER(model_name) = 'spica' AND part_code = '3V2YJ01090';
            """,
            (now_str,)
        )
        cnt_3v = c1.rowcount

        # Update remaining 16 TVC1TZ0030 rules
        c2 = conn.execute(
            """
            UPDATE bolocbom_rules
            SET notes = 'Tạo tự động từ Cây BOM Mô phỏng (Cắt con)',
                updated_at = ?
            WHERE LOWER(model_name) = 'spica' AND notes LIKE '%TVC1TZ0030%';
            """,
            (now_str,)
        )
        cnt_branch = c2.rowcount

        return cnt_3v, cnt_branch

    ok, result, used_fallback = mgr._execute_dual_write(_do_update, operation_name="cập nhật quy tắc lọc Spica")
    if not ok:
        raise RuntimeError(f"Dual-write failed: {result}")

    cnt_3v, cnt_branch = result
    print(f"Updated in DB: 3V2YJ01090={cnt_3v}, branch rules={cnt_branch}, used_fallback={used_fallback}")

    # Also sync to local root ssbom_master.db if it exists
    local_root_db = Path("ssbom_master.db")
    if mgr.local_cache_path.exists():
        print(f"Syncing local cache to {local_root_db.resolve()}...")
        shutil.copy2(mgr.local_cache_path, local_root_db)
        print("Synchronized to local root ssbom_master.db.")

    # Verification
    rules = mgr.get_rules_for_model("Spica", active_only=True)
    print(f"Retrieved {len(rules)} Spica rules from BOMFilterManager.")
    assert len(rules) == 70, f"Expected 70 rules, got {len(rules)}"

    pwb_rule = next((r for r in rules if r.part_code == "3V2YJ01090"), None)
    assert pwb_rule is not None, "3V2YJ01090 rule not found in DB!"
    print(f"3V2YJ01090 record: id={pwb_rule.id}, notes={pwb_rule.notes}")
    assert pwb_rule.notes == "Tạo tự động từ Cây BOM Mô phỏng (Loại bỏ cả cụm)"

    tvc_remaining = [r for r in rules if "TVC1TZ0030" in (r.notes or "")]
    print(f"Remaining TVC1TZ0030 rules count: {len(tvc_remaining)}")
    assert len(tvc_remaining) == 0, f"Found {len(tvc_remaining)} rules still having TVC1TZ0030!"

    model_rules = mgr.get_model_rules_for_pruner("Spica")
    pwb_mr = next((mr for mr in model_rules if mr.part_code == "3V2YJ01090"), None)
    assert pwb_mr is not None, "3V2YJ01090 ModelRule not found!"
    print(f"3V2YJ01090 ModelRule: action={pwb_mr.action}, branch={pwb_mr.parent_part_code}")
    assert pwb_mr.action == "prune_node"
    assert pwb_mr.parent_part_code is None

    # Check the other 16 rules in model_rules
    for part in ["31C1T04120", "31C1T04270", "31C1T04280", "31C0L46060", "TVC1T01010", "TVC1TJ0030"]:
        mr = next((r for r in model_rules if r.part_code == part), None)
        assert mr is not None, f"Rule {part} not found in model_rules!"
        assert mr.action == "prune_children", f"Rule {part} action is {mr.action}, expected prune_children"
        assert mr.parent_part_code is None, f"Rule {part} branch is {mr.parent_part_code}, expected None"

    print("[SUCCESS] Spica rules successfully synchronized and verified across DB and cache!")

if __name__ == "__main__":
    sync_spica()
