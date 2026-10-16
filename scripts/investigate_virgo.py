import sys
import sqlite3
from pathlib import Path

def investigate():
    print("=== INVESTIGATING SQLITE DATABASES ===")
    db_candidates = [
        Path("config/ssbom_master.db"),
        Path("ssbom_master.db"),
        Path(r"\\fstvn01\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)\Hang muc can luu\Vinh\Pm_sosanhBOM\ssbom_master.db")
    ]
    
    for db_path in db_candidates:
        print(f"\nChecking: {db_path}")
        if not db_path.exists():
            print("  File not found or not accessible.")
            continue
        try:
            conn = sqlite3.connect(str(db_path))
            cur = conn.cursor()
            cur.execute("SELECT DISTINCT model_name, count(*) FROM bolocbom_rules GROUP BY model_name")
            rows = cur.fetchall()
            print("  Models found in bolocbom_rules:")
            for m, count in rows:
                print(f"    - {m}: {count} rules")
                
            cur.execute("SELECT id, model_name, item_name, match_mode, part_code, notes, is_active FROM bolocbom_rules WHERE UPPER(model_name) LIKE '%VIRGO%'")
            virgo_rules = cur.fetchall()
            print(f"  Virgo rules found: {len(virgo_rules)}")
            for r in virgo_rules[:10]:
                print(f"    {r}")
            conn.close()
        except Exception as e:
            print(f"  Error: {e}")

if __name__ == "__main__":
    investigate()
