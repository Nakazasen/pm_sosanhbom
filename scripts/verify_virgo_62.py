import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import sqlite3

# Ensure utf-8 output
sys.stdout.reconfigure(encoding='utf-8')

from scripts.update_virgo_rules_62 import VIRGO_RULES_62

db_path = Path(r"\\fstvn01\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)\Hang muc can luu\Vinh\Pm_sosanhBOM\ssbom_master.db")
conn = sqlite3.connect(str(db_path))
cur = conn.cursor()
cur.execute("SELECT id, item_name, match_mode, part_code FROM bolocbom_rules WHERE UPPER(model_name) = 'VIRGO' ORDER BY id")
db_rows = cur.fetchall()
conn.close()

print(f"Tổng số quy tắc trong CSDL: {len(db_rows)} / Yêu cầu: {len(VIRGO_RULES_62)}")
mismatches = 0
for i, (act, exp) in enumerate(zip(db_rows, VIRGO_RULES_62), start=1):
    _id, item, mode, code = act
    e_item, e_mode, e_code = exp
    if item != e_item or mode != e_mode or code != e_code:
        print(f"Lệch dòng {i}: DB({item}, {mode}, {code}) != YÊU CẦU({e_item}, {e_mode}, {e_code})")
        mismatches += 1

if mismatches == 0 and len(db_rows) == len(VIRGO_RULES_62):
    print("XÁC NHẬN: TẤT CẢ 62 QUY TẮC ĐÃ KHỚP CHÍNH XÁC 100%!")
