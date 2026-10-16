import sys
import json
import sqlite3
from pathlib import Path
import openpyxl

if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# 1. Read Virgo from Excel
excel_path = Path(r"\\fstvn01\Data\10_Production Engineering Department(製造技術部)\02.製造技術課\PE Dept\15. FORM（BIEU MAU）-形式\Form_VBA\Form_VBA_SS_BOM\tonghop_new12052026_ma1.xlsm")
wb = openpyxl.load_workbook(str(excel_path), data_only=True)
ws = wb["BolocBom"]

excel_virgo_rules = []
for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
    if row_idx == 1:
        continue
    item_name = row[0]
    match_mode = row[1]
    part_code = row[2]
    
    if item_name is not None and str(item_name).strip():
        item_str = str(item_name).strip()
        mode_str = str(match_mode).strip() if match_mode else "Full_name"
        code_str = str(part_code).strip() if part_code else ""
        excel_virgo_rules.append({
            "row": row_idx,
            "item_name": item_str,
            "match_mode": mode_str,
            "part_code": code_str
        })

print(f"Total Virgo rules in Excel file: {len(excel_virgo_rules)}")

# 2. Read Virgo from SQLite DB (network path)
db_path = Path(r"\\fstvn01\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)\Hang muc can luu\Vinh\Pm_sosanhBOM\ssbom_master.db")
conn = sqlite3.connect(str(db_path))
cur = conn.cursor()
cur.execute("SELECT id, model_name, item_name, match_mode, part_code, notes, is_active FROM bolocbom_rules WHERE UPPER(model_name) = 'VIRGO' ORDER BY id")
db_virgo_rules = []
for r in cur.fetchall():
    db_virgo_rules.append({
        "id": r[0],
        "model_name": r[1],
        "item_name": r[2].strip() if r[2] else "",
        "match_mode": r[3].strip() if r[3] else "Full_name",
        "part_code": r[4].strip() if r[4] else "",
        "notes": r[5] or "",
        "is_active": r[6]
    })
conn.close()

print(f"Total Virgo rules in SQLite DB: {len(db_virgo_rules)}")

# 3. Compare rules
excel_items = {r["item_name"]: r for r in excel_virgo_rules}
db_items = {r["item_name"]: r for r in db_virgo_rules}

in_excel_not_in_db = [r for r in excel_virgo_rules if r["item_name"] not in db_items]
in_db_not_in_excel = [r for r in db_virgo_rules if r["item_name"] not in excel_items]
mode_mismatch = []
for item_name, ex_r in excel_items.items():
    if item_name in db_items:
        db_r = db_items[item_name]
        if ex_r["match_mode"].lower() != db_r["match_mode"].lower() or (ex_r["part_code"] or "") != (db_r["part_code"] or ""):
            mode_mismatch.append({"item": item_name, "excel": ex_r, "db": db_r})

print(f"\nRules in Excel but NOT in DB: {len(in_excel_not_in_db)}")
for r in in_excel_not_in_db:
    print(f"  + Row {r['row']}: item='{r['item_name']}', mode='{r['match_mode']}', code='{r['part_code']}'")

print(f"\nRules in DB but NOT in Excel: {len(in_db_not_in_excel)}")
for r in in_db_not_in_excel:
    print(f"  - ID {r['id']}: item='{r['item_name']}', mode='{r['match_mode']}', active={r['is_active']}, notes='{r['notes']}'")

print(f"\nRules with Mismatch (match_mode or part_code): {len(mode_mismatch)}")
for m in mode_mismatch:
    print(f"  * {m['item']}: Excel(mode={m['excel']['match_mode']}, code={m['excel']['part_code']}) vs DB(mode={m['db']['match_mode']}, code={m['db']['part_code']})")

# Check is_active in DB
inactive_in_db = [r for r in db_virgo_rules if not r["is_active"]]
print(f"\nInactive rules in DB: {len(inactive_in_db)}")
for r in inactive_in_db:
    print(f"  ! ID {r['id']}: '{r['item_name']}' is_active=0")
