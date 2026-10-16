import sys
import sqlite3
from pathlib import Path

db_path = Path(r"\\fstvn01\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)\Hang muc can luu\Vinh\Pm_sosanhBOM\ssbom_master.db")
conn = sqlite3.connect(str(db_path))
cur = conn.cursor()
cur.execute("SELECT id, model_name, item_name, match_mode, part_code, notes, is_active, created_at, updated_at FROM bolocbom_rules WHERE UPPER(model_name) = 'VIRGO' ORDER BY id")
rows = cur.fetchall()

with open("scripts/db_virgo_rules_full.txt", "w", encoding="utf-8") as f:
    f.write(f"Total rows in DB for Virgo: {len(rows)}\n")
    for r in rows:
        f.write(f"ID: {r[0]:<4} | Active: {r[6]} | Mode: {r[3]:<10} | Code: {str(r[4] or ''):<12} | Item: {str(r[2]):<42} | Notes: {r[5]}\n")

conn.close()
print("Wrote to scripts/db_virgo_rules_full.txt")
