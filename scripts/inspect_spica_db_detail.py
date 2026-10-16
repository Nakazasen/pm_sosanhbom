import sqlite3

conn = sqlite3.connect('data/ssbom_master_cache.db')
cursor = conn.cursor()
cursor.execute("SELECT id, model_name, item_name, match_mode, part_code, notes, is_active FROM bolocbom_rules WHERE LOWER(model_name) = 'spica' ORDER BY id")
rows = cursor.fetchall()
print(f"Total Spica rules in DB: {len(rows)}")

tvc_rules = [r for r in rows if 'TVC1TZ0030' in str(r[5])]
print(f"Rules with TVC1TZ0030 in notes: {len(tvc_rules)}")
for r in tvc_rules:
    print(f"ID={r[0]}, item={r[2]!r}, mode={r[3]!r}, code={r[4]!r}, notes={r[5]!r}")
