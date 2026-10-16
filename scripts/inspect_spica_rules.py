import sqlite3
from pathlib import Path

paths = {
    'remote': Path(r'\\fstvn01\Data\00_KDTVN Common(KDTVN共通)\⑤Production Engineering(製造技術)\Hang muc can luu\Vinh\Pm_sosanhBOM\ssbom_master.db'),
    'local_root': Path('ssbom_master.db'),
    'data_cache': Path('data/ssbom_master_cache.db'),
    'data_master': Path('data/ssbom_master.db'),
}

for name, p in paths.items():
    if p.exists():
        try:
            conn = sqlite3.connect(p)
            c = conn.cursor()
            c.execute("SELECT count(*) FROM bolocbom_rules WHERE LOWER(model_name) = 'spica'")
            cnt = c.fetchone()[0]
            c.execute("SELECT id, item_name, part_code, notes FROM bolocbom_rules WHERE LOWER(model_name) = 'spica' AND notes LIKE '%TVC1TZ0030%'")
            tvc_rows = c.fetchall()
            c.execute("SELECT id, item_name, part_code, notes FROM bolocbom_rules WHERE LOWER(model_name) = 'spica' AND part_code = '3V2YJ01090'")
            pwb_rows = c.fetchall()
            print(f'{name} ({p}): total Spica={cnt}, with TVC1TZ0030={len(tvc_rows)}, 3V2YJ01090={pwb_rows}')
            conn.close()
        except Exception as e:
            print(f'{name} error: {e}')
    else:
        print(f'{name} does not exist')
