import sys
from oletools.olevba import VBA_Parser

sys.stdout.reconfigure(encoding='utf-8')

files = ['form_ssbom.xlsm', 'formnguoidung.xlsm', 'tonghop_new12052026_ma1.xlsm']
for f in files:
    vb = VBA_Parser(f)
    for (fn, sp, vba_name, code) in vb.extract_macros():
        for i, line in enumerate(code.splitlines()):
            if 'R3' in line or 'ws_R3' in line:
                s = line.strip()
                if any(k in s for k in ['Sheets', 'Range', 'Cells', 'Pivot', 'Open', 'Copy', 'Path', 'Dir']):
                    print(f"{f} -> {vba_name}:{i+1} -> {s[:130]}")
