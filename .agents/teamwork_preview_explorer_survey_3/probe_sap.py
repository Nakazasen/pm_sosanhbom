import os
import sys
import psutil

print("=== Checking SAP GUI executables ===")
candidate_paths = [
    r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",
    r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe",
    r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapgui.exe",
    r"C:\Program Files\SAP\FrontEnd\SAPgui\sapgui.exe"
]

for p in candidate_paths:
    exists = os.path.exists(p)
    print(f"  {p}: {'EXISTS' if exists else 'NOT FOUND'}")

print("\n=== Checking running SAP processes ===")
sap_procs = []
for proc in psutil.process_iter(['pid', 'name', 'exe']):
    try:
        name = proc.info['name'] or ""
        if 'sap' in name.lower():
            sap_procs.append(proc.info)
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        pass

if sap_procs:
    for sp in sap_procs:
        print(f"  PID {sp['pid']}: {sp['name']} ({sp['exe']})")
else:
    print("  No active SAP processes found.")

print("\n=== Checking COM Object GetObject('SAPGUI') ===")
try:
    import win32com.client
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    print(f"  GetObject('SAPGUI') returned: {sap_gui_auto}")
    app = sap_gui_auto.GetScriptingEngine
    print(f"  GetScriptingEngine returned: {app}")
    print(f"  Connections count: {app.Connections.Count}")
except Exception as e:
    print(f"  COM check error (expected if saplogon not running): {e}")
