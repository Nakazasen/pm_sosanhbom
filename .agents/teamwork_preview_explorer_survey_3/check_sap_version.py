import os
import winreg
import subprocess

exe_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
print("Checking version of:", exe_path)

res = subprocess.run(['powershell', '-NoProfile', f'(Get-Item "{exe_path}").VersionInfo.FileVersion'], capture_output=True, text=True)
print("FileVersion:", res.stdout.strip())

res = subprocess.run(['powershell', '-NoProfile', f'(Get-Item "{exe_path}").VersionInfo.ProductVersion'], capture_output=True, text=True)
print("ProductVersion:", res.stdout.strip())

reg_paths = [
    (winreg.HKEY_CURRENT_USER, r"Software\SAP\SAPGUI Front\SAP Frontend Server\Scripting"),
    (winreg.HKEY_CURRENT_USER, r"Software\SAP\SAPGUI Front\SAP Frontend Server\Security"),
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\SAP\SAPGUI Front\SAP Frontend Server\Scripting"),
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\SAP\SAPGUI Front\SAP Frontend Server\Security"),
]

for hkey, subkey in reg_paths:
    hname = "HKCU" if hkey == winreg.HKEY_CURRENT_USER else "HKLM"
    try:
        k = winreg.OpenKey(hkey, subkey)
        print(f"\nRegistry Key: {hname}\\{subkey}")
        i = 0
        while True:
            try:
                name, val, typ = winreg.EnumValue(k, i)
                print(f"  {name} = {val} ({typ})")
                i += 1
            except OSError:
                break
        winreg.CloseKey(k)
    except FileNotFoundError:
        print(f"\nRegistry Key not found: {hname}\\{subkey}")
