"""Export real Excel views of the updated BOM (CopyPicture -> chart export; works without a visible desktop)."""
import pathlib
import time

import pythoncom
import win32com.client

p = r"D:\Sandbox\pm_sosanhbom\sample_data\run_feat\Virgo\110C3M3AK0\BOM_110C3M3AK0.xlsm"
out = pathlib.Path(r"D:\Sandbox\pm_sosanhbom\docs\images_feat")
out.mkdir(parents=True, exist_ok=True)
pythoncom.CoInitialize()
x = win32com.client.DispatchEx("Excel.Application")
x.AutomationSecurity = 3
x.DisplayAlerts = False
x.Visible = False
try:
    wb = x.Workbooks.Open(p, UpdateLinks=0, ReadOnly=True)
    jobs = [
        ("Tongket", "A1:L36", "x01_excel_tongket"),
        ("CTTT", "A1:Q40", "x02_excel_cttt_filtered"),
        ("PLM", "A1:T40", "x03_excel_plm"),
        ("R3", "A1:T40", "x04_excel_r3"),
        ("CTTT_Total", "A1:M40", "x05_excel_cttt_total"),
    ]
    for sheet, addr, name in jobs:
        ws = wb.Worksheets(sheet)
        ws.Activate()
        rng = ws.Range(addr)
        hidden = sum(1 for r in range(1, 41) if ws.Rows(r).Hidden)
        rng.CopyPicture(1, 2)  # xlScreen, xlBitmap
        w, h = rng.Width, rng.Height
        co = ws.ChartObjects().Add(0, 0, w, h)
        co.Activate()
        co.Chart.Paste()
        co.Chart.Export(str(out / f"{name}.png"))
        co.Delete()
        print(sheet, addr, "hidden rows in first 40:", hidden, "->", name, flush=True)
    wb.Close(False)
finally:
    x.Quit()
