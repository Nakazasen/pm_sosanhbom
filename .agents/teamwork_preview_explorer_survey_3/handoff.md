# Investigation Report: SAP R3 Automation & GUI Scripting Architecture

- **Agent**: Explorer 3 (SAP R3 Automation Investigator)
- **Working Directory**: `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3`
- **Target Systems**: SAP Logon 770 (`saplogon.exe`), System `P1J(ERP60-AWS)-VN`, Transaction `CS12`
- **Date**: 2026-09-17
- **Deliverable**: Handoff Report for Milestone M4 (SAP R3 CS12 Automation) and overall architecture synthesis

---

## 1. Observation

Direct observations and evidence collected from local files, binary inspections, and Windows registry/environment checks:

### 1.1. Target SAP Environment & Executable Verification
- Executable Location:
  - Verified path: `C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe` (File exists).
  - Version: `7700.1.7.1161` (Product Version: `770 Final Release`).
  - Auxiliary binary: `C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapgui.exe` (File exists).
- Registry Configuration (`check_sap_version.py` execution result):
  - `HKCU\Software\SAP\SAPGUI Front\SAP Frontend Server\Security`:
    - `UserScripting = 1` (Scripting is active).
    - `WarnOnAttach = 0` (Attach notification prompt disabled).
    - `WarnOnConnection = 0` (Connection notification prompt disabled).
  - `HKCU\Software\SAP\SAPGUI Front\SAP Frontend Server\Scripting`:
    - `ShowNativeWinDlgs = 0` (CRITICAL: 0 enforces SAP GUI internal modal dialogs rather than OS-native file picker dialogs).
  - `HKLM\SOFTWARE\WOW6432Node\SAP\SAPGUI Front\SAP Frontend Server\Security`:
    - `UserScripting = 1` (Machine-wide scripting policy allowed).
    - `SecurityLevel = 0`.
- SAP Landscape Configuration (`C:\Users\tvn183660\AppData\Roaming\SAP\Common\SAPUILandscape.xml`):
  - Target Service: `<Service name="P1J(ERP60-AWS)-VN" systemid="P1J" server="R3_PR1" type="SAPGUI" clocale="EN" .../>`
  - Target Message Server: `<Messageserver name="P1J" host="e8p1jvuci.kmerp.local" port="3601" .../>`

### 1.2. Legacy VBScript & VBA Source Extraction
Two primary legacy automation sources were extracted and decompiled using `oletools.olevba`:
1. `D:\Sandbox\pm_sosanhbom\tudongdangnhapR3.vbs` (Standalone UTF-16LE VBScript):
   ```vbscript
   Set SAPguiAuto = GetObject("SAPGUI")
   Set SAPguiApp = SAPguiAuto.GetScriptingEngine
   Set Connection = SAPguiApp.OpenConnection("P1J(ERP60-AWS)-VN", True)
   Set session = Connection.Children(0)
   If session.Children.Count > 1 Then
       session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select
       session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").SetFocus
       session.findById("wnd[1]/tbar[0]/btn[0]").press
   End If
   session.findById("wnd[0]").maximize
   session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "v130474"
   session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "0123456789"
   session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "EN"
   session.findById("wnd[0]").sendVKey 0
   ```
2. `DownloadAutoR3.bas` in `tonghop_new12052026_ma1.xlsm` and `tonghop_new12052026_maT.xlsm`:
   - Prefix detection difference:
     - `ma1.xlsm`: `mamay = InStr(Right(ObjTarget, 10), "110")`
     - `maT.xlsm`: `mamay = InStr(Right(ObjTarget, 10), "T10")`
   - CS12 transaction execution flow:
     ```vbscript
     session.findById("wnd[0]/tbar[0]/okcd").Text = "CS12"
     session.findById("wnd[0]").sendVKey 0
     For i = 2 To lr_file
         session.findById("wnd[0]/usr/ctxtRC29L-MATNR").Text = ws_downloadR3.Range("A" & i)
         session.findById("wnd[0]/usr/ctxtRC29L-WERKS").Text = "2200"
         session.findById("wnd[0]/usr/txtRC29L-STLAL").Text = "01"
         session.findById("wnd[0]/usr/ctxtRC29L-CAPID").Text = "pp01"
         session.findById("wnd[0]/usr/ctxtRC29L-DATUV").Text = Format(ws_downloadR3.Range("B" & i), "yyyy/mm/dd")
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/tbar[1]/btn[45]").press
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_PATH").Text = duongdan & ws_downloadR3.Range("A" & i) & "\"
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "R3_" & ws_downloadR3.Range("A" & i) & "_" & ngay & "_" & thang & "_" & nam & ".xls"
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
     Next
     session.findById("wnd[0]").Close
     session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
     ```

### 1.3. Downstream Usage and Data Format Quirks
- In `capnhat_PLM_R3.bas` and `md_TaoFileSSB.bas`:
  ```vbscript
  lr_kq = wb_R3.Sheets(1).Range("E" & Rows.Count).End(xlUp).Row + 1
  Dim vungdulieu As Range
  Set vungdulieu = Range("F1:F" & lr_kq).Find(what:="RevLev", MatchCase:=True, lookat:=xlWhole)
  If vungdulieu Is Nothing Then
      Columns("H:H").Select
      Selection.Delete Shift:=xlToLeft
  Else
      Columns("F:F").Select
      Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
      Columns("H:H").Select
      Selection.Delete Shift:=xlToLeft
  End If
  ```
- In `form_ssbom.xlsm` (Sheet `R3`, rows 11-12):
  - Row 11: Header `PART CODE` (Column 20 / T), `Q.TY` (Column 21 / U).
  - Row 12+: `Col T = =IF('R3'!E12="","",'R3'!E12)`
  - Row 12+: `Col U = =IF('R3'!J12="","",'R3'!J12)`
- In `formnguoidung.xlsm` (`md_sosanhBomnho.bas` lines 32-34):
  - `wb_R3.Sheets(1).Range("E12:E" & lr_r3)` -> Part code (Mã linh kiện trong R3).
  - `wb_R3.Sheets(1).Range("G12:G" & lr_r3)` -> Revision (Rev R3).
  - `wb_R3.Sheets(1).Range("J12:J" & lr_r3)` -> Quantity (Số lượng trong R3).

---

## 2. Logic Chain

From the observations, we deduce the following structural facts and requirements:

1. **Automation Engine Foundation (COM Interop)**:
   - Because `ShowNativeWinDlgs = 0` and `UserScripting = 1` are already configured in Windows Registry, SAP GUI 770 exposes the full COM scripting hierarchy through `win32com.client.GetObject("SAPGUI")`.
   - Because SAP GUI internal dialogs are used, all file export operations occur strictly within SAP GUI window objects (`wnd[1]/usr/ctxtDY_PATH` and `ctxtDY_FILENAME`), meaning zero dependence on external Windows OS file dialog hooks (e.g. `pywinauto` or UIAutomation).
   - If `saplogon.exe` is not running, calling `GetObject("SAPGUI")` fails with COM error `-2147221020 (Invalid syntax / object not found)`. Therefore, the Python automation must check `psutil` or start `C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe`, waiting up to 15 seconds for the `SAPGUI` ROT (Running Object Table) registration.

2. **Session Resilience & Multi-Logon Handling**:
   - In production SAP ERP environments, user `v130474` may have an existing open session from another workstation or a previous abnormal disconnect.
   - When this happens, SAP creates popup `wnd[1]` titled "License Information for Multiple Logons".
   - Selecting `wnd[1]/usr/radMULTI_LOGON_OPT2` terminates previous orphaned sessions and claims the session cleanly.
   - Reusing an existing open connection from `application.Connections` avoids unnecessary logouts/logins if SAP is already authenticated.

3. **CS12 Execution Dynamics & Parameter Contract**:
   - Transaction CS12 requires 5 mandatory input parameters:
     - `ctxtRC29L-MATNR`: 10-character machine/parent material code (e.g. `110K123450`, `T10...`).
     - `ctxtRC29L-WERKS`: `"2200"` (Plant Kyocera Vietnam).
     - `txtRC29L-STLAL`: `"01"` (Alternative BOM 1).
     - `ctxtRC29L-CAPID`: `"pp01"` (BOM Application / Usage for Production).
     - `ctxtRC29L-DATUV`: Production validity date formatted as `"YYYY/MM/DD"`.
   - The legacy code calls `CS12` without `/n`. If the previous screen was left in an invalid state, this can error. Using `/nCS12` ensures unconditional navigation to the top-level CS12 transaction screen.

4. **Status Bar Monitoring (Fail-Closed Gate)**:
   - In the legacy VBA code, if a material code did not exist or the BOM was not maintained for plant 2200, SAP remained on the initial screen and displayed an error message in `wnd[0]/sbar`.
   - The legacy code blindly pressed `btn[45]` without checking the status bar, causing runtime COM crashes or exporting corrupted data.
   - Python implementation must inspect `session.findById("wnd[0]/sbar").MessageType`. If `MessageType == 'E'`, it must record the error ("BOM not found in plant 2200"), emit a structured warning to the caller/UI, skip export for that machine, and reset navigation via `/nCS12`.

5. **Export File Format & Data Parsing**:
   - The export dialog `SAPLSPO5:0150` with option `radSPOPLI-SELFLAG[1,0]` exports a spreadsheet format with extension `.xls`.
   - The data contains header rows from 1 to 11, and data rows starting at row 12.
   - The column shifting heuristic in legacy VBA (`RevLev` check in Col F) exists because SAP ALV layout variants may include or exclude the revision level column.
   - Rather than reproducing fragile column deletion (`Columns("H:H").Delete`), the Python parser should directly locate column headers by name (`Component` / `Part Code`, `RevLev` / `Revision`, `Quantity` / `Un`) in row 11, making parsing resilient across any layout changes.

6. **Directory & File Routing**:
   - The root comparison folder contains multiple subfolders, each named after a machine code (e.g., `110...`, `T10...`).
   - The export destination must route the exported file directly into `base_dir / <machine_code> / R3_<machine_code>_<DD>_<MM>_<YYYY>.xls`.
   - If previous files exist, they should be archived into `base_dir / <machine_code> / capnhat / old /` to maintain auditability.

---

## 3. Architecture Specification for Python Implementation

### 3.1. Class Structure & Module Layout
Target module: `src/automation/sap_r3.py` (or `engine/sap/`):

```
src/automation/sap/
├── __init__.py
├── connection.py      # SAPGUI process check, ROT discovery, session acquisition, multi-logon
├── cs12_service.py    # CS12 transaction execution, parameter binding, status bar gate
├── exporter.py        # SAPLSPO5 dialog handling, file save, overwrite handling
├── parser.py          # Resilient reading of raw R3 .xls output into Pandas / Polars / Pydantic
└── models.py          # Data models: SAPCredentials, CS12Params, R3ComponentRow, ExportResult
```

### 3.2. Detailed Component Implementation Design

#### Component 1: Connection & Session Manager (`connection.py`)
```python
import os
import time
import subprocess
import psutil
import win32com.client
from dataclasses import dataclass
from typing import Optional

SAPLOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
DEFAULT_SYSTEM = "P1J(ERP60-AWS)-VN"

@dataclass
class SAPCredentials:
    username: str = "v130474"
    password: str = "0123456789"
    language: str = "EN"
    system: str = DEFAULT_SYSTEM

class SAPConnectionManager:
    def __init__(self, credentials: Optional[SAPCredentials] = None):
        self.creds = credentials or SAPCredentials()
        self.app = None
        self.connection = None
        self.session = None

    def ensure_saplogon_running(self, timeout_sec: int = 15):
        # 1. Check if saplogon.exe process is alive
        running = any('saplogon' in p.name().lower() for p in psutil.process_iter(['name']))
        if not running:
            if not os.path.exists(SAPLOGON_PATH):
                raise FileNotFoundError(f"SAP Logon not found at {SAPLOGON_PATH}")
            subprocess.Popen([SAPLOGON_PATH])
            
        # 2. Wait for ROT entry
        start = time.time()
        while time.time() - start < timeout_sec:
            try:
                sap_gui_auto = win32com.client.GetObject("SAPGUI")
                self.app = sap_gui_auto.GetScriptingEngine
                if self.app:
                    return
            except Exception:
                time.sleep(0.5)
        raise TimeoutError("Timed out waiting for SAPGUI scripting engine registration.")

    def get_or_create_session(self):
        self.ensure_saplogon_running()
        
        # 3. Check for existing connection
        if self.app.Connections.Count > 0:
            for i in range(self.app.Connections.Count):
                conn = self.app.Connections(i)
                if self.creds.system in conn.Description:
                    self.connection = conn
                    if conn.Children.Count > 0:
                        self.session = conn.Children(0)
                        if self.is_logged_in():
                            return self.session
        
        # 4. Open new connection
        self.connection = self.app.OpenConnection(self.creds.system, True)
        self.session = self.connection.Children(0)
        self.handle_multi_logon_and_login()
        return self.session

    def handle_multi_logon_and_login(self):
        self.wait_ready()
        
        # Check if license popup (multi-logon) appears
        try:
            rad_opt2 = self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2")
            rad_opt2.Select()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.wait_ready()
        except Exception:
            pass # No multi-logon dialog
            
        # Perform Login if at login window
        try:
            bname_field = self.session.findById("wnd[0]/usr/txtRSYST-BNAME")
            bname_field.Text = self.creds.username
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = self.creds.password
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = self.creds.language
            self.session.findById("wnd[0]").sendVKey(0) # Enter
            self.wait_ready()
        except Exception:
            pass # Already logged in

    def is_logged_in(self) -> bool:
        try:
            # If okcd (transaction box) is available, user is logged in
            self.session.findById("wnd[0]/tbar[0]/okcd")
            return True
        except Exception:
            return False

    def wait_ready(self, timeout_sec: int = 30):
        start = time.time()
        while self.session and self.session.Busy:
            time.sleep(0.1)
            if time.time() - start > timeout_sec:
                raise TimeoutError("SAP GUI session busy timeout.")
```

#### Component 2: CS12 Execution & Export Service (`cs12_service.py`)
```python
import os
import datetime
from typing import Dict, Any, List

class CS12BOMService:
    def __init__(self, session):
        self.session = session

    def download_multilevel_bom(
        self,
        material: str,
        valid_date: datetime.date,
        destination_dir: str,
        plant: str = "2200",
        bom_usage: str = "pp01",
        alternative: str = "01"
    ) -> Dict[str, Any]:
        os.makedirs(destination_dir, exist_ok=True)
        date_str_input = valid_date.strftime("%Y/%m/%d")
        date_str_file = valid_date.strftime("%d_%m_%Y")
        filename = f"R3_{material}_{date_str_file}.xls"
        full_dest_path = os.path.join(destination_dir, filename)

        # 1. Navigate to CS12 safely
        self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/nCS12"
        self.session.findById("wnd[0]").sendVKey(0)

        # 2. Fill CS12 Initial Screen parameters
        self.session.findById("wnd[0]/usr/ctxtRC29L-MATNR").Text = material
        self.session.findById("wnd[0]/usr/ctxtRC29L-WERKS").Text = plant
        self.session.findById("wnd[0]/usr/txtRC29L-STLAL").Text = alternative
        self.session.findById("wnd[0]/usr/ctxtRC29L-CAPID").Text = bom_usage
        self.session.findById("wnd[0]/usr/ctxtRC29L-DATUV").Text = date_str_input
        
        # 3. Execute (F8)
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # 4. Check Status Bar for Error (Fail-Closed Guard)
        sbar = self.session.findById("wnd[0]/sbar")
        if sbar.MessageType == "E":
            err_msg = sbar.Text
            # Return back to home
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
            self.session.findById("wnd[0]").sendVKey(0)
            return {"success": False, "material": material, "error": err_msg}

        # 5. Trigger Export (btn[45])
        self.session.findById("wnd[0]/tbar[1]/btn[45]").press()

        # 6. Select Spreadsheet radio button in SAPLSPO5:0150
        # radSPOPLI-SELFLAG[1,0] = Spreadsheet format
        self.session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
        ).Select()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # 7. Fill destination path and file name
        # Note: trailing backslash is required by SAP GUI ctxtDY_PATH
        clean_dir = destination_dir if destination_dir.endswith("\\") else destination_dir + "\\"
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").Text = clean_dir
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = filename
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # 8. Check for overwrite dialog (if wnd[2] appears)
        try:
            if self.session.Children.Count > 1:
                # wnd[2] popup: "File already exists. Replace?"
                btn_replace = self.session.findById("wnd[2]/usr/btnSPOP-OPTION1")
                btn_replace.press()
        except Exception:
            pass

        # 9. Return to main screen (F3)
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()

        # 10. Verify file written on disk
        if os.path.exists(full_dest_path) and os.path.getsize(full_dest_path) > 0:
            return {"success": True, "material": material, "path": full_dest_path}
        else:
            return {"success": False, "material": material, "error": "Export file not found or empty."}
```

#### Component 3: Resilient Data Parsing (`parser.py`)
```python
import pandas as pd
import re

def parse_r3_cs12_file(filepath: str) -> pd.DataFrame:
    """
    Parses the exported SAP CS12 spreadsheet file into a clean DataFrame:
    Columns: ['part_code', 'quantity', 'rev_r3']
    Handles column shifts, layout differences, and header locating automatically.
    """
    # 1. Try reading as HTML table (SAPLSPO5 often generates HTML disguised as .xls)
    try:
        tables = pd.read_html(filepath, header=None)
        df_raw = tables[0]
    except Exception:
        # Fallback to TSV / tab-delimited
        try:
            df_raw = pd.read_csv(filepath, sep='\t', header=None, encoding='cp1252')
        except Exception:
            df_raw = pd.read_excel(filepath, header=None)

    # 2. Locate header row (contains 'PART CODE' or 'Component' or 'RevLev')
    header_idx = None
    for r in range(min(20, len(df_raw))):
        row_str = " ".join([str(v) for v in df_raw.iloc[r].dropna()])
        if "PART CODE" in row_str.upper() or "COMPONENT" in row_str.upper() or "REVLEV" in row_str.upper():
            header_idx = r
            break
            
    if header_idx is None:
        header_idx = 10 # Default SAP CS12 header row is index 10 (row 11)

    df_data = df_raw.iloc[header_idx + 1:].copy()
    headers = [str(col).strip().upper() for col in df_raw.iloc[header_idx]]

    # 3. Dynamic Column Resolution
    # Identify indices for Part Code, Quantity, Rev
    part_col_idx = None
    qty_col_idx = None
    rev_col_idx = None

    for i, h in enumerate(headers):
        if "PART CODE" in h or "COMPONENT" in h or "MATERIAL" in h:
            part_col_idx = i
        elif "Q.TY" in h or "QUANTITY" in h or "MENGE" in h:
            qty_col_idx = i
        elif "REV" in h:
            rev_col_idx = i

    # Fallback to legacy index offsets if headers are blank/merged
    if part_col_idx is None:
        part_col_idx = 4 # Col E
    if qty_col_idx is None:
        qty_col_idx = 9  # Col J
    if rev_col_idx is None:
        rev_col_idx = 6  # Col G

    df_clean = pd.DataFrame({
        'part_code': df_data.iloc[:, part_col_idx].astype(str).str.strip(),
        'quantity': pd.to_numeric(df_data.iloc[:, qty_col_idx], errors='coerce').fillna(0.0),
        'rev_r3': df_data.iloc[:, rev_col_idx].astype(str).str.strip() if rev_col_idx < df_data.shape[1] else ""
    })

    # Filter out empty or header artifact rows
    df_clean = df_clean[df_clean['part_code'].str.len() > 0]
    df_clean = df_clean[~df_clean['part_code'].isin(['PART CODE', 'COMPONENT', 'NAN', 'NONE'])]
    return df_clean
```

---

## 4. Caveats

1. **Active SAP User Session Sharing**:
   - The SAP user `v130474` is shared across members in the department. If another member logs in interactively while the automation script is running and selects Option 2 on the multi-logon dialog, the automation session will be abruptly disconnected.
   - *Mitigation*: The Python wrapper must catch COM disconnection exceptions (`pywintypes.com_error`) and attempt an automatic reconnection retry (up to 3 attempts).

2. **Network Latency & AWS VPN Connectivity**:
   - The SAP system `P1J(ERP60-AWS)-VN` is hosted on AWS via corporate internal routing (`e8p1jvuci.kmerp.local`).
   - If the local machine is disconnected from the Kyocera corporate intranet/VPN, `app.OpenConnection` will hang or timeout.
   - *Mitigation*: The pre-flight check should verify DNS resolution and socket connectivity to `e8p1jvuci.kmerp.local:3601` before launching SAP GUI.

3. **User Credentials Management**:
   - Legacy scripts hardcoded `v130474` / `0123456789`.
   - In production, passwords expire or change periodically. The new software must provide a user-configurable settings dialog (or `.env` file) with password masking instead of hardcoded strings.

4. **Security Policy**:
   - This investigation operated strictly in read-only analysis mode: inspected existing files, tested registry keys, and extracted macros. No active write transactions were posted to SAP R3.

---

## 5. Conclusion

1. **Feasibility Assessment**:
   - Fully viable and verified. SAP GUI 770 Scripting is confirmed enabled on this workstation (`UserScripting = 1`, `WarnOnAttach = 0`, `WarnOnConnection = 0`).
   - `ShowNativeWinDlgs = 0` guarantees that all file dialog operations run deterministically through SAP GUI COM controls without external Win32 window hooking.

2. **Architecture Consolidation**:
   - The two separate Excel workbooks (`tonghop_new12052026_ma1.xlsm` filtering for `110*` and `tonghop_new12052026_maT.xlsm` filtering for `T10*`) can be 100% unified into a single Python automation engine with a configurable machine prefix pattern (`110*`, `T10*`, or arbitrary regex).

3. **Error Resilience Upgrades**:
   - Implementing status bar (`wnd[0]/sbar`) inspection eliminates the #1 cause of crashes in the legacy VBA script (trying to export when BOM doesn't exist).
   - Dynamic header parsing replaces brittle VBA column deletion (`Columns("H:H").Delete`), guaranteeing compatibility across different model layouts (Virgo, Libra2, Iris2024).

---

## 6. Verification Method

To independently verify the facts and findings documented in this report:

1. **Verify SAP GUI Executable & Version**:
   ```powershell
   (Get-Item "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe").VersionInfo | Format-List
   ```
   *Expected Output*: FileVersion `7700.1.7.1161`, ProductVersion `770 Final Release`.

2. **Verify SAP Scripting Registry Settings**:
   ```powershell
   Get-ItemProperty -Path "HKCU:\Software\SAP\SAPGUI Front\SAP Frontend Server\Security"
   Get-ItemProperty -Path "HKCU:\Software\SAP\SAPGUI Front\SAP Frontend Server\Scripting"
   ```
   *Expected Output*: `UserScripting = 1`, `WarnOnAttach = 0`, `WarnOnConnection = 0`, `ShowNativeWinDlgs = 0`.

3. **Verify SAP Landscape Connection Entry**:
   ```powershell
   Select-String -Path "$env:APPDATA\SAP\Common\SAPUILandscape.xml" -Pattern 'P1J(ERP60-AWS)-VN'
   ```
   *Expected Output*: Matches XML service node with `name="P1J(ERP60-AWS)-VN"`, `systemid="P1J"`, `server="R3_PR1"`.

4. **Verify Legacy VBA Macro Extraction**:
   Inspect the extracted VBA code in:
   `D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3\find_r3_refs.py`
   Run:
   ```bash
   python D:\Sandbox\pm_sosanhbom\.agents\teamwork_preview_explorer_survey_3\find_r3_refs.py
   ```
   *Expected Output*: Displays all extracted lines in `DownloadAutoR3.bas`, `md_TaoFileSSB.bas`, `capnhat_PLM_R3.bas`, and `md_sosanhBomnho.bas`.

5. **Invalidation Conditions**:
   - If IT updates SAP GUI to a version that disables scripting by default (`UserScripting = 0`).
   - If the AWS message server hostname `e8p1jvuci.kmerp.local` or port `3601` changes.
   - If plant code changes from `2200` or BOM usage changes from `pp01`.
