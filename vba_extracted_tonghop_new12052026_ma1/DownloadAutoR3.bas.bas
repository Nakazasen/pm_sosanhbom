Attribute VB_Name = "DownloadAutoR3"
Sub Nhapduongdan()
    'Declarations.
    Dim ObjFileSystem As Object
    Dim ObjFolder As Object
    Dim ObjTarget As Object
    Dim ObjSubFolder As Object
    Dim mamay As String
    '---
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_downloadR3 = wbth.Worksheets("Download_R3_tudong")
    duongdan = wbth.Path & "\"
    lr_file = ws_downloadR3.Range("A" & Rows.Count).End(xlUp).Row
    'cai dat bien.
    Set ObjFileSystem = CreateObject("Scripting.FileSystemObject")
    Set ObjFolder = ObjFileSystem.GetFolder(duongdan)
    Set ObjSubFolder = ObjFolder.SubFolders
    For Each ObjTarget In ObjSubFolder
            mamay = InStr(Right(ObjTarget, 10), "110")
            If mamay = 1 Then
                ws_downloadR3.Range("A" & lr_file + 1) = Right(ObjTarget, 10)
                lr_file = ws_downloadR3.Range("A" & Rows.Count).End(xlUp).Row
            Else
            End If
    Next
    ws_downloadR3.Activate
End Sub

Sub Nhapduongdan_auto()
 Call Focus(True)
    'Declarations.
    Dim ObjFileSystem As Object
    Dim ObjFolder As Object
    Dim ObjTarget As Object
    Dim ObjSubFolder As Object
    Dim mamay As String
    '---
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_downloadR3 = wbth.Worksheets("Download_R3_tudong")
    duongdan = wbth.Path & "\"
    lr_file = ws_downloadR3.Range("A" & Rows.Count).End(xlUp).Row
    'cai dat bien.
    Set ObjFileSystem = CreateObject("Scripting.FileSystemObject")
    Set ObjFolder = ObjFileSystem.GetFolder(duongdan)
    Set ObjSubFolder = ObjFolder.SubFolders
    For Each ObjTarget In ObjSubFolder
            mamay = InStr(Right(ObjTarget, 10), "110")
            If mamay = 1 Then
                ws_downloadR3.Range("A" & lr_file + 1) = Right(ObjTarget, 10)
                lr_file = ws_downloadR3.Range("A" & Rows.Count).End(xlUp).Row
            Else
            End If
    Next
    Call sap_2
Call Focus(False)
End Sub




' Dang nhap R3
Sub sap_1()
Dim Application As Variant
Dim homnay As String
Dim wb As Workbook
Dim ws As Worksheet
Dim namewb As String
Dim ngay As String
Dim thang As String
Dim nam As String
    homnay = Format(Date, "ddmmyyyy")
    ngay = Left(homnay, 2)
    thang = Mid(homnay, 3, 2)
    nam = Right(homnay, 4)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_downloadR3 = wbth.Worksheets("Download_R3_tudong")
    '--
    duongdan = wbth.Path & "\"
    '--
lr_file = ws_downloadR3.Range("A" & Rows.Count).End(xlUp).Row
If Not IsObject(SAPguiApp) Then
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
On Error GoTo 0
Set WshShell = CreateObject("WScript.Shell")
Set oExec = WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
On Error Resume Next
Do While Not WshShell.AppActivate("SAP Logon 770")
WScript.Sleep 100
Loop
Set SapGuiAuto = GetObject("SAPGUI")
Else
On Error GoTo 0
End If
Set SAPguiApp = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
Set Connection = SAPguiApp.OpenConnection("P1J(ERP60-AWS)-VN", True)
End If
If Not IsObject(session) Then
Set session = Connection.Children(0)
End If

If session.Children.Count > 1 Then
session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select
session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
End If

If Not IsObject(Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = Application.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If
homnay = Format(Date, "dd/mm/yy")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "v130474"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "0123456789"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").SetFocus
session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 10
session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "EN"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/okcd").Text = "CS12"
session.findById("wnd[0]").sendVKey 0
For i = 2 To lr_file
session.findById("wnd[0]/usr/ctxtRC29L-MATNR").Text = ws_downloadR3.Range("A" & i)
session.findById("wnd[0]/usr/ctxtRC29L-WERKS").Text = "2200"
session.findById("wnd[0]/usr/txtRC29L-STLAL").Text = "01"
session.findById("wnd[0]/usr/ctxtRC29L-CAPID").Text = "pp01"
session.findById("wnd[0]/usr/ctxtRC29L-DATUV").Text = Format(ws_downloadR3.Range("B" & i), "yyyy/mm/dd")
session.findById("wnd[0]/usr/ctxtRC29L-DATUV").SetFocus
session.findById("wnd[0]/usr/ctxtRC29L-DATUV").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = duongdan & ws_downloadR3.Range("A" & i) & "\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "R3_" & ws_downloadR3.Range("A" & i) & "_" & ngay & "_" & thang & "_" & nam & ".xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
Next
session.findById("wnd[0]").Close
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
ws_downloadR3.Range("A2:B" & lr_file).ClearContents
CreateObject("WScript.Shell").Popup "Bom R3 " & ChrW(273) & ChrW(227) & " t" & ChrW(7843) & "i v" & ChrW(7873) & " xong!", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o th" & ChrW(224) & "nh c" & ChrW(244) & "ng", 0 + 64
End Sub


' Dang nhap R3
Sub sap_2()
Dim Application As Variant
Dim homnay As String
Dim wb As Workbook
Dim ws As Worksheet
Dim namewb As String
Dim ngay As String
Dim thang As String
Dim nam As String
    homnay = Format(Date, "ddmmyyyy")
    ngay = Left(homnay, 2)
    thang = Mid(homnay, 3, 2)
    nam = Right(homnay, 4)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_downloadR3 = wbth.Worksheets("Download_R3_tudong")
    '--
    duongdan = wbth.Path & "\"
    '--
lr_file = ws_downloadR3.Range("A" & Rows.Count).End(xlUp).Row
If Not IsObject(SAPguiApp) Then
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
On Error GoTo 0
Set WshShell = CreateObject("WScript.Shell")
Set oExec = WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
On Error Resume Next
Do While Not WshShell.AppActivate("SAP Logon 770")
WScript.Sleep 100
Loop
Set SapGuiAuto = GetObject("SAPGUI")
Else
On Error GoTo 0
End If
Set SAPguiApp = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
Set Connection = SAPguiApp.OpenConnection("P1J(ERP60-AWS)-VN", True)
End If
If Not IsObject(session) Then
Set session = Connection.Children(0)
End If

If session.Children.Count > 1 Then
session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select
session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
End If

If Not IsObject(Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = Application.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If
homnay = Format(Date, "dd/mm/yy")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "v130474"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "0123456789"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").SetFocus
session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 10
session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "EN"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/okcd").Text = "CS12"
session.findById("wnd[0]").sendVKey 0
For i = 2 To lr_file
session.findById("wnd[0]/usr/ctxtRC29L-MATNR").Text = ws_downloadR3.Range("A" & i)
session.findById("wnd[0]/usr/ctxtRC29L-WERKS").Text = "2200"
session.findById("wnd[0]/usr/txtRC29L-STLAL").Text = "01"
session.findById("wnd[0]/usr/ctxtRC29L-CAPID").Text = "pp01"
session.findById("wnd[0]/usr/ctxtRC29L-DATUV").Text = Format(ws_downloadR3.Range("B2"), "yyyy/mm/dd")
session.findById("wnd[0]/usr/ctxtRC29L-DATUV").SetFocus
session.findById("wnd[0]/usr/ctxtRC29L-DATUV").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = duongdan & ws_downloadR3.Range("A" & i) & "\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "R3_" & ws_downloadR3.Range("A" & i) & "_" & ngay & "_" & thang & "_" & nam & ".xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
Next
session.findById("wnd[0]").Close
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
ws_downloadR3.Range("A2:B" & lr_file).ClearContents
CreateObject("WScript.Shell").Popup "Bom R3 " & ChrW(273) & ChrW(227) & " t" & ChrW(7843) & "i v" & ChrW(7873) & " xong!", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o th" & ChrW(224) & "nh c" & ChrW(244) & "ng", 0 + 64
End Sub

