Attribute VB_Name = "tonghopdl"
Sub kiemtra_trangthainhapdlok_ng()
Call Focus(True)
'1 khai bao bien
Dim wb_KQ As Workbook
Dim wb_file As Workbook
Dim myPath As String
Dim lr_kq As Long
Dim lr_file As Long
Dim tenngpt As String
Dim FSO As Object
Dim myfolderPath As String
Dim myfile As Object
Dim files As Object

'2 gan gia tri cho bien
Set FSO = CreateObject("Scripting.FileSystemObject")
Set wb_KQ = Workbooks(ThisWorkbook.Name)
myPath = wb_KQ.Path
myfolderPath = myPath & "\"
Set myfile = FSO.GetFolder(myfolderPath)
'4 su dung vong lap for each de tong hop du lieu

For Each files In myfile.files
'  If files.Name <> wb_KQ.Name And files.Name <> "~$" & wb_KQ.Name Then
If files.Name <> Dir(myfolderPath & "*tonghop*") And files.Name <> Dir(myfolderPath & "BOM*") _
And files.Name <> Dir(myfolderPath & "R3*") And files.Name <> Dir(myfolderPath & "PLM*") _
And files.Name <> Dir(myfolderPath & "~$*", vbHidden) Then
    On Error Resume Next
  Application.DisplayAlerts = False
    Workbooks.Open Filename:=myfolderPath & "\" & files.Name
  Application.DisplayAlerts = True
    Set wb_file = Workbooks(files.Name)
    Set ws_Cttt_pt = wb_file.Sheets("CTTT")
    If ws_Cttt_pt.Range("Q2") = "OK" Then
    Else
    tenngpt = files.Name & "; " & tenngpt
    End If
    wb_file.Close False
  End If
Next
If tenngpt <> "" Then
CreateObject("WScript.Shell").Popup tenngpt & "ch" & ChrW(432) & "a nh" & ChrW(7853) & "p danh s" & ChrW(225) & "ch linh ki" & ChrW(7879) & "n", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
End If
If tenngpt = "" Then
Call tonghop_dulieu
End If
Call Focus(False)
End Sub


Sub tonghop_dulieu()
Call Focus(True)
'1 khai bao bien
'->bien da duoc khai bao trong Md_khaibaobienlocal
'2 gan gia tri cho bien
Dim dc As Long
Set FSO = CreateObject("Scripting.FileSystemObject")
Set wb_KQ = Workbooks(ThisWorkbook.Name)
Set ws_thlabel_7980 = wb_KQ.Worksheets("tonghopMSI_7980_7990") 'ws 7980 cua file th
duongdan = wb_KQ.Path
Set myfile = FSO.GetFolder(duongdan)
duongdan = duongdan & "\"
folder_phutrach = duongdan & "phutrach"
    If (Not FSO.FolderExists(folder_phutrach)) Then
        FSO.CreateFolder (folder_phutrach)
    End If
dc = wb_KQ.Sheets("tonghopdl").Range("C" & Rows.Count).End(xlUp).Row
If dc > 3 Then
CreateObject("WScript.Shell").Popup "H" & ChrW(227) & "y x" & ChrW(243) & "a d" & ChrW(7919) & " li" & ChrW(7879) & "u c" & ChrW(361) & " tr" & ChrW(432) & ChrW(7899) & "c khi copy m" & ChrW(7899) & "i, ch" & ChrW(7885) & "n 1.X" & ChrW(243) & "a d" & ChrW(7919) & " li" & ChrW(7879) & "u c" & ChrW(361) & " trong file t" & ChrW(7893) & "ng h" & ChrW(7907) & "p", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
End If
If dc < 3 Then
'3 su dung vong lap for each de tong hop du lieu
For Each files In myfile.files
If files.Name <> Dir(duongdan & "*tonghop*") And files.Name <> Dir(duongdan & "BOM*") _
And files.Name <> Dir(duongdan & "R3*") And files.Name <> Dir(duongdan & "PLM*") _
And files.Name <> Dir(duongdan & "~$*", vbHidden) Then
    On Error Resume Next
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=duongdan & files.Name
    Application.DisplayAlerts = True
    Set wb_file = Workbooks(files.Name)
    Set ws_label_7980 = wb_file.Worksheets("Label_7980_7990")
    '
    On Error Resume Next
    wb_file.Worksheets("CTTT").ShowAllData
    '
    lr_file = wb_file.Worksheets(1).Range("C" & Rows.Count).End(xlUp).Row
    'tim dong cuoi cua sheet label_7980_7990 cua wb nguoi pt
    lr_label_7980 = ws_label_7980.Range("C" & Rows.Count).End(xlUp).Row
    lr_label_7990 = ws_label_7980.Range("I" & Rows.Count).End(xlUp).Row
    wb_file.Worksheets(1).Range("F3:F" & lr_file) = Split(files.Name, ".xlsm") ' dien ten nguoi phu trach
    'neu dong cuoi >=2 thi copy tu dong 2 den dong cuoi vao file th 7980
    If lr_label_7980 >= 2 Or lr_label_7990 >= 2 Then
    'tim dong cuoi cua cot C sheet tonghopMSI_7980_7990 cua wb thdl
'    lr_th_label7980 = ws_thlabel_7980.Range("C" & Rows.Count).End(xlUp).Row
'        ws_label_7980.Range("A2:F" & lr_label_7980).Copy _
'               Destination:=ws_thlabel_7980.Range("A" & lr_th_label7980 + 1)
    Dim qllbth As Integer
    Dim qllb As Integer
    qllb = 2
    For qllbth = 49 To 73
    If wb_file.Worksheets("Label_7980_7990").Range("A" & qllb) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("A" & qllbth) = wb_file.Worksheets("Label_7980_7990").Range("A" & qllb)
    End If
    If wb_file.Worksheets("Label_7980_7990").Range("B" & qllb) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("B" & qllbth) = wb_file.Worksheets("Label_7980_7990").Range("B" & qllb)
    End If
    If wb_file.Worksheets("Label_7980_7990").Range("C" & qllb) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("C" & qllbth) = wb_file.Worksheets("Label_7980_7990").Range("C" & qllb)
    End If
    If wb_file.Worksheets("Label_7980_7990").Range("D" & qllb) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("D" & qllbth) = wb_file.Worksheets("Label_7980_7990").Range("D" & qllb)
    End If
    If wb_file.Worksheets("Label_7980_7990").Range("E" & qllb) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("E" & qllbth) = wb_file.Worksheets("Label_7980_7990").Range("E" & qllb)
    End If
    If wb_file.Worksheets("Label_7980_7990").Range("F" & qllb) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("F" & qllbth) = wb_file.Worksheets("Label_7980_7990").Range("F" & qllb)
    End If
    If wb_file.Worksheets("Label_7980_7990").Range("G" & qllb) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("G" & qllbth) = wb_file.Worksheets("Label_7980_7990").Range("G" & qllb)
    End If
    If wb_file.Worksheets("Label_7980_7990").Range("H" & qllb) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("H" & qllbth) = wb_file.Worksheets("Label_7980_7990").Range("H" & qllb)
    End If
    If wb_file.Worksheets("Label_7980_7990").Range("I" & qllb) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("I" & qllbth) = wb_file.Worksheets("Label_7980_7990").Range("I" & qllb)
    End If
    If wb_file.Worksheets("Label_7980_7990").Range("J" & qllb) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("J" & qllbth) = wb_file.Worksheets("Label_7980_7990").Range("J" & qllb)
    End If
    If wb_file.Worksheets("Label_7980_7990").Range("L" & qllb) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("L" & qllbth) = wb_file.Worksheets("Label_7980_7990").Range("L" & qllb)
    End If
    qllb = qllb + 1
    Next
    End If
    
    If i = 0 Then
    lr_kq = 3
    wb_file.Worksheets("CTTT").Range("A3:F" & lr_file).Copy _
               Destination:=wb_KQ.Worksheets("tonghopdl").Range("A" & lr_kq)
    wb_file.Worksheets("CTTT").Range("O3:O" & lr_file).Copy _
               Destination:=wb_KQ.Worksheets("tonghopdl").Range("G" & lr_kq)
    Else
    lr_kq = wb_KQ.Worksheets("tonghopdl").Range("A" & Rows.Count).End(xlUp).Row + 1
    wb_file.Worksheets("CTTT").Range("A3:F" & lr_file).Copy _
               Destination:=wb_KQ.Worksheets("tonghopdl").Range("A" & lr_kq)
    wb_file.Worksheets("CTTT").Range("O3:O" & lr_file).Copy _
               Destination:=wb_KQ.Worksheets("tonghopdl").Range("G" & lr_kq)
    End If
    For u = 2 To 36
    If wb_file.Worksheets("MSI").Range("A" & u) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("A" & u) = wb_file.Worksheets("MSI").Range("A" & u)
    End If
    If wb_file.Worksheets("MSI").Range("B" & u) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("B" & u) = wb_file.Worksheets("MSI").Range("B" & u)
    End If
    If wb_file.Worksheets("MSI").Range("D" & u) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("D" & u) = wb_file.Worksheets("MSI").Range("D" & u)
    End If
    If wb_file.Worksheets("MSI").Range("E" & u) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("E" & u) = wb_file.Worksheets("MSI").Range("E" & u)
    End If
    If wb_file.Worksheets("MSI").Range("F" & u) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("F" & u) = wb_file.Worksheets("MSI").Range("F" & u)
    End If
    If wb_file.Worksheets("MSI").Range("G" & u) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("G" & u) = wb_file.Worksheets("MSI").Range("G" & u)
    End If
    If wb_file.Worksheets("MSI").Range("H" & u) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("H" & u) = wb_file.Worksheets("MSI").Range("H" & u)
    End If
    If wb_file.Worksheets("MSI").Range("I" & u) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("I" & u) = wb_file.Worksheets("MSI").Range("I" & u)
    End If
    If wb_file.Worksheets("MSI").Range("J" & u) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("J" & u) = wb_file.Worksheets("MSI").Range("J" & u)
    End If
    If wb_file.Worksheets("MSI").Range("L" & u) <> "" Then
    wb_KQ.Worksheets("tonghopMSI_7980_7990").Range("L" & u) = wb_file.Worksheets("MSI").Range("L" & u)
    End If
    Next
    wb_file.Close False
    FSO.CopyFile duongdan & files.Name, archiveTo & duongdan & "phutrach" & "\", True
    FSO.DeleteFile duongdan & files.Name
  End If
  i = i + 1
Next
CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " c" & ChrW(7853) & "p nh" & ChrW(7853) & "t xong danh s" & ChrW(225) & "ch linh ki" & ChrW(7879) & "n v" & ChrW(224) & "o file t" & ChrW(7893) & "ng h" & ChrW(7907) & "p", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o c" & ChrW(7853) & "p nh" & ChrW(7853) & "t", 0 + 64
End If
Call Focus(False)
End Sub




