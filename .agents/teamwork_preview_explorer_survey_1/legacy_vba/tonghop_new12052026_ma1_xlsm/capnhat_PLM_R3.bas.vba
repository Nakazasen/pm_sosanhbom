' Stream: VBA/capnhat_PLM_R3
' File: capnhat_PLM_R3.bas

Attribute VB_Name = "capnhat_PLM_R3"
Sub capnhat_PLM()
Call Focus(True)
    Const sheetname = "PLM_old"
    Dim ok As Boolean
    ok = False
    Dim ws As Object 'bien ws chay trong file bom de tim ten file plm(old)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    duongdan = wbth.Path
    duongdan = duongdan & "\"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    folder_capnhat = duongdan & "capnhat" & "\" & "old"
    If (Not FSO.FolderExists(folder_capnhat)) Then
        FSO.CreateFolder (duongdan & "capnhat")
        FSO.CreateFolder (folder_capnhat)
    End If
    tenmay = Dir(duongdan & "BOM*")
    If tenmay = "" Then
    CreateObject("WScript.Shell").Popup "Kh" & ChrW(244) & "ng t" & ChrW(7891) & "n t" & ChrW(7841) & "i file BOM trong th" & ChrW(432) & " m" & ChrW(7909) & "c so s" & ChrW(225) & "nh BOM", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
    Exit Sub
    End If
    
    ' Ki?m tra file PLM c? t?n t?i hay kh?ng
    tenfilePLM = Dir(duongdan & "PLM*")
    If tenfilePLM = "" Then
       CreateObject("WScript.Shell").Popup "Kh" & ChrW(244) & "ng t" & ChrW(7891) & "n t" & ChrW(7841) & "i file PLM trong th" & ChrW(432) & " m" & ChrW(7909) & "c BOM", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
        Exit Sub
    End If
   
    If tenmay <> "" Then
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=duongdan & tenmay
    Application.DisplayAlerts = True
    Set wb_Bom = Workbooks(tenmay) 'wb_Bom la wb file Bom
    For Each ws In wb_Bom.Sheets
      If ws.Name = sheetname Then
         ok = True
      Exit For
    End If
    Next
    If ok = False Then
    wb_Bom.Sheets("PLM").Copy Before:=Sheets(6)
    wb_Bom.Sheets("PLM (2)").Select
    wb_Bom.Sheets("PLM (2)").Name = "PLM_old"
    Set ws_Plm = wb_Bom.Sheets("PLM")
    ws_Plm.Activate
    For i = 1 To 19
    ws_Plm.Range("A1:S5000").AutoFilter Field:=i
    Next i
    Range("A2:M5000").Select
    Selection.ClearContents
    Range("O2:Q5000").Select
    Selection.ClearContents
    ActiveSheet.PivotTables("pivot_table_plm").PivotFields("PART CODE"). _
        Orientation = xlHidden
    ActiveSheet.PivotTables("pivot_table_plm").PivotFields("Sum of Q.TY"). _
        Orientation = xlHidden
    tenfilePLM = Dir(duongdan & "PLM*")
    Workbooks.Open Filename:=duongdan & tenfilePLM 'mo file plm
    Set wb_Plm = Workbooks(tenfilePLM)
    lr_kq = wb_Plm.Sheets(1).Range("D" & Rows.Count).End(xlUp).Row + 1
    wb_Plm.Sheets(1).Range("B2:N" & lr_kq).Copy Destination:=ws_Plm.Range("A2:M" & lr_kq)
    wb_Plm.Close False
    'di chuyen plm cu vao folder cap nhat
    'FSO.MoveFile Source:=duongdan & tenfilePLM, Destination:=duongdan & "capnhat" & "\"
     FSO.CopyFile duongdan & tenfilePLM, archiveTo & duongdan & "capnhat" & "\"
     FSO.DeleteFile duongdan & tenfilePLM
    '
    ws_Plm.Activate
    lr_kq = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row + 1
    ws_Plm.Range("O2:Q" & lr_kq).ClearContents 'xoa du lieu giai thich cu trong sheet plm moi
     'chuyen du lieu cot C cua sheet PLM->number
        ws_Plm.Range("C2:C" & lr_kq).Select
        Selection.TextToColumns Destination:=Range("C2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    'chuyen du lieu cot E cua sheet PLM->number
        ws_Plm.Range("E2:E" & lr_kq).Select
        Selection.TextToColumns Destination:=Range("E2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    'chuyen du lieu cot L cua sheet PLM->number
        ws_Plm.Range("L2:L" & lr_kq).Select
        Selection.TextToColumns Destination:=Range("L2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        'goi ham match index de tim kiem gia tri Giai_Thich, Phu_trach, Quan_Ly_check
        Call ham_match_index_mix
'    '==============tao pivot table cho plm
  ws_Plm.Range("T3").Select
    wb_Bom.ShowPivotTableFieldList = True
    With ws_Plm.PivotTables("pivot_table_plm").PivotFields("PART CODE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ws_Plm.PivotTables("pivot_table_plm").PivotFields("Q.TY")
        .Orientation = xlRowField
        .Position = 2
    End With
    ws_Plm.PivotTables("pivot_table_plm").AddDataField ws_Plm.PivotTables _
        ("pivot_table_plm").PivotFields("Q.TY"), "Count of Q.TY", xlCount
    With ws_Plm.PivotTables("pivot_table_plm").PivotFields("Count of Q.TY")
        .Caption = "Sum of Q.TY"
        .Function = xlSum
    End With
    wb_Bom.ShowPivotTableFieldList = False
    ws_Plm.PivotTables("pivot_table_plm").PivotCache.Refresh
    wb_Bom.Close True
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=duongdan & tenmay
    Application.DisplayAlerts = True
    Set wb_Bom = Workbooks(tenmay)
    wb_Bom.RefreshAll
    wb_Bom.Close True
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " c" & ChrW(7853) & "p nh" & ChrW(7853) & "t n" & ChrW(7897) & "i dung PLM m" & ChrW(7899) & "i", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o ho" & ChrW(224) & "n th" & ChrW(224) & "nh", 0 + 64
    End If
    'truong hop da ton tai sheet PLM(old)
    If ok = True Then
    'wb_Bom.Sheets("PLM(old)").Delete
    'MsgBox "da xoa sheet plm old"
    wb_Bom.Sheets("PLM").Copy Before:=Sheets(6)
    Dim slsPlm As Integer
    ' dem so luong sheet de doi ten PLM
        slsPlm = wb_Bom.Sheets.Count
        For u = 1 To slsPlm
        For j = 8 To slsPlm
        On Error Resume Next
        wb_Bom.Sheets("PLM (2)").Select
        wb_Bom.Sheets("PLM (2)").Name = "PLM_old" & "_" & u
        Exit For
        Next j
        Next u
' doi ten sheet
'    wb_Bom.Sheets("PLM (2)").Select
'    wb_Bom.Sheets("PLM (2)").Name = "PLM_old" & "_" & slsPlm
    Set ws_Plm = wb_Bom.Sheets("PLM")
    ws_Plm.Activate
    For i = 1 To 19
    ws_Plm.Range("A2:S5000").AutoFilter Field:=i
    Next i
    Range("A2:M5000").Select
    Selection.ClearContents
    Range("O2:Q5000").Select
    Selection.ClearContents
    ActiveSheet.PivotTables("pivot_table_plm").PivotFields("PART CODE"). _
        Orientation = xlHidden
    ActiveSheet.PivotTables("pivot_table_plm").PivotFields("Sum of Q.TY"). _
        Orientation = xlHidden
    tenfilePLM = Dir(duongdan & "PLM*")
    Workbooks.Open Filename:=duongdan & tenfilePLM 'mo file plm
    Set wb_Plm = Workbooks(tenfilePLM)
    lr_kq = wb_Plm.Sheets(1).Range("D" & Rows.Count).End(xlUp).Row + 1
    wb_Plm.Sheets(1).Range("B2:N" & lr_kq).Copy Destination:=ws_Plm.Range("A2:M" & lr_kq)
    wb_Plm.Close False
    
    'chuyen file plm cu vao thu muc cap nhat
     FSO.CopyFile duongdan & tenfilePLM, archiveTo & duongdan & "capnhat" & "\"
     FSO.DeleteFile duongdan & tenfilePLM
    '
     ws_Plm.Activate
    lr_kq = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row + 1
    ws_Plm.Range("O2:Q" & lr_kq).ClearContents 'xoa du lieu giai thich cu trong sheet plm moi
     'chuyen du lieu cot C cua sheet PLM->number
        ws_Plm.Range("C2:C" & lr_kq).Select
        Selection.TextToColumns Destination:=Range("C2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    'chuyen du lieu cot E cua sheet PLM->number
        ws_Plm.Range("E2:E" & lr_kq).Select
        Selection.TextToColumns Destination:=Range("E2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    'chuyen ,du lieu cot L cua sheet PLM->number
        ws_Plm.Range("L2:L" & lr_kq).Select
        Selection.TextToColumns Destination:=Range("L2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        'goi ham match index de tim kiem gia tri Giai_Thich, Phu_trach, Quan_Ly_check
        Call ham_match_index_mix1
'    '==============tao pivot table cho plm
  ws_Plm.Range("T3").Select
    wb_Bom.ShowPivotTableFieldList = True
    With ws_Plm.PivotTables("pivot_table_plm").PivotFields("PART CODE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ws_Plm.PivotTables("pivot_table_plm").PivotFields("Q.TY")
        .Orientation = xlRowField
        .Position = 2
    End With
    ws_Plm.PivotTables("pivot_table_plm").AddDataField ws_Plm.PivotTables _
        ("pivot_table_plm").PivotFields("Q.TY"), "Count of Q.TY", xlCount
    With ws_Plm.PivotTables("pivot_table_plm").PivotFields("Count of Q.TY")
        .Caption = "Sum of Q.TY"
        .Function = xlSum
    End With
    wb_Bom.ShowPivotTableFieldList = False
    ws_Plm.PivotTables("pivot_table_plm").PivotCache.Refresh
    wb_Bom.Close True
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=duongdan & tenmay
    Application.DisplayAlerts = True
    Set wb_Bom = Workbooks(tenmay)
    wb_Bom.RefreshAll
    wb_Bom.Close True
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " c" & ChrW(7853) & "p nh" & ChrW(7853) & "t n" & ChrW(7897) & "i dung PLM m" & ChrW(7899) & "i", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o ho" & ChrW(224) & "n th" & ChrW(224) & "nh", 0 + 64
    End If
    End If
Call Focus(False)
End Sub

Sub capnhat_R3()
Call Focus(True)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    duongdan = wbth.Path
    duongdan = duongdan & "\"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    folder_capnhat = duongdan & "capnhat" & "\" & "old"
    If (Not FSO.FolderExists(folder_capnhat)) Then
        FSO.CreateFolder (duongdan & "capnhat")
        FSO.CreateFolder (folder_capnhat)
    End If
    tenmay = Dir(duongdan & "BOM*")
    If tenmay = "" Then
    CreateObject("WScript.Shell").Popup "Kh" & ChrW(244) & "ng t" & ChrW(7891) & "n t" & ChrW(7841) & "i file BOM trong th" & ChrW(432) & " m" & ChrW(7909) & "c so s" & ChrW(225) & "nh BOM", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
    Exit Sub
    End If
    
    'kiem tra xem co file r3 trong duong dan khong
    
    tenfileR3 = Dir(duongdan & "R3*")
    If tenfileR3 = "" Then
     CreateObject("WScript.Shell").Popup "Kh" & ChrW(244) & "ng t" & ChrW(7891) & "n t" & ChrW(7841) & "i file R3 trong th" & ChrW(432) & " m" & ChrW(7909) & "c BOM", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
    Exit Sub
    End If
    If tenmay <> "" Then
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=duongdan & tenmay
    Application.DisplayAlerts = True
    Set wb_Bom = Workbooks(tenmay) 'wb_Bom la wb file Bom
    Set ws_R3 = wb_Bom.Sheets("R3")
    ws_R3.Activate
    Range("A1:Q5000").Select
    Selection.ClearContents
    Range("W1:X5000").Select
    Selection.Delete Shift:=xlToLeft
    tenfileR3 = Dir(duongdan & "R3*")
    Workbooks.Open Filename:=duongdan & tenfileR3 'mo file r3
    Set wb_R3 = Workbooks(tenfileR3)
    lr_kq = wb_R3.Sheets(1).Range("E" & Rows.Count).End(xlUp).Row + 1
    'xu ly cot du lieu trong R3 cua Virgo va phan con lai
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
    '---OK
    wb_R3.Sheets(1).Range("A1:Q" & lr_kq).Copy Destination:=ws_R3.Range("A1:Q" & lr_kq)
    wb_R3.Close False
    'di chuyen R3
    'FSO.MoveFile Source:=duongdan & tenfileR3, Destination:=duongdan & "capnhat" & "\"
    FSO.CopyFile duongdan & tenfileR3, archiveTo & duongdan & "capnhat" & "\"
    FSO.DeleteFile duongdan & tenfileR3
    '
'    '==============tao pivot table cho r3
    ws_R3.Select
    ws_R3.Range("W1").Select
    wb_Bom.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "R3!R11C20:R2520C21", Version:=6).CreatePivotTable TableDestination:= _
        "R3!R1C23", TableName:="PivotTable3", DefaultVersion:=6
    ws_R3.Select
    ws_R3.Cells(1, 23).Select
    wb_Bom.ShowPivotTableFieldList = True
    With ws_R3.PivotTables("PivotTable3").PivotFields("PART CODE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ws_R3.PivotTables("PivotTable3").PivotFields("Q.TY")
        .Orientation = xlRowField
        .Position = 2
    End With
    ws_R3.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Q.TY"), "Count of Q.TY", xlCount
    With ws_R3.PivotTables("PivotTable3").PivotFields("Count of Q.TY")
        .Caption = "Sum of Q.TY"
        .Function = xlSum
    End With
    wb_Bom.ShowPivotTableFieldList = False
    wb_Bom.RefreshAll
    wb_Bom.Close True
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=duongdan & tenmay
    Application.DisplayAlerts = True
    Set wb_Bom = Workbooks(tenmay)
    wb_Bom.RefreshAll
    wb_Bom.Close True
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " c" & ChrW(7853) & "p nh" & ChrW(7853) & "t n" & ChrW(7897) & "i dung R3 m" & ChrW(7899) & "i", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o ho" & ChrW(224) & "n th" & ChrW(224) & "nh", 0 + 64
Call Focus(False)
End If
End Sub
Sub ham_match_index_mix()
    Dim i As Integer
    Dim lastrow As Long
    Dim lastrow_nhap As Long
    Dim gttim As Variant, hang_match As Variant
    Dim vung_mlk As Variant, vung_gt As Variant, vung_tennpt As Variant, vung_quanly As Variant
    
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    duongdan = wbth.Path
    duongdan = duongdan & "\"
    tenmay = Dir(duongdan & "BOM*")
    Set wb_Bom = Workbooks(tenmay) 'wb_Bom la wb file Bom
    Set ws_Plm = wb_Bom.Sheets("PLM")
    Set ws_Plm_old = wb_Bom.Sheets("PLM_old")
    lastrow = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row + 1
    lastrow_nhap = ws_Plm_old.Range("A" & Rows.Count).End(xlUp).Row + 1
    '---
    If lastrow < lastrow_nhap Then
        lastrow = lastrow_nhap
    End If
    '---
    Set vung_mlk = ws_Plm_old.Range("C2:C" & lastrow)
    Set vung_gt = ws_Plm_old.Range("O2:O" & lastrow)
    Set vung_tennpt = ws_Plm_old.Range("P2:P" & lastrow)
    Set vung_quanly = ws_Plm_old.Range("Q2:Q" & lastrow)
    i = 2
    Do While Len(ws_Plm.Cells(i, 3).Value) <> 0
    On Error Resume Next
    gttim = ws_Plm.Range("C" & i).Value
    hang_match = Application.WorksheetFunction.Match(gttim, vung_mlk, 0)
    ws_Plm.Range("O" & i).Value = Application.WorksheetFunction.Index(vung_gt, hang_match)
    ws_Plm.Range("P" & i).Value = Application.WorksheetFunction.Index(vung_tennpt, hang_match)
    ws_Plm.Range("Q" & i).Value = Application.WorksheetFunction.Index(vung_quanly, hang_match)
    i = i + 1
    Loop
End Sub
Sub ham_match_index_mix1()
    Dim i As Integer
    Dim lastrow As Long
    Dim lastrow_nhap As Long
    Dim gttim As Variant, hang_match As Variant
    Dim vung_mlk As Variant, vung_gt As Variant, vung_tennpt As Variant, vung_quanly As Variant
    
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    duongdan = wbth.Path
    duongdan = duongdan & "\"
    tenmay = Dir(duongdan & "BOM*")
    Set wb_Bom = Workbooks(tenmay) 'wb_Bom la wb file Bom
    Set ws_Plm = wb_Bom.Sheets("PLM")
    Set ws_Plm_old = wb_Bom.Sheets(6)
    lastrow = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row + 1
    lastrow_nhap = ws_Plm_old.Range("A" & Rows.Count).End(xlUp).Row + 1
    '---
    If lastrow < lastrow_nhap Then
        lastrow = lastrow_nhap
    End If
    '---
    Set vung_mlk = ws_Plm_old.Range("C2:C" & lastrow)
    Set vung_gt = ws_Plm_old.Range("O2:O" & lastrow)
    Set vung_tennpt = ws_Plm_old.Range("P2:P" & lastrow)
    Set vung_quanly = ws_Plm_old.Range("Q2:Q" & lastrow)
    i = 2
    Do While Len(ws_Plm.Cells(i, 3).Value) <> 0
    On Error Resume Next
    gttim = ws_Plm.Range("C" & i).Value
    hang_match = Application.WorksheetFunction.Match(gttim, vung_mlk, 0)
    ws_Plm.Range("O" & i).Value = Application.WorksheetFunction.Index(vung_gt, hang_match)
    ws_Plm.Range("P" & i).Value = Application.WorksheetFunction.Index(vung_tennpt, hang_match)
    ws_Plm.Range("Q" & i).Value = Application.WorksheetFunction.Index(vung_quanly, hang_match)
    i = i + 1
    Loop
End Sub





