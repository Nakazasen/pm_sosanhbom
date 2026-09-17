' Stream: VBA/md_TaoFileSSB
' File: md_TaoFileSSB.bas

Attribute VB_Name = "md_TaoFileSSB"
Sub TaoFileSSB()
Call Focus(True)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_thdl = wbth.Sheets("tonghopdl")
    Set ws_link = wbth.Sheets("duongdan")
    duongdan = wbth.Path
    duongdan = duongdan & "\"
    tenmay = Right(duongdan, 11)
    tenmay = "BOM_" & Left(tenmay, 10)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    folder_capnhat = duongdan & "capnhat" & "\" & "old"
    If (Not FSO.FolderExists(folder_capnhat)) Then
        FSO.CreateFolder (duongdan & "capnhat")
        FSO.CreateFolder (folder_capnhat)
        Else
    End If
    Call FSO.CopyFile(ws_link.Range("A1").Value & "form_ssbom.xlsm", duongdan)
    OldName = duongdan & "form_ssbom.xlsm"
    NewName = duongdan & tenmay & ".xlsm"
    On Error GoTo loi
    Name OldName As NewName
    Call copydslk_toBom
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=duongdan & tenmay
    Application.DisplayAlerts = True
    Set wb_Bom = Workbooks(tenmay)
    Set ws_Cttt = wb_Bom.Sheets("CTTT")
    Set ws_tongket_CTTT = wb_Bom.Sheets("Tongket")
    wb_Bom.RefreshAll
    wb_Bom.Close True
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " t" & ChrW(7841) & "o xong file BOM c" & ChrW(7847) & "n so s" & ChrW(225) & "nh", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o ho" & ChrW(224) & "n th" & ChrW(224) & "nh", 0 + 64
Exit Sub
loi:
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " t" & ChrW(7891) & "n t" & ChrW(7841) & "i file BOM c" & ChrW(361), , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
    Kill (duongdan & "form_ssbom.xlsm")
  Call Focus(False)
    End Sub
Sub copydslk_toBom()
Call Focus(True)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_thdl = wbth.Sheets("tonghopdl")
    Set ws_tonghopMsi = wbth.Sheets("tonghopMSI_7980_7990")
    Set FSO = CreateObject("Scripting.FileSystemObject")
    duongdan = wbth.Path
    duongdan = duongdan & "\"
    tenmay = Right(duongdan, 11)
    tenmay = "BOM_" & Left(tenmay, 10) & ".xlsm"
    tenfileR3 = Dir(duongdan & "R3*")
    tenfilePLM = Dir(duongdan & "PLM*")
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=duongdan & tenmay
    Application.DisplayAlerts = True
    lr_kq = ws_thdl.Range("C" & Rows.Count).End(xlUp).Row + 1
    Set wb_Bom = Workbooks(tenmay) 'wb_Bom la wb file Bom
    Set ws_Cttt = wb_Bom.Sheets("CTTT")
    Set ws_Plm = wb_Bom.Sheets("PLM")
    Set ws_R3 = wb_Bom.Sheets("R3")
    Set ws_msi = wb_Bom.Sheets("MSI_7980_7990")
    ws_thdl.Range("A3:F" & lr_kq).Copy Destination:=ws_Cttt.Range("A3:F" & lr_kq) 'copy dslk vao file bom
    ws_thdl.Range("G3:G" & lr_kq).Copy Destination:=ws_Cttt.Range("Q3") 'copy dslk vao file bom
    ws_tonghopMsi.Range("A2:L36").Copy Destination:=ws_msi.Range("A2:L36") ' copy msi vao file bom
    ws_tonghopMsi.Range("A49:L73").Copy Destination:=ws_msi.Range("A49:L73") ' copy 7980_7990 vao file bom
    'chuyen du lieu cot C cua sheet CTTT->number
        ws_Cttt.Activate
        ws_Cttt.Range("C3:C" & lr_kq).Select
        Selection.TextToColumns Destination:=Range("C3"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    '---
    If tenfileR3 <> "" Or tenfilePLM <> "" Then 'truong hop khong co file R3 hoac PLM
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
    '
    'FSO.MoveFile Source:=duongdan & tenfileR3, Destination:=duongdan & "capnhat" & "\" & "old" & "\"
    FSO.CopyFile duongdan & tenfileR3, archiveTo & duongdan & "capnhat" & "\" & "old" & "\"
    FSO.DeleteFile duongdan & tenfileR3
    '
    Workbooks.Open Filename:=duongdan & tenfilePLM 'mo file plm
    Set wb_Plm = Workbooks(tenfilePLM)
    lr_kq = wb_Plm.Sheets(1).Range("D" & Rows.Count).End(xlUp).Row + 1
    wb_Plm.Sheets(1).Range("B2:N" & lr_kq).Copy Destination:=ws_Plm.Range("A2:M" & lr_kq)
    wb_Plm.Close False
    '
    'FSO.MoveFile Source:=duongdan & tenfilePLM, Destination:=duongdan & "capnhat" & "\" & "old" & "\"
    FSO.CopyFile duongdan & tenfilePLM, archiveTo & duongdan & "capnhat" & "\" & "old" & "\"
     FSO.DeleteFile duongdan & tenfilePLM
    '
    ws_Plm.Activate
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
        '==============tao pivot table cho plm va r3
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
    Columns("T:T").Select
    Selection.ColumnWidth = 24
    Columns("U:U").Select
    Selection.ColumnWidth = 24
    '-----tao pivot table cho R3
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
'    ws_Plm.Range("$A$1:$N$5603").AutoFilter Field:=5, Criteria1:="<>*0.0**"
'    ActiveSheet.Range("$A$1:$U$5603").AutoFilter Field:=14, Criteria1:="#N/A"
    Unload uf_mainmenu
    wb_Bom.Close True
    End If
    If tenfileR3 = "" Or tenfilePLM = "" Then
    CreateObject("WScript.Shell").Popup "Ch" & ChrW(432) & "a c" & ChrW(243) & " file R3 ho" & ChrW(7863) & "c PLM trong th" & ChrW(432) & " m" & ChrW(7909) & "c so s" & ChrW(225) & "nh BOM", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
    wb_Bom.Close False
    Kill (duongdan & tenmay)
    End
    End If
Call Focus(False)
End Sub



