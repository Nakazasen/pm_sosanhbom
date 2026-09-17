' Stream: VBA/loc_dulieucanss
' File: loc_dulieucanss.bas

Attribute VB_Name = "loc_dulieucanss"
Sub locdl_ss()
Call Focus(True)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    duongdan = wbth.Path
    duongdan = duongdan & "\"
    tenmay = Dir(duongdan & "BOM*")
    If tenmay = "" Then
    CreateObject("WScript.Shell").Popup "Kh" & ChrW(244) & "ng t" & ChrW(7891) & "n t" & ChrW(7841) & "i file BOM trong th" & ChrW(432) & " m" & ChrW(7909) & "c so s" & ChrW(225) & "nh BOM", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
    End If
    If tenmay <> "" Then
     Application.DisplayAlerts = False
    Workbooks.Open Filename:=duongdan & tenmay
    Application.DisplayAlerts = True
    Set wb_Bom = Workbooks(tenmay) 'wb_Bom la wb file Bom
    Set ws_Cttt = wb_Bom.Sheets("CTTT")
    Set ws_Plm = wb_Bom.Sheets("PLM")
    Set ws_R3 = wb_Bom.Sheets("R3")
    Set ws_cttt_total = wb_Bom.Sheets("CTTT_Total")
    Set ws_tongket_CTTT = wb_Bom.Sheets("Tongket")
    On Error Resume Next
    ws_Cttt.ShowAllData
    ws_Cttt.Range("$A$2:$Y$5000").AutoFilter Field:=18, Criteria1:="=NG", _
        Operator:=xlOr, Criteria2:="=#N/A"

        ws_Plm.Select
        On Error Resume Next
        ws_Plm.ShowAllData
        ws_Plm.Range("A1:S1").Select
        Selection.AutoFilter
'        ws_Plm.Range("A1:S1").AutoFilter Field:=5, Criteria1:="<>0"
        ws_Plm.Range("A1:S1").AutoFilter Field:=14, Criteria1:="#N/A"
'        ws_Plm.Range("A1:S1").AutoFilter Field:=15, Criteria1:="="
'        ws_Plm.Range("A1:S1").AutoFilter Field:=16, Criteria1:="="
        ws_cttt_total.Select
        Columns("A:A").Select
        Selection.ColumnWidth = 24
        Columns("B:B").Select
        Selection.ColumnWidth = 24
    On Error Resume Next
    ws_cttt_total.ShowAllData
    ws_cttt_total.Range("C1:H1").AutoFilter Field:=4, Criteria1:="=NG", _
    Operator:=xlOr, Criteria2:="=#N/A"
'    ws_cttt_total.Range("C1:H1").AutoFilter Field:=6, Criteria1:="="
    ws_Cttt.Activate
    ThisWorkbook.RefreshAll
    wb_Bom.Close True
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " l" & ChrW(7885) & "c d" & ChrW(7919) & " li" & ChrW(7879) & "u c" & ChrW(7847) & "n so s" & ChrW(225) & "nh trong sheet_CTTT_PLM_Total", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o ho" & ChrW(224) & "n th" & ChrW(224) & "nh", 0 + 64
    End If
Call Focus(False)
End Sub
