' Stream: VBA/Locdl_focus
' File: Locdl_focus.bas

Attribute VB_Name = "Locdl_focus"
Sub locdl_ss_cttt()
Call Focus(True)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    duongdan = wbth.Path
    Set ws_Cttt = wbth.Sheets("CTTT")
    Set ws_Plm = wbth.Sheets("PLM")
    Set ws_R3 = wbth.Sheets("R3")
    Set ws_cttt_total = wbth.Sheets("CTTT_Total")
    Set ws_Tongket = wbth.Sheets("Tongket")
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
'    wb_Bom.Close True
    On Error Resume Next
    ThisWorkbook.RefreshAll
    ws_Cttt.Activate
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " l" & ChrW(7885) & "c d" & ChrW(7919) & " li" & ChrW(7879) & "u c" & ChrW(7847) & "n so s" & ChrW(225) & "nh trong sheet_CTTT_BOM_Total", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o ho" & ChrW(224) & "n th" & ChrW(224) & "nh", 0 + 64
Call Focus(False)
End Sub

Sub locdl_ss_plm()
Call Focus(True)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    duongdan = wbth.Path
    Set ws_Cttt = wbth.Sheets("CTTT")
    Set ws_Plm = wbth.Sheets("PLM")
    Set ws_R3 = wbth.Sheets("R3")
    Set ws_cttt_total = wbth.Sheets("CTTT_Total")
    Set ws_Tongket = wbth.Sheets("Tongket")
    On Error Resume Next
    ws_Cttt.ShowAllData
    ws_Cttt.Range("$A$2:$Y$5000").AutoFilter Field:=18, Criteria1:="=NG", _
        Operator:=xlOr, Criteria2:="=#N/A"
'        ws_Cttt.Range("$A$2:$Y$5000").AutoFilter Field:=17, Criteria1:="="
'        ws_Cttt.Range("$A$2:$Y$5000").AutoFilter Field:=19, Criteria1:="="
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
     On Error Resume Next
     ws_cttt_total.ShowAllData
        Columns("A:A").Select
        Selection.ColumnWidth = 24
        Columns("B:B").Select
        Selection.ColumnWidth = 24
    'ws_cttt_total.ShowAllData
    Range("C1:H1").AutoFilter Field:=4, Criteria1:="=NG", _
    Operator:=xlOr, Criteria2:="=#N/A"
'    Range("C1:H1").AutoFilter Field:=6, Criteria1:="="
'    wb_Bom.Close True
  On Error Resume Next
    ThisWorkbook.RefreshAll
    ws_Plm.Activate
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " l" & ChrW(7885) & "c d" & ChrW(7919) & " li" & ChrW(7879) & "u c" & ChrW(7847) & "n so s" & ChrW(225) & "nh trong sheet_CTTT_BOM_Total", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o ho" & ChrW(224) & "n th" & ChrW(224) & "nh", 0 + 64
Call Focus(False)
End Sub

Sub locdl_ss_total()
Call Focus(True)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    duongdan = wbth.Path

    Set ws_Cttt = wbth.Sheets("CTTT")
    Set ws_Plm = wbth.Sheets("PLM")
    Set ws_R3 = wbth.Sheets("R3")
    Set ws_cttt_total = wbth.Sheets("CTTT_Total")
    Set ws_Tongket = wbth.Sheets("Tongket")
    On Error Resume Next
    ws_Cttt.ShowAllData
    ws_Cttt.Range("$A$2:$Y$5000").AutoFilter Field:=18, Criteria1:="=NG", _
        Operator:=xlOr, Criteria2:="=#N/A"
'        ws_Cttt.Range("$A$2:$Y$5000").AutoFilter Field:=17, Criteria1:="="
'        ws_Cttt.Range("$A$2:$Y$5000").AutoFilter Field:=19, Criteria1:="="
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
         On Error Resume Next
        ws_cttt_total.ShowAllData
        Columns("A:A").Select
        Selection.ColumnWidth = 24
        Columns("B:B").Select
        Selection.ColumnWidth = 24
    'ws_cttt_total.ShowAllData
    ws_cttt_total.Range("C1:H1").AutoFilter Field:=4, Criteria1:="=NG", _
    Operator:=xlOr, Criteria2:="=#N/A"
'    ws_cttt_total.Range("C1:H1").AutoFilter Field:=6, Criteria1:="="
'    wb_Bom.Close True
  On Error Resume Next
    ThisWorkbook.RefreshAll
    ws_cttt_total.Activate
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " l" & ChrW(7885) & "c d" & ChrW(7919) & " li" & ChrW(7879) & "u c" & ChrW(7847) & "n so s" & ChrW(225) & "nh trong sheet_CTTT_BOM_Total", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o ho" & ChrW(224) & "n th" & ChrW(224) & "nh", 0 + 64
Call Focus(False)
End Sub
Sub Focus(ByVal Flag As Boolean)
    With Application
        .EnableEvents = Not Flag
        .ScreenUpdating = Not Flag
        .Calculation = IIf(Flag, xlCalculationManual, xlCalculationAutomatic)
    End With
End Sub


