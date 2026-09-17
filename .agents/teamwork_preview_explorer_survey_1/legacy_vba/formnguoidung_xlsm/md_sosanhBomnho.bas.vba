' Stream: VBA/md_sosanhBomnho
' File: md_sosanhBomnho.bas

Attribute VB_Name = "md_sosanhBomnho"
Sub sosanhBompt()
Call Focus(True)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    thisfile = ActiveWorkbook.Name
    Set wb_file = Workbooks(thisfile)
    Set ws_Cttt = wb_file.Sheets("CTTT")
    Set ws_Plm_R3 = wb_file.Sheets("PLM_R3")
    duongdan = wb_file.Path
    duongdan = duongdan & "\"
    tenfilePLM = Dir(duongdan & "PLM*")
    tenfileR3 = Dir(duongdan & "R3*")
If tenfileR3 <> "" Then
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=duongdan & tenfileR3
    Application.DisplayAlerts = True
    Set wb_R3 = Workbooks(tenfileR3)
    lr_r3 = wb_R3.Sheets(1).Range("E" & Rows.Count).End(xlUp).Row + 1
    'xu ly cot du lieu trong R3 cua Virgo va phan con lai
    Dim vungdulieu As Range
    Set vungdulieu = Range("F1:F" & lr_r3).Find(what:="RevLev", MatchCase:=True, lookat:=xlWhole)
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
    wb_R3.Sheets(1).Range("E12:E" & lr_r3).Copy Destination:=ws_Plm_R3.Range("D2") ' ten linh kien
    wb_R3.Sheets(1).Range("J12:J" & lr_r3).Copy Destination:=ws_Plm_R3.Range("E2") ' so luong
    wb_R3.Sheets(1).Range("G12:G" & lr_r3).Copy Destination:=ws_Plm_R3.Range("F2") 'rev
    wb_R3.Close False
     '
    
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=duongdan & tenfilePLM 'mo file plm
    Application.DisplayAlerts = True
    
    Set wb_Plm = Workbooks(tenfilePLM)
    lr_plm = wb_Plm.Sheets(1).Range("D" & Rows.Count).End(xlUp).Row + 1
    wb_Plm.Sheets(1).Range("D3:D" & lr_plm).Copy Destination:=ws_Plm_R3.Range("A2")
    wb_Plm.Sheets(1).Range("F3:F" & lr_plm).Copy Destination:=ws_Plm_R3.Range("B2")
    wb_Plm.Sheets(1).Range("M3:M" & lr_plm).Copy Destination:=ws_Plm_R3.Range("C2")
    wb_Plm.Sheets(1).Range("K3:K" & lr_plm).Copy Destination:=ws_Plm_R3.Range("G2")
    wb_Plm.Close False
    '
    '
    ws_Plm_R3.Activate
    'chuyen du lieu cot A cua sheet PLM_R3 -> number
        ws_Plm_R3.Range("A2:A" & lr_plm).Select
        Selection.TextToColumns Destination:=Range("A2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    'chuyen du lieu cot B cua sheet PLM_R3->number
    ws_Plm_R3.Range("B2:B" & lr_plm).Select
        Selection.TextToColumns Destination:=Range("B2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
     'chuyen du lieu cot C cua sheet PLM_R3->number
    ws_Plm_R3.Range("C2:C" & lr_plm).Select
        Selection.TextToColumns Destination:=Range("C2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
             'chuyen du lieu cot D cua sheet PLM_R3->number
    ws_Plm_R3.Range("D2:D" & lr_plm).Select
        Selection.TextToColumns Destination:=Range("D2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
             'chuyen du lieu cot E cua sheet PLM_R3->number
    ws_Plm_R3.Range("E2:E" & lr_plm).Select
        Selection.TextToColumns Destination:=Range("E2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
             'chuyen du lieu cot F cua sheet PLM_R3->number
    ws_Plm_R3.Range("F2:F" & lr_plm).Select
        Selection.TextToColumns Destination:=Range("F2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    ThisWorkbook.RefreshAll
    ws_Cttt.Activate
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " t" & ChrW(7843) & "i PLM v" & ChrW(224) & " R3 v" & ChrW(224) & "o file c" & ChrW(7911) & "a b" & ChrW(7841) & "n.", , "Th" & ChrW(224) & "nh c" & ChrW(244) & "ng", 0 + 0
 End If
     If tenfileR3 = "" Or tenfilePLM = "" Then
        CreateObject("WScript.Shell").Popup "Ch" & ChrW(432) & "a c" & ChrW(243) & " file R3 ho" & ChrW(7863) & "c PLM trong th" & ChrW(432) & " m" & ChrW(7909) & "c so s" & ChrW(225) & "nh BOM", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
        End
     End If
Call Focus(False)
End Sub

Sub menuchinh()
uf_ssbnpt.Show
End Sub
