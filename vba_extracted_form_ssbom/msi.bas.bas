Attribute VB_Name = "msi"
Sub sosanhmsi()
Call Focus(True)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_Cttt = wbth.Sheets("CTTT")
    Set ws_Plm = wbth.Sheets("PLM")
    Set ws_R3 = wbth.Sheets("R3")
    Set ws_msi = wbth.Sheets("MSI_7980_7990")
    Set ws_Tongket = wbth.Sheets("Tongket")
    ws_msi.Activate
    Range("M2:N36").ClearContents
    Range("K2:K36").ClearContents
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        Range("B2:B36").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
            Range("D2:D36").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    For i = 2 To 36
            ws_Plm.Activate
            lr_kq = Range("A" & Rows.Count).End(xlUp).Row
            Range("A1:E1").Select
            Selection.AutoFilter
            'tim ma unit
            Range("$A$1:$E$" & lr_kq).AutoFilter Field:=3, Criteria1:="*" & ws_msi.Range("B" & i).Value & "*"
            ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 3).Select
            Selection.Copy
            ws_msi.Activate
            Range("M" & i).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

            'tim 3 ky tu msi unit
            Workbooks.Open Filename:=ws_Tongket.Range("Y31").Value
            duongdan_fixserial = ws_Tongket.Range("Y32").Value
            tenfile_fixserial = Dir(duongdan_fixserial & "FIX_SERIAL_DLTOOL_VER010.xls")
            Set wb_FixSerial = Workbooks(tenfile_fixserial)
            Set ws_msiunit = wb_FixSerial.Sheets("UNIT")
            Set ws_msimay = wb_FixSerial.Sheets("MACHINE")
                If i < 36 And ws_msi.Range("B" & i).Value <> "" Then
                ws_msiunit.Activate
                    lr_fix = ws_msiunit.Range("A" & Rows.Count).End(xlUp).Row + 1
                    Range("A1:F1").Select
                    Selection.AutoFilter
                    Range("$A$1:$F$" & lr_fix).AutoFilter Field:=1, Criteria1:="*" & Left(ws_msi.Range("B" & i).Value, 9) & "*"
                    ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 6).Select
                    Selection.Copy
                    ws_msi.Activate
                    Range("N" & i).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
                   ws_msi.Range("N" & i) = CStr(ws_msi.Range("N" & i))
                End If
                         'tim LABEL_COMMENT(SERVICE)
                If i < 36 And ws_msi.Range("B" & i).Value <> "" Then
                ws_msiunit.Activate
                    lr_fix = ws_msiunit.Range("A" & Rows.Count).End(xlUp).Row + 1
                    Range("A1:H1").Select
                    Selection.AutoFilter
                    Range("$A$1:$H$" & lr_fix).AutoFilter Field:=1, Criteria1:="*" & Left(ws_msi.Range("B" & i).Value, 9) & "*"
                    ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 8).Select
                    Selection.Copy
                    ws_msi.Activate
                    Range("O" & i).Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
                   ws_msi.Range("O" & i) = CStr(ws_msi.Range("O" & i))
                End If
            'tim 3 ky tu msi hontai
            If i = 36 And ws_msi.Range("B" & i).Value <> "" Then
                ws_msimay.Activate
                lr_fix = ws_msimay.Range("A" & Rows.Count).End(xlUp).Row + 1
                Range("A1:F1").Select
                Selection.AutoFilter
                Range("$A$1:$F$" & lr_fix).AutoFilter Field:=1, Criteria1:="*" & Left(ws_msi.Range("B" & i).Value, 10) & "*"
                ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 6).Select
                Selection.Copy
                ws_msi.Activate
                Range("N" & i).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            End If
            '---- Truong hop dong 36 ma may bi bo trong se copy du lieu trong PLM de paste vao
            If i = 36 And ws_msi.Range("B" & i).Value = "" Then
                ws_msimay.Activate
                ws_msi.Range("B" & i).Value = ws_Plm.Range("C2").Value
                lr_fix = ws_msimay.Range("A" & Rows.Count).End(xlUp).Row + 1
                Range("A1:F1").Select
                Selection.AutoFilter
                Range("$A$1:$F$" & lr_fix).AutoFilter Field:=1, Criteria1:="*" & Left(ws_msi.Range("B" & i).Value, 10) & "*"
                ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 6).Select
                Selection.Copy
                ws_msi.Activate
                Range("N" & i).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            End If
            'truong hop du lieu cot M trong
            If Range("M" & i).Value = "" Then
            'boi mau do cot B
                Range("B" & i).Select
                    With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                    End With
               'boi mau do cot K va dien NG
                Range("K" & i).Select
                   With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                    Range("K" & i) = "NG"
            End If
            
           'truong hop du lieu cot E trong thi dien gia tri cua cot E = -
            If Range("E" & i).Value = "" Then
            'boi mau do cot B
            Range("E" & i).Value = "-"
            End If

            'M trong va D <> N
                        If Range("M" & i).Value = "" And CVar(Range("N" & i).Value) <> CVar(Range("D" & i).Value) Then
                    'Boi do cot B
                    Range("B" & i).Select
                    With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                    End With
                   'Boi do cot D
                 Range("D" & i).Select
                 With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                'Boi do va ghi NG vao cot K
                 Range("K" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                    Range("K" & i) = "NG"
            End If
            'truong hop du lieu cot m <> trong, dl cot n = trong-> chua cap nhat fix serial tool
            If Range("M" & i).Value <> "" And Range("N" & i).Value = "" Then
            'boi do cot B
                    Range("B" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
             'boi do va ghi NG cot K
                     Range("K" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                    Range("K" & i) = "NG"
            End If
            'M khac trong,  va cot D = cot N, cot E = cot O -> dien mau xanh va in ket qua ra cot K
            If Range("M" & i).Value <> "" And CVar(Range("N" & i).Value) = CVar(Range("D" & i).Value) And CVar(Range("E" & i).Value) = CVar(Range("O" & i).Value) Then
                    'boi mau xanh cot B
                    Range("B" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 6750054
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                'boi mau xanh cot D
                    Range("D" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 6750054
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                'boi mau xanh cot E
                    Range("E" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 6750054
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                 'boi mau xanh cot K va dien OK.
                    Range("K" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 6750054
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                Range("K" & i) = "OK"
            End If
                'M khac trong,  va cot D = cot N, cot E ="-" va cot O = trong -> dien mau xanh va in ket qua ra cot K
            If Range("M" & i).Value <> "" And CVar(Range("N" & i).Value) = CVar(Range("D" & i).Value) And CVar(Range("E" & i).Value) = "-" And CVar(Range("O" & i).Value) = "" Then
                    'boi mau xanh cot B
                    Range("B" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 6750054
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                'boi mau xanh cot D
                    Range("D" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 6750054
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                'boi mau xanh cot E
                    Range("E" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 6750054
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                 'boi mau xanh cot K va dien OK.
                    Range("K" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 6750054
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                Range("K" & i) = "OK"
            End If
            'M khac trong,  va cot D = cot N, cot E ="-" va cot O <> trong -> dien mau xanh vao cot B,D va in ket qua OK ra cot K, boi do cot E
            If Range("M" & i).Value <> "" And CVar(Range("N" & i).Value) = CVar(Range("D" & i).Value) And CVar(Range("E" & i).Value) = "-" And CVar(Range("O" & i).Value) <> "" Then
                    'boi mau xanh cot B
                    Range("B" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 6750054
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                'boi mau xanh cot D
                    Range("D" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 6750054
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                'boi mau do cot E
                  Range("E" & i).Select
                 With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                'boi mau xanh cot K va dien OK.
                    Range("K" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 6750054
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                Range("K" & i) = "OK"
            End If
            'cot M khac trong va cot D va cot N khac nhau->boi mau do va dien NG vao cot K
            If Range("M" & i).Value <> "" And CVar(Range("N" & i).Value) <> CVar(Range("D" & i).Value) Then
                    If CVar(Range("E" & i).Value) = "-" And CVar(Range("O" & i).Value) = "" Then
                       'boi mau xanh cot E
                      Range("E" & i).Select
                         With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 6750054
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    End If
                    Range("B" & i).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 6750054
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                    Range("D" & i).Select
                 With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                    Range("K" & i).Select
                 With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                Range("K" & i) = "NG"
            End If
            'cot M khac trong va cot E va cot O khac nhau, cot E khac gia tri "-" ->boi mau do va dien NG vao cot K
                        If Range("M" & i).Value <> "" And CVar(Range("E" & i).Value) <> CVar(Range("O" & i).Value) And CVar(Range("E" & i).Value) <> "-" Then
                    Range("E" & i).Select
                 With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                    Range("K" & i).Select
                 With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                Range("K" & i) = "NG"
            End If
    Next i
'    Dim FillColor As Integer
'    MsgBox Range("C2").Interior.ColorIndex
    wb_FixSerial.Close False
Call Focus(False)
End Sub






