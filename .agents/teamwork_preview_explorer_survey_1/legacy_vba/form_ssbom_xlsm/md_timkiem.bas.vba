' Stream: VBA/md_timkiem
' File: md_timkiem.bas

Attribute VB_Name = "md_timkiem"
Sub FillVLookup()
Call Focus(True)
    Dim wsPLM As Worksheet
    Dim wsCTTT As Worksheet
    Dim lastRowPLM As Long
    Dim lastRowCTTT As Long
    Dim i As Long
    Dim lookupValue As String
    Dim resultValue As Variant
    
    ' Xac ??nh cac sheet
    Set wsPLM = ThisWorkbook.Sheets("PLM")
    Set wsCTTT = ThisWorkbook.Sheets("CTTT")
    
    ' Xac ??nh dong cu?i cung c?a c? hai sheet
    lastRowPLM = wsPLM.Cells(wsPLM.Rows.Count, "C").End(xlUp).Row
    lastRowCTTT = wsCTTT.Cells(wsCTTT.Rows.Count, "C").End(xlUp).Row
    'wsPLM.Range("AJ1") = "T" & ChrW(234) & "n ng" & ChrW(432) & ChrW(7901) & "i ph" & ChrW(7909) & " tr" & ChrW(225) & "ch"
    ' Duy?t qua t?ng o trong c?t C c?a sheet PLM
    For i = 3 To lastRowPLM
        lookupValue = wsPLM.Cells(i, "C").Value
        ' S? d?ng ham VLOOKUP ?? tim gia tr?
        resultValue = Application.VLookup(lookupValue, wsCTTT.Range("C3:F" & lastRowCTTT), 4, False)
        
        ' Ki?m tra n?u tim th?y gia tr?
        If Not IsError(resultValue) Then
            wsPLM.Cells(i, "AH").Value = resultValue
        Else
            wsPLM.Cells(i, "AH").Value = ""
        End If
    Next i
    Call Focus(False)
End Sub

Sub CopyDataFromLinkedFile()
Call Focus(True)
    Dim wsPLM As Worksheet
    Dim wbSource As Workbook
    Dim sourceFile As String
    Dim lastRowSource As Long
    Dim lastRowDest As Long
    Dim rngSource As Range
    Dim rngDest As Range
    
    ' Xac ??nh sheet PLM
    Set wsPLM = ThisWorkbook.Sheets("PLM")
    
    ' L?y lien k?t file t? o W1000
    sourceFile = wsPLM.Range("W1000").Value
      'Delete cot X den AI
    On Error Resume Next
    wsPLM.ShowAllData
    wsPLM.Columns("X:AI").Delete
    
    ' M? file ngu?n
    On Error Resume Next
    Set wbSource = Workbooks.Open(sourceFile)
    On Error GoTo 0
    
    ' Ki?m tra n?u file ngu?n khong m? ???c
    If wbSource Is Nothing Then
        CreateObject("WScript.Shell").Popup "Kh" & ChrW(244) & "ng th" & ChrW(7875) & " m" & ChrW(7903) & " file", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
        Exit Sub
    End If
    
    ' Xac ??nh dong cu?i cung c?a d? li?u trong file ngu?n
    With wbSource.Sheets(1) ' Gi? s? d? li?u n?m tren sheet ??u tien
        lastRowSource = .Cells(.Rows.Count, "X").End(xlUp).Row
        Set rngSource = .Range("X1:AI" & lastRowSource) ' Gi? s? d? li?u b?t ??u t? dong 2
    End With
    
    ' Xac ??nh dong cu?i cung c?a d? li?u trong sheet PLM
    lastRowDest = wsPLM.Cells(wsPLM.Rows.Count, "X").End(xlUp).Row + 1
    Set rngDest = wsPLM.Range("X" & lastRowDest - 1 & ":AI" & (lastRowDest + rngSource.Rows.Count - 1))
    
    ' Sao chep d? li?u t? file ngu?n vao sheet PLM
    rngSource.Copy Destination:=rngDest
    wsPLM.Columns("AG").ColumnWidth = 20
    wsPLM.Columns("AH").ColumnWidth = 20
    wsPLM.Columns("AI").ColumnWidth = 20
    ' ?ong file ngu?n
    wbSource.Close SaveChanges:=False
    Call FillVLookup
    wsPLM.Columns("X:AI").Copy
    wsPLM.Columns("X:AI").PasteSpecial Paste:=xlPasteValues
    ' H?y ch?n vung ?a sao chep
      Application.CutCopyMode = False
   
   CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " t" & ChrW(236) & "m ki" & ChrW(7871) & "m Unit m" & ChrW(7865) & " th" & ChrW(224) & "nh c" & ChrW(244) & "ng, b" & ChrW(7841) & "n h" & ChrW(227) & "y x" & ChrW(225) & "c nh" & ChrW(7853) & "n d" & ChrW(7919) & " li" & ChrW(7879) & "u t" & ChrW(7915) & " c" & ChrW(7897) & "t X " & ChrW(273) & ChrW(7871) & "n c" & ChrW(7897) & "t AI nh" & ChrW(233), , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o th" & ChrW(224) & "nh c" & ChrW(244) & "ng", 0 + 64
    Call Focus(False)
End Sub






