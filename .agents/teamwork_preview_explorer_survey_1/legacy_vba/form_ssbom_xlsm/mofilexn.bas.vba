' Stream: VBA/mofilexn
' File: mofilexn.bas

Attribute VB_Name = "mofilexn"
Sub mofile_hangmuc1()
Dim file1 As String
file1 = ThisWorkbook.Sheets("Tongket").Range("Y33")
Workbooks.Open file1
End Sub

Sub mofile_hangmuc2()
Dim file2 As String
file2 = ThisWorkbook.Sheets("Tongket").Range("Y34")
Workbooks.Open file2
End Sub
Sub mofile_hangmuc3()
Dim file3 As String
file3 = ThisWorkbook.Sheets("Tongket").Range("Y35")
Workbooks.Open file3
End Sub

