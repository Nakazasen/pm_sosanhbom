' Stream: VBA/md_dntd
' File: md_dntd.bas

Attribute VB_Name = "md_dntd"
Option Explicit

Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" ( _
ByVal hWnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Function OpenAnyFile(FileToOpen As String)

Call ShellExecute(0, "Open", FileToOpen & vbNullString, _
vbNullString, vbNullString, 1)

End Function

Sub OpenFile()
Call Focus(True)
Dim file1 As String
file1 = ThisWorkbook.Sheets("duongdan").Range("C1")
Call OpenAnyFile(file1)
Call Focus(False)
End Sub

