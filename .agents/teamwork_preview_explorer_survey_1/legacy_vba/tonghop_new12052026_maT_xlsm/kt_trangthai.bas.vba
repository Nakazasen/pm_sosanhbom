' Stream: VBA/kt_trangthai
' File: kt_trangthai.bas

Attribute VB_Name = "kt_trangthai"
Sub kiemtra_trangthainhapdl()
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
CreateObject("WScript.Shell").Popup "To" & ChrW(224) & "n b" & ChrW(7897) & " ng" & ChrW(432) & ChrW(7901) & "i ph" & ChrW(7909) & " tr" & ChrW(225) & "ch " & ChrW(273) & ChrW(227) & " copy danh s" & ChrW(225) & "ch linh ki" & ChrW(7879) & "n", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o ki" & ChrW(7875) & "m tra ho" & ChrW(224) & "n th" & ChrW(224) & "nh", 0 + 64
End If
Call Focus(False)
End Sub


