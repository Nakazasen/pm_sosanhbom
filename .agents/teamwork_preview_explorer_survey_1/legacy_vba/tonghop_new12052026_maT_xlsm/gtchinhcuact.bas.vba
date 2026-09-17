' Stream: VBA/gtchinhcuact
' File: gtchinhcuact.bas

Attribute VB_Name = "gtchinhcuact"
Sub ct_ssb()
Call Focus(True)
    Dim wbth As Workbook
    Dim ws_link As Worksheet ' sheet duongdan cua file tonghop
    Dim ws_thdl As Worksheet
    Dim tenwb As String 'ten workbook dang chay chuong trinh vba
    Dim tenwbth As String
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_thdl = wbth.Sheets("tonghopdl")
    Set ws_link = wbth.Sheets("duongdan")
If ws_link.Range("A6").Text = True Then
uf_mainmenu.Show
End If
If ws_link.Range("A7").Text = True Then
uf_mainmenu_new.Show
End If
If ws_link.Range("A6").Text = False Then
uf_sosanhbom.Show
End If
If ws_link.Range("A7").Text = False Then
uf_sosanhbom.Show
End If
If ws_link.Range("A6").Text = "" And ws_link.Range("A7").Text = "" Then
uf_sosanhbom.Show
End If
Call Focus(False)
End Sub

