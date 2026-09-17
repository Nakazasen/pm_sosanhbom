' Stream: VBA/ThisWorkbook
' File: ThisWorkbook.cls

Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
Dim wbth As Workbook
Dim ws_link As Worksheet ' sheet duongdan cua file tonghop
Dim ws_thdl As Worksheet
Dim tenwb As String 'ten workbook dang chay chuong trinh vba
Dim tenwbth As String
tenwbth = ActiveWorkbook.Name
Set wbth = Workbooks(tenwbth)
Set ws_thdl = wbth.Sheets("tonghopdl")
Set ws_link = wbth.Sheets("duongdan")
Dim Ten_Sheet As Worksheet
Dim MatKhauKhoa As String
If ws_link.Range("A6").Value <> "" Or ws_link.Range("A7").Value <> "" Then
ws_link.Range("A6").ClearContents
ws_link.Range("A7").ClearContents
End If
For Each Ten_Sheet In wbth.Worksheets
Ten_Sheet.Protect Password = ws_link.Range("A3").Text
Next Ten_Sheet
ActiveWorkbook.Save
End Sub

