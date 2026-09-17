' Stream: VBA/uf_sosanhbom
' File: uf_sosanhbom.frm

Attribute VB_Name = "uf_sosanhbom"
Attribute VB_Base = "0{87627E65-5D42-459B-B3B8-88D7DB2D7CC1}{70C6A459-1317-4460-B91C-988B302C9E45}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Public Sub CommandButton6_Click()
    Dim wbth As Workbook
    Dim ws_link As Worksheet ' sheet duongdan cua file tonghop
    Dim ws_thdl As Worksheet
    Dim tenwb As String 'ten workbook dang chay chuong trinh vba
    Dim tenwbth As String
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_thdl = wbth.Sheets("tonghopdl")
    Set ws_link = wbth.Sheets("duongdan")
    If ws_link.Range("A3") <> "" And ws_link.Range("A6").Value = "" And ws_link.Range("A7").Value = "" Then
    uf_ps.MultiPage1.Value = 1
    uf_ps.Show
    End If
    If ws_link.Range("A3") <> "" And ws_link.Range("A6").Value = False And ws_link.Range("A7").Value = False Then
    uf_ps.MultiPage1.Value = 1
    uf_ps.Show
    End If
    If ws_link.Range("A6").Value = True Then
    uf_mainmenu.Show
    Unload uf_sosanhbom
    End If
    If ws_link.Range("A7").Value = True Then
    uf_mainmenu_new.Show
    Unload uf_sosanhbom
    End If
End Sub
