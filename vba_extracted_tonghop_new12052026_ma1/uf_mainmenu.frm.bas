Attribute VB_Name = "uf_mainmenu"
Attribute VB_Base = "0{EC8D0125-51F8-49C9-AE80-9426669976DA}{2817B975-A34A-4DD1-9231-873E5EF6374C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub Bo_loc_Click()
lb1.Visible = True
tb_tenloaimay.Visible = True
cb_1.Visible = True
lb2.Visible = True
tb_tenloaimay1.Visible = True
cb_2.Visible = True
End Sub

Private Sub Bosung_CTTT_Click()
Call capnhatDSLK_toBom
End Sub

Private Sub cb_1_Click()
Call bom_tc14_full
End Sub

Private Sub cb_2_Click()
Call bom_tc14_full_all
End Sub

Private Sub cb_dntd_Click()
Call OpenFile
End Sub

Private Sub cb14_Click()
uf_mainmenu.cb13.Visible = False
uf_mainmenu.cb14.Visible = False
uf_mainmenu.lb_tm.Visible = True
uf_mainmenu.tb_ngaysxdd.Visible = True
uf_mainmenu.cb15.Visible = True
uf_mainmenu.CommandButton12.Visible = False
End Sub
Private Sub cb13_Click()
Unload uf_mainmenu
Call Nhapduongdan
End Sub

Private Sub cb15_Click()
Unload uf_mainmenu
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_downloadR3 = wbth.Worksheets("Download_R3_tudong")
    If tb_ngaysxdd.Value = "" Then
    CreateObject("WScript.Shell").Popup "Ch" & ChrW(432) & "a nh" & ChrW(7853) & "p ng" & ChrW(224) & "y s" & ChrW(7843) & "n xu" & ChrW(7845) & "t", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
    Exit Sub
    Else
    ws_downloadR3.Range("B2") = tb_ngaysxdd.Value
    End If
Call Nhapduongdan_auto
End Sub

Private Sub cmd_guimailnhapBom_Click()
Unload uf_mainmenu
Call SendEmails
End Sub

Private Sub cmd_guimailql_Click()
Unload uf_mainmenu
Call ForwardEmails
End Sub

Private Sub CommandButton1_Click()
Unload uf_mainmenu
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_Lichsu = wbth.Worksheets("Lichsu")
    ws_Lichsu.Activate
    ws_Lichsu.Range("C1").Select
End Sub

Private Sub CommandButton10_Click()
Call locdl_ss
End Sub


Private Sub CommandButton12_Click()
uf_mainmenu.cb13.Visible = True
uf_mainmenu.cb14.Visible = True
End Sub

Private Sub CommandButton13_Click()
Call taifiletonghopma1
Unload uf_mainmenu
End Sub

Private Sub CommandButton2_Click()
Call kiemtra_trangthainhapdl
End Sub

Private Sub CommandButton3_Click()
Call kiemtra_trangthainhapdlok_ng
End Sub

Private Sub CommandButton4_Click()
Call TaoFileSSB
End Sub

Private Sub CommandButton7_Click()
Call xoa_dlold
End Sub

Private Sub CommandButton8_Click()
Call capnhat_PLM
End Sub

Private Sub CommandButton9_Click()
Call capnhat_R3
End Sub

Private Sub UserForm_Activate()
Dim dc As Long
dc = ThisWorkbook.Sheets("Lichsu").Range("A" & Rows.Count).End(xlUp).Row
If dc >= 2 Then
lb_yccn.Visible = True
End If
End Sub

