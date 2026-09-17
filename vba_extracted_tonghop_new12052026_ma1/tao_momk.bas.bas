Attribute VB_Name = "tao_momk"
Sub taomk()
    Dim wbth As Workbook
    Dim ws_link As Worksheet ' sheet duongdan cua file tonghop
    Dim ws_thdl As Worksheet
    Dim tenwb As String 'ten workbook dang chay chuong trinh vba
    Dim tenwbth As String
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_thdl = wbth.Sheets("tonghopdl")
    Set ws_link = wbth.Sheets("duongdan")
    If uf_ps.tb_taops.Value = "" Then
        CreateObject("WScript.Shell").Popup "Kh" & ChrW(244) & "ng " & ChrW(273) & ChrW(432) & ChrW(7907) & "c " & ChrW(273) & ChrW(7863) & "t m" & ChrW(7853) & "t kh" & ChrW(7849) & "u tr" & ChrW(7889) & "ng", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o b" & ChrW(7843) & "o m" & ChrW(7853) & "t", 0 + 64
        Unload uf_ps
        uf_ps.Show
    End If
    If Len(uf_ps.tb_taops.Value) < 8 Then
    CreateObject("WScript.Shell").Popup "M" & ChrW(7853) & "t kh" & ChrW(7849) & "u kh" & ChrW(244) & "ng " & ChrW(273) & ChrW(432) & ChrW(7907) & "c " & ChrW(273) & ChrW(7863) & "t d" & ChrW(432) & ChrW(7899) & "i 8 k" & ChrW(253) & " t" & ChrW(7921), , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o b" & ChrW(7843) & "o m" & ChrW(7853) & "t", 0 + 64
    Unload uf_ps
    uf_ps.Show
    End If
    If Len(uf_ps.tb_taops.Value) > 8 Then
    ws_link.Range("A3").Value = uf_ps.tb_taops.Value
    CreateObject("WScript.Shell").Popup "B" & ChrW(7841) & "n " & ChrW(273) & ChrW(227) & " t" & ChrW(7841) & "o xong m" & ChrW(7853) & "t kh" & ChrW(7849) & "u, h" & ChrW(227) & "y nh" & ChrW(7899) & " m" & ChrW(7853) & "t kh" & ChrW(7849) & "u " & ChrW(273) & ChrW(7875) & " " & ChrW(273) & ChrW(259) & "ng nh" & ChrW(7853) & "p l" & ChrW(7847) & "n ti" & ChrW(7871) & "p theo", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o t" & ChrW(7841) & "o m" & ChrW(7853) & "t kh" & ChrW(7849) & "u th" & ChrW(224) & "nh c" & ChrW(244) & "ng", 0 + 64
    uf_ps.MultiPage1.Value = 2
    End If
End Sub
Sub doimk()
    Dim ws_link As Worksheet ' sheet duongdan cua file tonghop
    Dim ws_thdl As Worksheet
    Dim tenwb As String 'ten workbook dang chay chuong trinh vba
    Dim tenwbth As String
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_thdl = wbth.Sheets("tonghopdl")
    Set ws_link = wbth.Sheets("duongdan")
    If uf_ps.tb_mkc.Text <> ws_link.Range("A3").Text Then
    CreateObject("WScript.Shell").Popup "M" & ChrW(7853) & "t kh" & ChrW(7849) & "u c" & ChrW(361) & " sai", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o m" & ChrW(7853) & "t kh" & ChrW(7849) & "u c" & ChrW(361), 0 + 64
    End If
    If uf_ps.tb_mkc.Text = ws_link.Range("A3").Text And Len(uf_ps.tb_mkm.Text) < 4 Then
    CreateObject("WScript.Shell").Popup "M" & ChrW(7853) & "t kh" & ChrW(7849) & "u kh" & ChrW(244) & "ng " & ChrW(273) & ChrW(432) & ChrW(7907) & "c " & ChrW(273) & ChrW(7863) & "t d" & ChrW(432) & ChrW(7899) & "i 4 k" & ChrW(253) & " t" & ChrW(7921), , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o b" & ChrW(7843) & "o m" & ChrW(7853) & "t", 0 + 64
    Unload uf_ps
    uf_ps.Show
    End If
    If uf_ps.tb_mkc.Text = ws_link.Range("A3").Text And Len(uf_ps.tb_mkm.Text) >= 4 Then
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " thay " & ChrW(273) & ChrW(7893) & "i m" & ChrW(7853) & "t kh" & ChrW(7849) & "u", , "Thay " & ChrW(273) & ChrW(7893) & "i m" & ChrW(7853) & "t kh" & ChrW(7849) & "u th" & ChrW(224) & "nh c" & ChrW(244) & "ng", 0 + 64
    Call MoKhoa_TatCa_Sheet
    ws_link.Range("A3").Value = uf_ps.tb_mkm.Value
    uf_mainmenu.Show
    Unload uf_ps
    Unload uf_sosanhbom
    End If
End Sub
Sub momk()
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
    If uf_ps.tb_dangnhap.Text <> ws_link.Range("A3").Text Then
    CreateObject("WScript.Shell").Popup "B" & ChrW(7841) & "n " & ChrW(273) & ChrW(227) & " nh" & ChrW(7853) & "p sai m" & ChrW(7853) & "t kh" & ChrW(7849) & "u, vui l" & ChrW(242) & "ng nh" & ChrW(7853) & "p l" & ChrW(7841) & "i", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o " & ChrW(273) & ChrW(259) & "ng nh" & ChrW(7853) & "p th" & ChrW(7845) & "t b" & ChrW(7841) & "i", 0 + 64
    End If
    If uf_ps.tb_dangnhap.Text = ws_link.Range("A3").Text And uf_ps.cb_kieussbom.Value = "So s" & ChrW(225) & "nh BOM ki" & ChrW(7875) & "u c" & ChrW(361) Then
    Call MoKhoa_TatCa_Sheet
    ws_link.Range("A6") = uf_ps.ckb_duytri.Value
    uf_mainmenu.Show
    Unload uf_ps
    Unload uf_sosanhbom
    End If
    If uf_ps.tb_dangnhap.Text = ws_link.Range("A3").Text And uf_ps.cb_kieussbom.Value = "So s" & ChrW(225) & "nh BOM ki" & ChrW(7875) & "u m" & ChrW(7899) & "i" Then
    Call MoKhoa_TatCa_Sheet
    ws_link.Range("A7") = uf_ps.ckb_duytri.Value
    uf_mainmenu_new.Show
    Unload uf_ps
    Unload uf_sosanhbom
    End If
Call Focus(False)
End Sub

Sub MoKhoa_TatCa_Sheet()
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
For Each Ten_Sheet In wbth.Worksheets
Ten_Sheet.Unprotect Password = ws_link.Range("A3").Text
Next Ten_Sheet
End Sub
'Sub PasswordBreaker()
'If ActiveSheet.ProtectContents = False Then
'    MsgBox "Sheet '" & ActiveSheet.Name & "' is unprotected!", vbInformation
'Else
'    If MsgBox("Sheet '" & ActiveSheet.Name & "' is protected, do you want to unprotect it?", _
'    vbYesNo + vbQuestion, "Unprotect Active Sheet") = vbNo Then Exit Sub
'    Dim i As Integer, j As Integer, k As Integer
'    Dim l As Integer, m As Integer, n As Integer
'    Dim i1 As Integer, i2 As Integer, i3 As Integer
'    Dim i4 As Integer, i5 As Integer, i6 As Integer
'    On Error Resume Next
'    For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
'    For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
'    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
'    For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
'        ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
'        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
'        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
'    Next: Next: Next: Next: Next: Next
'    Next: Next: Next: Next: Next: Next
'    If ActiveSheet.ProtectContents = False Then MsgBox "Sheet '" & ActiveSheet.Name & "' is unprotected!", vbInformation
'End If
'End Sub
