Attribute VB_Name = "locbomfull_all"
Sub bom_tc14_full_all()
Call Focus(True)
Dim SearchRange As Range
Dim diachi As String
Dim namesub As Range
Dim diachidautien As String
Dim sttdong As Variant
Dim chuoihieuluc As Variant
Dim ngayhieuluc As Variant
Dim vitri As Variant
Dim chenhlechnam As Variant
Dim chenhlechthang As Variant
Dim ngayhieuluchomnay As String
Dim levelbom As Variant
Dim vitrilevelbom As Variant
Dim dc As Long
Dim hasChildren As Variant
Dim sttdongitemname As Variant
Dim vitriitemname As Variant
Dim vitritenmay As Variant
Dim sttdongboloc As Variant
Dim vitridencungcap As Variant
Dim vttrunggian As Variant
Dim lr_boloc As Long
Dim f, fc

Set FSO = CreateObject("Scripting.FileSystemObject")
Set wbth = Workbooks(ThisWorkbook.Name) 'wb tonghop
    duongdan = wbth.Path
    duongdan = duongdan & "\"
    'tao folder backupTC14full
    folder_tc14backup = duongdan & "backupTC14full"
    If (Not FSO.FolderExists(folder_tc14backup)) Then
        FSO.CreateFolder (folder_tc14backup)
    End If
    '
Set f = FSO.GetFolder(duongdan)
Set fc = f.files
For Each f1 In fc
 If f1.Name <> Dir(duongdan & "tonghop*") And f1.Name <> Dir(duongdan & "~$tonghop*", vbHidden) Then
        tenfilePLM = f1.Name
        ' doi ten de copy vao thu muc backupTC14 full
    TC14Full = Left(tenfilePLM, Len(tenfilePLM) - 5)
    TC14Full = TC14Full & "_" & Format(Date, "dd_mm_yyyy") & ".xlsm"
    OldName = tenfilePLM
    NewName = TC14Full
    FSO.CopyFile duongdan & tenfilePLM, archiveTo & duongdan & "backupTC14full" & "\", True
    ' doi ten file tc14full
    Name duongdan & "backupTC14full" & "\" & OldName As duongdan & "backupTC14full" & "\" & NewName
        '
        On Error GoTo loi
Application.DisplayAlerts = False
Workbooks.Open Filename:=duongdan & tenfilePLM
Application.DisplayAlerts = True
    Set wb_Plm = Workbooks(tenfilePLM)
    Set ws_Plm = wb_Plm.Worksheets(1)
    wb_Plm.Activate
dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
For i = 3 To dc ' vong for xu ly hieu luc cu
chuoihieuluc = ws_Plm.Range("I" & i).Value 'lay chuoi lien quan den ngay hieu luc
vitri = InStr(1, chuoihieuluc, "to", vbTextCompare) 'tim vi tri sau "to"
ngayhieuluc = Mid(chuoihieuluc, vitri + 3, 11)
ngayhieuluc = Format(ngayhieuluc, "dd/mm/yy")
ngayhieuluchomnay = Format(Date, "dd/mm/yy")
If ws_Plm.Range("I" & i).Find(what:="UP", MatchCase:=True, lookat:=xlPart) Is Nothing Then
On Error Resume Next ' gap nhung dong hieu luc trong thi tu dong bo qua
chenhlechnam = Right(ngayhieuluchomnay, 2) - Right(ngayhieuluc, 2)
chenhlechthang = Mid(ngayhieuluchomnay, 4, 2) - Mid(ngayhieuluc, 4, 2)
If chenhlechnam > 0 And chuoihieuluc <> "" Then
levelbom = ws_Plm.Range("B" & i).Value
If levelbom >= ws_Plm.Range("B" & i + 1) Then
ws_Plm.Rows(i).Select
Selection.Delete Shift:=xlUp
i = i - 1
dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
End If
If levelbom < ws_Plm.Range("B" & i + 1) Then
            '2024/05/21 reset vitrilevelbom
            vitrilevelbom = ""
            '
vitrilevelbom = ws_Plm.Range("B" & i + 1 & ":" & "B" & dc).Find(what:=levelbom, MatchCase:=True, lookat:=xlWhole).Address
'2024/05/21 xu ly khi khong the tim duoc vitrilevelbom_ bom iris C2M3Nl0
                If vitrilevelbom = "" Then
                    For hai = 1 To 4
                        levelbom = levelbom - 1
                            vitrilevelbom = ws_Plm.Range("B" & i + 1 & ":" & "B" & dc).Find(what:=levelbom, MatchCase:=True, lookat:=xlWhole).Address
                            If vitrilevelbom <> "" Then
                                Exit For
                            End If
                    Next hai
                End If
                'ket thuc
diachi = Right(vitrilevelbom, Len(vitrilevelbom) - 3) 'luu dia chi cua dong cuoi unit muon xoa
    For u = i + 1 To diachi
        If ws_Plm.Range("B" & u).Value <= ws_Plm.Range("B" & diachi).Value And levelbom < 6 Then
            diachi = u
            ws_Plm.Rows(i & ":" & diachi - 1).Select
            Selection.Delete Shift:=xlUp
            i = i - 1
            dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
            Exit For
        Else
        End If
    Next
Else
End If
Else
End If
If chenhlechnam = 0 And chenhlechthang > 1 And chuoihieuluc <> "" Then
levelbom = ws_Plm.Range("B" & i).Value
If levelbom >= ws_Plm.Range("B" & i + 1) Then
ws_Plm.Rows(i).Select
Selection.Delete Shift:=xlUp
i = i - 1
dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
End If
If levelbom < ws_Plm.Range("B" & i + 1) Then
            '2024/05/21 reset vitrilevelbom
            vitrilevelbom = ""
            '
vitrilevelbom = ws_Plm.Range("B" & i + 1 & ":" & "B" & dc).Find(what:=levelbom, MatchCase:=True, lookat:=xlWhole).Address
'2024/05/21 xu ly khi khong the tim duoc vitrilevelbom_ bom iris C2M3Nl0
                If vitrilevelbom = "" Then
                    For hai = 1 To 4
                        levelbom = levelbom - 1
                            vitrilevelbom = ws_Plm.Range("B" & i + 1 & ":" & "B" & dc).Find(what:=levelbom, MatchCase:=True, lookat:=xlWhole).Address
                            If vitrilevelbom <> "" Then
                                Exit For
                            End If
                    Next hai
                End If
                'ket thuc
diachi = Right(vitrilevelbom, Len(vitrilevelbom) - 3) 'luu dia chi cua dong cuoi unit muon xoa
    For u = i + 1 To diachi
        If ws_Plm.Range("B" & u).Value <= ws_Plm.Range("B" & diachi).Value And levelbom < 6 Then
            diachi = u
            ws_Plm.Rows(i & ":" & diachi - 1).Select
            Selection.Delete Shift:=xlUp
            i = i - 1
            dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
            Exit For
        Else
        End If
    Next
Else
End If
End If
Else
End If
dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
Next 'end vong for xu ly ngay hieu luc cu

'----xu ly ngay hieu luc trong co haschildren = true

dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
For i = 3 To dc ' vong for xu ly ngay hieu luc trong
chuoihieuluc = ws_Plm.Range("I" & i).Value
hasChildren = ws_Plm.Range("E" & i).Value
levelbom = ws_Plm.Range("B" & i).Value
If chuoihieuluc = "" And hasChildren = "True" And levelbom < ws_Plm.Range("B" & i + 1) Then
            '2024/05/21 reset vitrilevelbom
            vitrilevelbom = ""
            '
vitrilevelbom = ws_Plm.Range("B" & i + 1 & ":" & "B" & dc).Find(what:=levelbom, MatchCase:=True, lookat:=xlWhole).Address
'2024/05/21 xu ly khi khong the tim duoc vitrilevelbom_ bom iris C2M3Nl0
                If vitrilevelbom = "" Then
                    For hai = 1 To 4
                        levelbom = levelbom - 1
                            vitrilevelbom = ws_Plm.Range("B" & i + 1 & ":" & "B" & dc).Find(what:=levelbom, MatchCase:=True, lookat:=xlWhole).Address
                            If vitrilevelbom <> "" Then
                                Exit For
                            End If
                    Next hai
                End If
                'ket thuc
diachi = Right(vitrilevelbom, Len(vitrilevelbom) - 3) 'luu dia chi cua dong cuoi unit muon xoa
    For u = i + 1 To diachi
        If ws_Plm.Range("B" & u).Value <= ws_Plm.Range("B" & diachi).Value And levelbom < 6 Then
            diachi = u
            ws_Plm.Rows(i & ":" & diachi - 1).Select
            Selection.Delete Shift:=xlUp
            i = i - 1
            dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
            Exit For
        Else
        End If
    Next
End If
Next

'--xu ly ngay hieu luc trong co has children =false
dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
For i = 3 To dc ' vong for xu ly ngay hieu luc trong
chuoihieuluc = ws_Plm.Range("I" & i).Value
hasChildren = ws_Plm.Range("E" & i).Value
If chuoihieuluc = "" And hasChildren = "False" Then
    ws_Plm.Rows(i).Select
    Selection.Delete Shift:=xlUp
    i = i - 1
    dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
    Else
End If
Next

'---Loc theo bo loc dung rieng cho doi may

    'tenwbth = Dir(duongdan & "tonghop*")
    'Set wbth = Workbooks(tenwbth)
    Set ws_thdl = wbth.Worksheets("tonghopdl")
    Set ws_bolocbom = wbth.Worksheets("BolocBom")
    ws_bolocbom.Activate
    tenmay = uf_mainmenu.tb_tenloaimay1
    On Error GoTo loi1
    vitritenmay = ws_bolocbom.Range("A1:XFD1").Find(what:=tenmay, MatchCase:=True, lookat:=xlWhole).Address
    If vitritenmay = "" Then
loi1:
        CreateObject("WScript.Shell").Popup "Ch" & ChrW(432) & "a nh" & ChrW(7853) & "p t" & ChrW(234) & "n m" & ChrW(225) & "y ho" & ChrW(7863) & "c t" & ChrW(234) & "n m" & ChrW(225) & "y kh" & ChrW(244) & "ng t" & ChrW(7891) & "n t" & ChrW(7841) & "i trong c" & ChrW(417) & " s" & ChrW(7903) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u c" & ChrW(7911) & "a sheet BolocBom!!!", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
        wb_Plm.Close False
        Exit Sub
        Else
    End If
    sttdongboloc = Mid(vitritenmay, 2, 1) 'lay cot du lieu ten may tim dc
    lr_boloc = ws_bolocbom.Range(sttdongboloc & Rows.Count).End(xlUp).Row
    wb_Plm.Activate 'chuyen sang wb_plm
    For y = 2 To lr_boloc
    dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
        For i = 3 To dc
            On Error Resume Next
            'truong hop loc full name
            If ws_bolocbom.Range(sttdongboloc & y).Offset(0, 1) = "Full_name" And ws_bolocbom.Range(sttdongboloc & y).Offset(0, 2) = "" Then
                vitriitemname = ws_Plm.Range("K" & i - 1 & ":" & "K" & dc).Find(what:=ws_bolocbom.Range(sttdongboloc & y), MatchCase:=True, lookat:=xlWhole).Address
                sttdongitemname = Right(vitriitemname, Len(vitriitemname) - 3)
            End If
            'truong hop loc part name
            If ws_bolocbom.Range(sttdongboloc & y).Offset(0, 1) = "Part_name" And ws_bolocbom.Range(sttdongboloc & y).Offset(0, 2) = "" Then
                vitriitemname = ws_Plm.Range("K" & i - 1 & ":" & "K" & dc).Find(what:=ws_bolocbom.Range(sttdongboloc & y), MatchCase:=True, lookat:=xlPart).Address
                sttdongitemname = Right(vitriitemname, Len(vitriitemname) - 3)
            End If
            'truong hop loc bang malk
            If ws_bolocbom.Range(sttdongboloc & y).Offset(0, 2) <> "" Then
            vitriitemname = ws_Plm.Range("D" & i - 1 & ":" & "D" & dc).Find(what:=ws_bolocbom.Range(sttdongboloc & y).Offset(0, 2), MatchCase:=True, lookat:=xlWhole).Address
            sttdongitemname = Right(vitriitemname, Len(vitriitemname) - 3)
            End If
            If vitriitemname = "" Then
            Exit For
            ElseIf vitriitemname <> "" And ws_Plm.Range("E" & sttdongitemname) = "True" And ws_Plm.Range("B" & sttdongitemname) = 6 Then
            ws_Plm.Rows(sttdongitemname).Select
            Selection.Delete Shift:=xlUp
            dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
            i = sttdongitemname - 1
            ElseIf vitriitemname <> "" And ws_Plm.Range("E" & sttdongitemname) = "True" And ws_Plm.Range("B" & sttdongitemname) < ws_Plm.Range("B" & sttdongitemname + 1) Then
            levelbom = ws_Plm.Range("B" & sttdongitemname)
            '2024/05/21 reset vitrilevelbom
            vitrilevelbom = ""
            '
            vitrilevelbom = ws_Plm.Range("B" & sttdongitemname + 1 & ":" & "B" & dc).Find(what:=levelbom, MatchCase:=True, lookat:=xlWhole).Address
            '2024/05/21 xu ly khi khong the tim duoc vitrilevelbom_ bom iris C2M3Nl0
                If vitrilevelbom = "" Then
                    For hai = 1 To 4
                        levelbom = levelbom - 1
                            vitrilevelbom = ws_Plm.Range("B" & sttdongitemname + 1 & ":" & "B" & dc).Find(what:=levelbom, MatchCase:=True, lookat:=xlWhole).Address
                            If vitrilevelbom <> "" Then
                                Exit For
                            End If
                    Next hai
                End If
                'ket thuc
            vitridencungcap = Right(vitrilevelbom, Len(vitrilevelbom) - 3)
                For vttrunggian = sttdongitemname + 1 To vitridencungcap
                    If ws_Plm.Range("B" & vttrunggian).Value <= ws_Plm.Range("B" & vitridencungcap).Value And levelbom < 6 Then
                        vitridencungcap = vttrunggian
                        ws_Plm.Rows(sttdongitemname + 1 & ":" & vitridencungcap - 1).Select
                        Selection.Delete Shift:=xlUp
                        dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
                        i = sttdongitemname
                        Exit For
                    Else
                    End If
                Next
            ElseIf vitriitemname <> "" And ws_Plm.Range("E" & sttdongitemname) = "True" And ws_Plm.Range("B" & sttdongitemname) >= ws_Plm.Range("B" & sttdongitemname + 1) And ws_Plm.Range("B" & sttdongitemname) < 6 Then
            ElseIf vitriitemname <> "" And ws_Plm.Range("E" & sttdongitemname) = "False" Then
            ws_Plm.Rows(sttdongitemname).Select
            Selection.Delete Shift:=xlUp
            dc = ws_Plm.Range("A" & Rows.Count).End(xlUp).Row
            i = sttdongitemname - 1
            Else
            End If
            vitriitemname = ""
        Next
   Next
  wb_Plm.Close True
    Else
End If
Next
  Unload uf_mainmenu
  ws_thdl.Activate
  Call Nhapduongdan_xulyPLM
  CreateObject("WScript.Shell").Popup "Bom " & ChrW(273) & ChrW(227) & " " & ChrW(273) & ChrW(432) & ChrW(7907) & "c x" & ChrW(7917) & " l" & ChrW(253) & " xong.", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o th" & ChrW(224) & "nh c" & ChrW(244) & "ng", 0 + 64
Call Focus(False)
Exit Sub
loi:
CreateObject("WScript.Shell").Popup "Kh" & ChrW(244) & "ng t" & ChrW(7891) & "n t" & ChrW(7841) & "i file PLM c" & ChrW(249) & "ng th" & ChrW(432) & " m" & ChrW(7909) & "c v" & ChrW(7899) & "i file t" & ChrW(7893) & "ng h" & ChrW(7907) & "p!!!", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
End Sub

'move file PLM sau khi xu ly ve thu muc chi dinh

Sub Nhapduongdan_xulyPLM()
    'khai bao.
    Dim ObjFileSystem As Object
    Dim ObjFolder As Object
    Dim ObjTarget As Object
    Dim ObjSubFolder As Object
    Dim mamay As String
    Dim f, fc

    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_downloadR3 = wbth.Worksheets("Download_R3_tudong")
    duongdan = wbth.Path & "\"
    lr_file = ws_downloadR3.Range("A" & Rows.Count).End(xlUp).Row
    'cai dat bien.
    Set ObjFileSystem = CreateObject("Scripting.FileSystemObject")
    Set ObjFolder = ObjFileSystem.GetFolder(duongdan)
    Set ObjSubFolder = ObjFolder.SubFolders
    For Each ObjTarget In ObjSubFolder
            mamay = InStr(Right(ObjTarget, 10), "110")
            If mamay = 1 Then
                ws_downloadR3.Range("A" & lr_file + 1) = Right(ObjTarget, 10)
                lr_file = ws_downloadR3.Range("A" & Rows.Count).End(xlUp).Row
            Else
            End If
    Next
    lr_file = ws_downloadR3.Range("A" & Rows.Count).End(xlUp).Row
    Set f = ObjFileSystem.GetFolder(duongdan)
    Set fc = f.files
    For Each f1 In fc
         If f1.Name <> Dir(duongdan & "tonghop*") And f1.Name <> Dir(duongdan & "~$tonghop*", vbHidden) Then
            tenfilePLM = f1.Name
                For i = 2 To lr_file
                If Mid(tenfilePLM, 5, 10) = ws_downloadR3.Range("A" & i) Then
                On Error Resume Next
                ObjFileSystem.movefile Source:=duongdan & tenfilePLM, Destination:=duongdan & ws_downloadR3.Range("A" & i) & "\"
                Else
                End If
                Next
         End If
     Next
     ws_downloadR3.Range("A2:B" & lr_file).ClearContents
End Sub


