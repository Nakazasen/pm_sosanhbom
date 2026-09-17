' Stream: VBA/tudong_folder_filengpt
' File: tudong_folder_filengpt.bas

Attribute VB_Name = "tudong_folder_filengpt"
Sub ditoisheetdk()
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_Lichsu = wbth.Worksheets("Lichsu")
ws_Lichsu.Activate
End Sub
Sub xoadulieucudk()
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_Lichsu = wbth.Worksheets("Lichsu")
    ws_Lichsu.Range("C2:C40").ClearContents
    ws_Lichsu.Range("E2:F40").ClearContents
End Sub
Sub taofolder_ngpt()
'tao folder huong xuat theo list dang ky
Call Focus(True)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_link = wbth.Sheets("duongdan")
    Set ws_Lichsu = wbth.Sheets("Lichsu")
    lr_lichsu_cotE = wbth.Sheets("Lichsu").Range("E" & Rows.Count).End(xlUp).Row
    lr_lichsu_cotC = wbth.Sheets("Lichsu").Range("C" & Rows.Count).End(xlUp).Row
    duongdan = wbth.Path
    duongdan = duongdan & "\"
    If lr_lichsu_cotE < 2 Then
    CreateObject("WScript.Shell").Popup "B" & ChrW(7841) & "n ch" & ChrW(432) & "a nh" & ChrW(7853) & "p m" & ChrW(227) & " m" & ChrW(225) & "y(h" & ChrW(432) & ChrW(7899) & "ng xu" & ChrW(7845) & "t) v" & ChrW(224) & "o list", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
    Exit Sub
    ElseIf lr_lichsu_cotC < 2 Then
    CreateObject("WScript.Shell").Popup "B" & ChrW(7841) & "n ch" & ChrW(432) & "a l" & ChrW(7921) & "a ch" & ChrW(7885) & "n ng" & ChrW(432) & ChrW(7901) & "i ph" & ChrW(7909) & " tr" & ChrW(225) & "ch " & ChrW(273) & ChrW(7875) & " t" & ChrW(7841) & "o file " & ChrW(273) & "i" & ChrW(7873) & "n linh ki" & ChrW(7879) & "n.", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
    Exit Sub
    End If
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For i = 2 To lr_lichsu_cotE
    folder_listhx = duongdan & Trim(wbth.Sheets("Lichsu").Range("E" & i))
    If (Not FSO.FolderExists(folder_listhx)) And wbth.Sheets("Lichsu").Range("F" & i) = "" Then
        FSO.CreateFolder (folder_listhx)
    End If
    Next
    Call taofile_nhapdslk_theolist
    Call taifiletonghopma1_theolist
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " t" & ChrW(7841) & "o xong danh s" & ChrW(225) & "ch BOM v" & ChrW(224) & " ng" & ChrW(432) & ChrW(7901) & "i ph" & ChrW(7909) & " tr" & ChrW(225) & "ch theo list c" & ChrW(224) & "i " & ChrW(273) & ChrW(7863) & "t.", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o th" & ChrW(224) & "nh c" & ChrW(244) & "ng", 0 + 64
  Call Focus(False)
End Sub
'tao file ngpt theo list dang ky
Sub taofile_nhapdslk_theolist()
Call Focus(True)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_link = wbth.Sheets("duongdan")
    Set ws_Lichsu = wbth.Sheets("Lichsu")
    duongdan = wbth.Path
    duongdan = duongdan & "\"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    lr_lichsu_cotC = wbth.Sheets("Lichsu").Range("C" & Rows.Count).End(xlUp).Row
    For i = 2 To lr_lichsu_cotC
    Call FSO.CopyFile(ws_link.Range("A2").Value & "formnguoidung.xlsm", duongdan)
    OldName = duongdan & "formnguoidung.xlsm"
    NewName = duongdan & wbth.Sheets("Lichsu").Range("D" & i) & ".xlsm"
    If (Not FSO.FileExists(NewName)) And wbth.Sheets("Lichsu").Range("C" & i) = "O" Then
    Name OldName As NewName
    Else
    Kill (OldName)
    End If
    Next
    Call Move_file_npt_theolist
    Exit Sub
  Call Focus(False)
End Sub

'chuyen file ngpt vao cac folder huong xuat

Sub Move_file_npt_theolist()
    'khai bao.
    Dim ObjFileSystem As Object
    Dim ObjFolder As Object
    Dim ObjTarget As Object
    Dim ObjSubFolder As Object
    Dim mamay As String
    Dim f, fc

    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_Lichsu = wbth.Sheets("Lichsu")
    lr_lichsu_cotE = wbth.Sheets("Lichsu").Range("E" & Rows.Count).End(xlUp).Row
    duongdan = wbth.Path & "\"
    'cai dat bien.
    Set ObjFileSystem = CreateObject("Scripting.FileSystemObject")
    Set ObjFolder = ObjFileSystem.GetFolder(duongdan)
    Set ObjSubFolder = ObjFolder.SubFolders
    '-- move file
    
    Set f = ObjFileSystem.GetFolder(duongdan)
    Set fc = f.files
    For Each f1 In fc
         If f1.Name = Dir(duongdan & "*mecha*") Then
            tennpt = ""
            tennpt = f1.Name
                For i = 2 To lr_lichsu_cotE
                If wbth.Sheets("Lichsu").Range("F" & i) = "" Then
                ObjFileSystem.CopyFile Source:=duongdan & tennpt, Destination:=duongdan & Trim(wbth.Sheets("Lichsu").Range("E" & i)) & "\"
                End If
                Next
         End If
         On Error Resume Next
         ObjFileSystem.DeleteFile duongdan & tennpt
     Next
End Sub

'copy file tong hop vao cac folder huong xuat
Sub taifiletonghopma1_theolist()
Call Focus(True)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_thdl = wbth.Sheets("tonghopdl")
    Set ws_link = wbth.Sheets("duongdan")
    duongdan = wbth.Path
    duongdan = duongdan & "\"
    tenform_thT = Dir(ws_link.Range("A2").Value & "*ma1*")
    Set FSO = CreateObject("Scripting.FileSystemObject")
    folder_taifile = duongdan & "taifileth"
    If (Not FSO.FolderExists(folder_taifile)) Then
        FSO.CreateFolder (folder_taifile)
    End If
    Call FSO.CopyFile(ws_link.Range("A2").Value & tenform_thT, folder_taifile & "\")
    ' doi ten file trong thu muc taifileth
    namerutgon = Left(tenform_thT, Len(tenform_thT) - 19)
    namerutgon = namerutgon & Format(Date, "dd.mm.yyyy") & "_ma1.xlsm"
    Name folder_taifile & "\" & tenform_thT As _
    folder_taifile & "\" & namerutgon
    Call Move_file_th_theolist
  Call Focus(False)
End Sub
Sub Move_file_th_theolist()
'khai bao.
    Dim ObjFileSystem As Object
    Dim ObjFolder As Object
    Dim ObjTarget As Object
    Dim ObjSubFolder As Object
    Dim mamay As String
    Dim f, fc

    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_Lichsu = wbth.Sheets("Lichsu")
    lr_lichsu_cotE = wbth.Sheets("Lichsu").Range("E" & Rows.Count).End(xlUp).Row
    duongdan = wbth.Path & "\"
    folder_taifile = duongdan & "taifileth\"
    'cai dat bien.
    Set ObjFileSystem = CreateObject("Scripting.FileSystemObject")
    Set ObjFolder = ObjFileSystem.GetFolder(duongdan)
    Set ObjSubFolder = ObjFolder.SubFolders
    '-- move file
    Set f = ObjFileSystem.GetFolder(folder_taifile)
    Set fc = f.files
    For Each f1 In fc
         If f1.Name = Dir(folder_taifile & "*ma1*") Then
            tenform_thT = f1.Name
                For i = 2 To lr_lichsu_cotE
                If wbth.Sheets("Lichsu").Range("F" & i) = "" Then
                ObjFileSystem.CopyFile Source:=folder_taifile & tenform_thT, Destination:=duongdan & Trim(wbth.Sheets("Lichsu").Range("E" & i)) & "\"
                End If
                Next
         End If
     Next
    On Error Resume Next
    'delete all files in folder
    Kill folder_taifile & "*.*"
    'delete empty folder
    If Dir(folder_taifile) = "" Then
    RmDir folder_taifile
    End If
End Sub



