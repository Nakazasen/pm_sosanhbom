' Stream: VBA/taofiledslk
' File: taofiledslk.bas

Attribute VB_Name = "taofiledslk"
Sub taofile_nhapdslk1()
Call Focus(True)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_thdl = wbth.Sheets("tonghopdl")
    Set ws_link = wbth.Sheets("duongdan")
    duongdan = wbth.Path
    duongdan = duongdan & "\"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Call FSO.CopyFile(ws_link.Range("A2").Value & "formnguoidung.xlsm", duongdan)
    OldName = duongdan & "formnguoidung.xlsm"
    NewName = duongdan & uf_taofilenpt.cb_tenpt.Value & ".xlsm"
    On Error GoTo loi
    If FSO.FileExists(OldName) Then FSO.movefile OldName, NewName
    Set FSO = Nothing
    Call Move_file_npt
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " t" & ChrW(7841) & "o file th" & ChrW(224) & "nh c" & ChrW(244) & "ng", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o th" & ChrW(224) & "nh c" & ChrW(244) & "ng", 0 + 64
    Exit Sub
loi:
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " t" & ChrW(7891) & "n t" & ChrW(7841) & "i file ng" & ChrW(432) & ChrW(7901) & "i ph" & ChrW(7909) & " tr" & ChrW(225) & "ch, h" & ChrW(227) & "y ki" & ChrW(7875) & "m tra l" & ChrW(7841) & "i!", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
    Kill (duongdan & "formnguoidung.xlsm")
  Call Focus(False)
End Sub

'move file ngpt sau khi tao ve thu muc chi dinh

Sub Move_file_npt()
    'khai bao.
    Dim ObjFileSystem As Object
    Dim ObjFolder As Object
    Dim ObjTarget As Object
    Dim ObjSubFolder As Object
    Dim mamay As String
    Dim f, fc

    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    
    duongdan = wbth.Path & "\"
    'cai dat bien.
    Set ObjFileSystem = CreateObject("Scripting.FileSystemObject")
    Set ObjFolder = ObjFileSystem.GetFolder(duongdan)
    Set ObjSubFolder = ObjFolder.SubFolders
    '
    Dim coll As Object
    Set coll = CreateObject("System.collections.ArrayList")
    '
    For Each ObjTarget In ObjSubFolder
            mamay = InStr(Right(ObjTarget, 10), "T10") ' ket qua la 1 boi vi chuoi tim duoc o vi tri thu nhat
            If mamay = 1 Then
                coll.Add Right(ObjTarget, 10)
            Else
            End If
    Next
    '-- move file
    
    Set f = ObjFileSystem.GetFolder(duongdan)
    Set fc = f.files
    For Each f1 In fc
         If f1.Name = Dir(duongdan & "*mecha*") Then
            tennpt = f1.Name
                For i = 0 To coll.Count - 1
                On Error Resume Next
                ObjFileSystem.CopyFile Source:=duongdan & tennpt, Destination:=duongdan & coll(i) & "\"
                Next
         End If
     Next
     ObjFileSystem.DeleteFile duongdan & tennpt
End Sub




