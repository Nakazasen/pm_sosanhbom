Attribute VB_Name = "md_tinhnangkhac"
Sub taifiletonghopma1()
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
    Call Move_file
  Call Focus(False)
End Sub
Sub Move_file()
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
    folder_taifile = duongdan & "taifileth\"
    'cai dat bien.
    Set ObjFileSystem = CreateObject("Scripting.FileSystemObject")
    Set ObjFolder = ObjFileSystem.GetFolder(duongdan)
    Set ObjSubFolder = ObjFolder.SubFolders
    '
    Dim coll As Object
    Set coll = CreateObject("System.collections.ArrayList")
    '
    For Each ObjTarget In ObjSubFolder
            mamay = InStr(Right(ObjTarget, 10), "110") ' ket qua la 1 boi vi chuoi tim duoc o vi tri thu nhat
            If mamay = 1 Then
                On Error Resume Next
               Kill (ObjTarget & "\" & Dir(duongdan & "*ma1*"))
                coll.Add Right(ObjTarget, 10)
            Else
            End If
    Next
    '-- move file
    
    Set f = ObjFileSystem.GetFolder(folder_taifile)
    Set fc = f.files
    For Each f1 In fc
         If f1.Name = Dir(folder_taifile & "*ma1*") Then
            tenform_thT = f1.Name
                For i = 0 To coll.Count - 1
                On Error Resume Next
                ObjFileSystem.CopyFile Source:=folder_taifile & tenform_thT, Destination:=duongdan & coll(i) & "\"
                Next
         End If
     Next
    FSO.movefile Source:=folder_taifile & tenform_thT, Destination:=duongdan
    Application.DisplayAlerts = False
    Workbooks.Open Filename:=duongdan & tenform_thT
    Application.DisplayAlerts = True
    Set wb_thdl_new = Workbooks(tenform_thT)
    On Error Resume Next
    'delete all files in folder
    Kill folder_taifile & "\*.*"
    'delete empty folder
    RmDir folder_taifile & "\"
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " t" & ChrW(7843) & "i file t" & ChrW(7893) & "ng h" & ChrW(7907) & "p m" & ChrW(7899) & "i v" & ChrW(7873) & " Project, h" & ChrW(227) & "y x" & ChrW(243) & "a 1 file c" & ChrW(361) & " b" & ChrW(234) & "n ngo" & ChrW(224) & "i folder h" & ChrW(432) & ChrW(7899) & "ng xu" & ChrW(7845) & "t nh" & ChrW(233) & ".", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o th" & ChrW(224) & "nh c" & ChrW(244) & "ng", 0 + 64
    wbth.Close False
End Sub

