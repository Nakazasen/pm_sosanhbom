Attribute VB_Name = "md_guimail"
'Option Explicit
Sub SendEmails()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim cell As Range
    Dim EmailList As Range
    Dim Subject As String
    Dim HTMLBody As String
    Dim AttachmentPath As String
    Dim CCAddress As String
    
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_Mail = wbth.Worksheets("Mail")
    duongdan = wbth.Path
    duongdan = duongdan & "\"
    

    ' L?y tieu ?? t? sheet Excel
    Subject = ws_Mail.Range("A2").Value ' Gi? s? tieu ?? ? o A2
    
    ' L?y n?i dung HTML t? sheet Excel
    HTMLBody = ws_Mail.Range("A3").Value ' Gi? s? n?i dung HTML ? o A3
    
    ' L?y d?a ch? CC t? � D3
    CCAddress = ws_Mail.Range("B3").Value
    
    ' Ki?m tra n?u ng??i dung ?a nh?p d? li?u
    If Subject = "" Or HTMLBody = "" Then
        CreateObject("WScript.Shell").Popup "Vui l" & ChrW(242) & "ng nh" & ChrW(7853) & "p ti" & ChrW(234) & "u " & ChrW(273) & ChrW(7873) & " " & ChrW(7903) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u A2 v" & ChrW(224) & " n" & ChrW(7897) & "i dung th" & ChrW(432) & " " & ChrW(7903) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u A3 trong sheet Mail", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", 0 + 64
        Exit Sub
    End If
    
    ' Thi?t l?p danh sach email
    Set EmailList = ws_Mail.Range("B2") ' Thay ??i ph?m vi theo nhu c?u
    
    ' T?o ??i t??ng Outlook
    Set OutApp = CreateObject("Outlook.Application")
    
    ' L?y ???ng d?n ??n th? m?c ch?a workbook
    AttachmentPath = ThisWorkbook.Path
    
    ' Thay th? ???ng d?n trong HTML n?u c?n
    HTMLBody = Replace(HTMLBody, "{AttachmentPath}", AttachmentPath)
    
    ' G?i email cho t?ng ??a ch? trong danh sach
    For Each cell In EmailList
        If cell.Value Like "*@*.?*" Then ' Ki?m tra ??a ch? email h?p l?
            Set OutMail = OutApp.CreateItem(0)
            With OutMail
                .To = cell.Value
                .CC = CCAddress
                .Subject = Subject
                .HTMLBody = HTMLBody
                .Send
            End With
            Set OutMail = Nothing
        End If
    Next cell
    
    ' ?ong Outlook
    'OutApp.Quit
    'Set OutApp = Nothing
    
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " g" & ChrW(7917) & "i mail th" & ChrW(224) & "nh c" & ChrW(244) & "ng!", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", 0 + 64
End Sub



Sub ForwardEmails()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim cell As Range
    Dim EmailList As Range
    Dim Subject As String
    Dim HTMLBody As String
    Dim AttachmentPath As String
    Dim CCAddress As String
    
    Dim tenwbth As String
    Dim wbth As Workbook
    Dim ws_Mail As Worksheet
    Dim duongdan As String
    
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_Mail = wbth.Worksheets("Mail")
    duongdan = wbth.Path & "\"
    
    ' Lay tieu de thu
    Subject = ws_Mail.Range("C2").Value ' Gi? s? ti黏 d? ? � A2
    
    ' L?y n?i dung HTML t? sheet Excel
    HTMLBody = ws_Mail.Range("C3").Value ' Gi? s? n?i dung HTML ? � C3
    
    ' L?y d?a ch? CC t? � D3
    CCAddress = ws_Mail.Range("D3").Value
    
    ' Ki?m tra n?u ngu?i dg d� nh?p d? li?u
    If Subject = "" Or HTMLBody = "" Then
        CreateObject("WScript.Shell").Popup "Vui l" & ChrW(242) & "ng nh" & ChrW(7853) & "p ti" & ChrW(234) & "u " & ChrW(273) & ChrW(7873) & " " & ChrW(7903) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u C2 v" & ChrW(224) & " n" & ChrW(7897) & "i dung th" & ChrW(432) & " " & ChrW(7903) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u C3 trong sheet Mail", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", 0 + 64
        Exit Sub
    End If
   
    ' Thi?t l?p danh s當h email
    Set EmailList = ws_Mail.Range("D2") ' Thay d?i ph?m vi theo nhu c?u
    
    ' T?o d?i tu?ng Outlook
    Set OutApp = CreateObject("Outlook.Application")
    
    ' L?y du?ng d?n d?n thu m?c ch?a workbook
    AttachmentPath = ThisWorkbook.Path
    
    ' Thay th? du?ng d?n trong HTML n?u c?n
    HTMLBody = Replace(HTMLBody, "{AttachmentPath}", AttachmentPath)
    
    ' G?i email cho t?ng d?a ch? trong danh s當h
    For Each cell In EmailList
        If cell.Value Like "*@*.?*" Then ' Ki?m tra d?a ch? email h?p l?
            Set OutMail = OutApp.CreateItem(0)
            With OutMail
                .To = cell.Value
                .CC = CCAddress
                .Subject = Subject
                .HTMLBody = HTMLBody
                .Send
            End With
            Set OutMail = Nothing
        End If
    Next cell
    
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " g" & ChrW(7917) & "i mail th" & ChrW(224) & "nh c" & ChrW(244) & "ng!", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", 0 + 64
End Sub

