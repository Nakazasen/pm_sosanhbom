Attribute VB_Name = "uf_jig"
Attribute VB_Base = "0{3B787C97-24D3-4793-9824-65BEBF96AC19}{27E6BC5B-3DB9-4F5A-B9AF-361667108879}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub cmd_taijig_Click()
Call Focus(True)
    Dim wsJIG As Worksheet
    Dim wbSource As Workbook
    Dim wsName As String 'bien luu gia tri nguoi dung lua chon cb1
    Dim ws As Worksheet 'ws cua file List JIG thay doi, khi bo sung ma hang.xlsx
    Dim sourceFile As String
    Dim lastRowSource As Long
    Dim lastRowDest As Long
    Dim rngSource As Range
    Dim destRange As Range
    'Dim rngYesNo As Range
    Dim destYesNo As Range
    
    ' Xac ??nh sheet PLM
    Set wsJIG = ThisWorkbook.Sheets("List JIG")
     wsName = cb1.Value 'gan jig nguoi dung lua chon vao bien wsName.
     
    ' Lay lien ket file tu o V24
    sourceFile = wsJIG.Range("V24").Value
      'Clear noi dung cot A3 den H50
    On Error Resume Next
    wsJIG.Range("A3:H50").UnMerge ' Bo tron o trong pham vi A3:H50
    wsJIG.Range("A3:H50").ClearContents
    ' Mo file nguon de tien hanh viec so sanh
    On Error Resume Next
    Set wbSource = Workbooks.Open(sourceFile)
    On Error GoTo 0
    
    ' Ki?m tra neu file nguon khong mo duoc
    If wbSource Is Nothing Then
        CreateObject("WScript.Shell").Popup "Kh" & ChrW(244) & "ng th" & ChrW(7875) & " m" & ChrW(7903) & " file", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 64
        Exit Sub
    End If
    
    '----
    
        ' Kiem tra neu sheet ma nguoi dung lua chon trong list jig co ton tai trong file List JIG thay doi, khi bo sung ma hang.xlsx
    On Error Resume Next
    Set ws = wbSource.Sheets(wsName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet '" & wsName & "' khong ton tai!", vbExclamation, "NG!!!"
    Else
        ' Chuyen den sheet tuong ung
        'ws.Activate
            ' Xac dinh dong cuoi cung cua du lieu trong file nguon
            
        lastRowSource = 50
        ' Xac dinh pham vi du lieu can sao chep
        Set rngSource = ws.Range("A3:G" & lastRowSource)
        'Set rngYesNo = ws.Range("V37:V40")
        Set destRange = wsJIG.Range("A3") ' Vi tri dan du lieu trong wsJIG
        'Set destYesNo = wsJIG.Range("V37") ' Vi tri dan du lieu trong wsJIG
        ' Sao chep du lieu + giu nguyen dinh dang
        rngSource.Copy
        destRange.PasteSpecial Paste:=xlPasteAll  ' Sao chep tat ca (ca gia tri, cong thuc, dinh dang)
        destRange.PasteSpecial Paste:=xlPasteColumnWidths ' Sao chep ca do rong cot
        'rngYesNo.Copy
        'destYesNo.PasteSpecial Paste:=xlPasteAll  ' Sao chep tat ca (ca gia tri, cong thuc, dinh dang)
        'destYesNo.PasteSpecial Paste:=xlPasteColumnWidths ' Sao chep ca do rong cot
        wsJIG.Columns("A:G").AutoFit
    ' Ta che do sao chep
    Application.CutCopyMode = False
    
    ' Dong file nguon ma khong luu thay doi
    wbSource.Close SaveChanges:=False

    ' Thong bao hoan thanh
    MsgBox "Du lieu da duoc sao chep vao 'List JIG'!", vbInformation, "Hoan thanh!"
    uf_jig.Hide
    End If
    '----
    
    Call Focus(False)
End Sub

Private Sub cmd_xoajig_Click()
Call Focus(True)
    Dim wsJIG As Worksheet
    
    ' Xac ??nh sheet PLM
    Set wsJIG = ThisWorkbook.Sheets("List JIG")
     
      'Clear noi dung cot A3 den H50
    On Error Resume Next
    wsJIG.Range("A3:H50").UnMerge ' Bo tron o trong pham vi A3:H50
    wsJIG.Range("A3:H50").ClearFormats ' X �?nh d?ng (bao g?m c? m炒 n?n)
    wsJIG.Range("A3:H50").ClearContents ' X n?i dung d? li?u
    uf_jig.Hide
      
    Call Focus(False)
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    Set ws = ThisWorkbook.Sheets("List JIG")
    Set rng = ws.Range("V3:V17") ' Vg d? li?u

    For Each cell In rng
        cb1.AddItem cell.Value
    Next cell
End Sub
