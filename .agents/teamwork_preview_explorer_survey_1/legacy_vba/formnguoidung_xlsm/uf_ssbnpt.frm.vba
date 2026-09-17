' Stream: VBA/uf_ssbnpt
' File: uf_ssbnpt.frm

Attribute VB_Name = "uf_ssbnpt"
Attribute VB_Base = "0{33827C17-8D80-4510-A73E-3C18DA9F8A41}{0DA33405-8490-4AC9-AEF0-368787CA8C79}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Sub cb_kqss_Click()
tenwbpt = ThisWorkbook.Name
Set wbpt = Workbooks(tenwbpt)
Set ws_Cttt = wbpt.Worksheets("CTTT")
ws_Cttt.Range("Q2") = "OK"
ws_Cttt.Range("Q2").Interior.Color = vbGreen
End Sub

Private Sub cb_sosanh_Click()
Call sosanhBompt
End Sub

Private Sub cb_xkq_Click()
tenwbpt = ThisWorkbook.Name
Set wbpt = Workbooks(tenwbpt)
Set ws_Cttt = wbpt.Worksheets("CTTT")
ws_Cttt.Range("Q2").ClearContents
ws_Cttt.Range("Q2").Interior.Color = RGB(255, 255, 255)
End Sub
