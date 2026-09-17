' Stream: VBA/uf_plm
' File: uf_plm.frm

Attribute VB_Name = "uf_plm"
Attribute VB_Base = "0{A5D14DE8-4596-4488-9E26-C9E749BF3C02}{C395EA20-DE9D-4209-9440-180086E64111}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Const WM_SETTEXT = &HC
#If Win64 Then
    Private Declare PtrSafe Function DefWindowProcW Lib "user32" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#Else
    Private Declare Function DefWindowProcW Lib "user32" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

Private Sub cmd_locdl_Click()
Call locdl_ss_plm
End Sub

Private Sub cmd_timkiemunit_Click()
Call CopyDataFromLinkedFile
End Sub

Private Sub UserForm_Initialize()
Dim HWnd&
    HWnd = FindWindow("ThunderDFrame", Caption)
    DefWindowProcW HWnd, WM_SETTEXT, 0, StrPtr("Ch" & ChrW(432) & ChrW(417) & "ng tr" & ChrW(236) & "nh h" & ChrW(7895) & " tr" & ChrW(7907) & " t" & ChrW(236) & "m ki" & ChrW(7871) & "m th" & ChrW(244) & "ng tin trong PLM")
End Sub
