' Stream: VBA/uf_taofilenpt
' File: uf_taofilenpt.frm

Attribute VB_Name = "uf_taofilenpt"
Attribute VB_Base = "0{06F8F340-2EFC-47FD-B35D-032709CC8543}{4F9193C5-1F8D-4822-BA85-F4F6C3CFE077}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Public Sub CommandButton1_Click()
Call taofile_nhapdslk1
End Sub

'Khai bao ban dau cho userform
Public Sub UserForm_Initialize()
'tao danh sach chon cho cb_phonglamviec
uf_taofilenpt.cb_phong.List = Sheets("tenphong_pt").Range("G2:G4").Value
End Sub
'tao danh sach phu thuoc
Public Sub cb_phong_Change()
'lay danh sach ten phu trach voi phong duoc chon
'B1: Lay gia tri cb_tenphong vao o I2
Sheets("tenphong_pt").Activate
Sheets("tenphong_pt").Range("I2").Value = uf_taofilenpt.cb_phong.Value
'B2: Chay cau lenh loc danh sach
Call Loc_danh_sach
'B3: tim dong cuoi cot K
Dim lr As Long
lr = Sheets("tenphong_pt").Range("K" & Rows.Count).End(xlUp).Row
'B4: cau lenh gan danh sach vao cb_ten phu trach
uf_taofilenpt.cb_tenpt.List = Sheets("tenphong_pt").Range("K2:K" & lr).Value
End Sub
