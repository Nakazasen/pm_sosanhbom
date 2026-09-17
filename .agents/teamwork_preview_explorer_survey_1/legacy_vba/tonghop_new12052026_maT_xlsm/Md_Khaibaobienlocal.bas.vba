' Stream: VBA/Md_Khaibaobienlocal
' File: Md_Khaibaobienlocal.bas

Attribute VB_Name = "Md_Khaibaobienlocal"
Public wbth As Workbook
Public FSO As Object
Public myfile As Object
Public files As Object
Public folder_phutrach As String
Public folder_capnhat As String
Public folder_tc14backup As String
Public folder_taifile As String
Public wb_KQ As Workbook
Public wb_file As Workbook
Public wb_Bom As Workbook 'wb_Bom la wb file Bom
Public wb_Plm As Workbook 'wb file plm
Public wb_R3 As Workbook 'wb file r3
Public wb_npt As Workbook 'wb npt
Public wb_thdl_new As Workbook
Public ws_link As Worksheet
Public ws_thdl As Worksheet
Public ws_Cttt As Worksheet 'ws_cttt la ws cttt file bom
Public ws_Plm As Worksheet 'ws plm file bom
Public ws_Plm_old As Worksheet 'ws plm_old file bom
Public ws_tonghopMsi As Worksheet 'ws Msi file tong hop
Public ws_R3 As Worksheet 'ws r3 file bom
Public ws_cttt_total As Worksheet 'ws cttt_total
Public ws_Lichsu As Worksheet 'ws lichsu trong file tonghop
Public ws_bosungCTTT As Worksheet 'ws bo sung chi thi thao tac
Public ws_bolocbom As Worksheet 'ws bo loc bom
Public ws_downloadR3 As Worksheet 'ws download r3 cua workbook tong hop
Public ws_thlabel_7980 As Worksheet 'ws 7980 cua workbook file thop
Public ws_label_7980 As Worksheet 'ws 7980 cua workbook nguoi npt
Public ws_tongket_CTTT As Worksheet 'ws tong ket cua workbook bom
Public duongdan As String
Public myfolderPath As String
Public tenmay As String
Public tennpt As String
Public tenwbth As String
Public tenform_thT As String
Public tenfileR3 As String 'ten file r3 trong thu muc ss bom
Public tenfilePLM As String ' ten file plm trong thu muc ss bom
Public TC14Full As String 'ten file plm full
Public OldName As String, NewName As String
Public myPath As String
Public lr_kq As Long
Public lr_file As Long
Public lr_MSI As Long 'dong cuoi cua sheet MSI trong file phu trach
Public lr_BosungCTTT As Long 'dong cuoi cua sheet BosungCTTT trong file phu trach
Public lr_lichsu As Long 'dong cuoi cua sheet lichsu
Public lr_cttt As Long 'dong cuoi cua sheet cttt cua file bom
Public lr_label_7980 As Long ' dong cuoi cua sheet label_7980 cua file nguoi pt
Public lr_label_7990 As Long ' dong cuoi cua sheet label_7990 cua file nguoi pt
Public lr_th_label7980 As Long ' dong cuoi cua sheet tonghopMSI_7980 cua file tong hop
Public lr_th_label7990 As Long ' dong cuoi cua sheet tonghopMSI_7990 cua file tong hop
Public mk_cttt As String



