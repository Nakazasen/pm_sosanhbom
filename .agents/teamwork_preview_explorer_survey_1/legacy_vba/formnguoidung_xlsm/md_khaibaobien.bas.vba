' Stream: VBA/md_khaibaobien
' File: md_khaibaobien.bas

Attribute VB_Name = "md_khaibaobien"
Public FSO As Object
Public myfile As Object
Public files As Object
'--khai bao duong dan
Public duongdan As String
Public myPath As String
Public myfolderPath As String
'--khai bao ten file
Public thisfile As String
Public tenfileth As String
Public tenwbth As String
Public tenwbpt  As String
Public tenfileR3 As String 'ten file r3 trong thu muc ss bom
Public tenfilePLM As String ' ten file plm trong thu muc ss bom
Public OldName As String, NewName As String
'--khai bao wb
Public wb_Bom As Workbook 'wb_Bom la wb file Bom
Public wb_Plm As Workbook 'wb file plm
Public wb_R3 As Workbook 'wb file r3
Public wb_KQ As Workbook
Public wb_file As Workbook
Public wbth As Workbook
Public wbpt As Workbook
'--khai bao ws
Public ws_Plm_R3 As Worksheet
Public ws_Cttt As Worksheet 'ws_cttt la ws cttt file bom
Public ws_Plm As Worksheet 'ws plm file bom
Public ws_Plm_old As Worksheet 'ws plm_old file bom
Public ws_R3 As Worksheet 'ws r3 file bom
Public ws_cttt_total As Worksheet 'ws cttt_total
Public ws_lichsu As Worksheet 'ws lich su cua file tonghop
Public ws_link As Worksheet
Public ws_thdl As Worksheet
Public ws_duongdan As Worksheet 'sheet duong dan cua file th
'--khai bao dong cuoi
Public lr_MSI As Long 'dong cuoi cua sheet MSI trong file phu trach
Public lr_BosungCTTT As Long 'dong cuoi cua sheet BosungCTTT trong file phu trach
Public lr_kq As Long
Public lr_file As Long
Public lr_r3 As Long
Public lr_plm As Long
