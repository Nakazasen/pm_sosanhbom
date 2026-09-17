Attribute VB_Name = "md_khaibaobienlocal_mainmenu"
Public wbth As Workbook
Public FSO As Object
Public duongdan As String
Public tenmay As String
Public ws_link As Worksheet
Public ws_thdl As Worksheet
Public tenwbth As String
Public OldName As String, NewName As String
Public wb_KQ As Workbook
Public wb_file As Workbook
Public myPath As String
Public lr_kq As Long
Public lr_file As Long
Public myfolderPath As String
Public myfile As Object
Public files As Object
Public wb_Bom As Workbook 'wb_Bom la wb file Bom
Public ws_Cttt As Worksheet 'ws_cttt la ws cttt file bom
Public ws_Plm As Worksheet 'ws plm file bom
Public ws_R3 As Worksheet 'ws r3 file bom
Public ws_cttt_total As Worksheet 'ws cttt_total
Public ws_msi As Worksheet 'ws msi
Public ws_Tongket As Worksheet 'ws tongket
Public tenfileR3 As String 'ten file r3 trong thu muc ss bom
Public tenfilePLM As String ' ten file plm trong thu muc ss bom
Public wb_Plm As Workbook 'wb file plm
Public wb_R3 As Workbook 'wb file r3
Public wb_FixSerial As Workbook 'wb file FixSerial
Public ws_msimay As Worksheet
Public ws_msiunit As Worksheet
Public duongdan_fixserial As String
Public tenfile_fixserial As String
Public lr_fix As Long 'dong cuoi cua file fix serial
Public mk_cttt As String

Sub mainPLM()
uf_plm.Show
End Sub
