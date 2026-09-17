Attribute VB_Name = "xoadlcu"
Sub xoa_dlold()
Call Focus(True)
    tenwbth = ActiveWorkbook.Name
    Set wbth = Workbooks(tenwbth)
    Set ws_thdl = wbth.Sheets("tonghopdl")
    Set ws_downloadR3 = wbth.Sheets("Download_R3_tudong")
    Set ws_tonghopMsi = wbth.Sheets("tonghopMSI_7980_7990")
    lr_kq = ws_thdl.Range("C" & Rows.Count).End(xlUp).Row + 1
    ws_thdl.Range("A3:G" & lr_kq).ClearContents
    ws_tonghopMsi.Range("A2:B36").ClearContents
    ws_tonghopMsi.Range("D2:L36").ClearContents
    ws_tonghopMsi.Range("A49:L73").ClearContents
    ws_downloadR3.Range("A2:B50").ClearContents
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " x" & ChrW(243) & "a d" & ChrW(7919) & " li" & ChrW(7879) & "u linh ki" & ChrW(7879) & "n c" & ChrW(361) & " trong sheet t" & ChrW(7893) & "ng h" & ChrW(7907) & "p", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o x" & ChrW(243) & "a th" & ChrW(224) & "nh c" & ChrW(244) & "ng", 0 + 64
Call Focus(False)
End Sub


