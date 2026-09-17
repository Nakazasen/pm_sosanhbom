' Stream: VBA/hamcon
' File: hamcon.bas

Attribute VB_Name = "hamcon"
Sub Loai_bo_Duplicate()
Range("A1:B39").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("G1"), Unique:=True
End Sub
Sub Loc_danh_sach()
Attribute Loc_danh_sach.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("A1:B39").AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Range("adf_DK"), CopyToRange:=Range("K1"), Unique:=True
End Sub
