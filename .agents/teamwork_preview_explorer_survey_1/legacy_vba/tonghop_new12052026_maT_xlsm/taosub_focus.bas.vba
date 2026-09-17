' Stream: VBA/taosub_focus
' File: taosub_focus.bas

Attribute VB_Name = "taosub_focus"
Sub Focus(ByVal Flag As Boolean)
    With Application
        .EnableEvents = Not Flag
        .ScreenUpdating = Not Flag
        .Calculation = IIf(Flag, xlCalculationManual, xlCalculationAutomatic)
    End With
End Sub
