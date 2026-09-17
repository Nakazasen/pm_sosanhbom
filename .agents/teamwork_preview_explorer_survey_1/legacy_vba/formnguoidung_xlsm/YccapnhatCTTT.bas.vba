' Stream: VBA/YccapnhatCTTT
' File: YccapnhatCTTT.bas

Attribute VB_Name = "YccapnhatCTTT"
Sub Focus(ByVal Flag As Boolean)
    With Application
        .EnableEvents = Not Flag
        .ScreenUpdating = Not Flag
        .Calculation = IIf(Flag, xlCalculationManual, xlCalculationAutomatic)
    End With
End Sub


