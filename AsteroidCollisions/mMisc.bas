Attribute VB_Name = "mMisc"
Option Explicit

' Define the name of this class/module for error-trap reporting.
Private Const mc_strModuleName As String = "mMisc"


Private Const g_sngPIDivideBy180 = 0.0174533!
Private Const g_sng180DivideByPI = 57.29578!

Public Sub DrawCrossHairs(CurrentForm As Form)

    CurrentForm.DrawStyle = vbDot ' or vbSolid
    CurrentForm.DrawWidth = 1
    
    ' Draw vertical line.
    CurrentForm.ForeColor = RGB(0, 64, 64)
    CurrentForm.Line (CurrentForm.ScaleWidth / 2, 0)-(CurrentForm.ScaleWidth / 2, CurrentForm.ScaleHeight)

    ' Draw horizontal line.
    CurrentForm.ForeColor = RGB(0, 42, 42)
    CurrentForm.Line (0, CurrentForm.ScaleHeight / 2)-(CurrentForm.ScaleWidth, CurrentForm.ScaleHeight / 2)
    
End Sub
Public Function ConvertDeg2Rad(Degress As Single) As Single

    ' Converts Degrees to Radians
    ConvertDeg2Rad = Degress * g_sngPIDivideBy180
    
End Function

Public Function ConvertRad2Deg(Radians As Single) As Single

    ' Converts Radians to Degrees
    ConvertRad2Deg = Radians * g_sng180DivideByPI

End Function



Public Function GetRNDNumberBetween(Min As Variant, Max As Variant) As Single

    GetRNDNumberBetween = (Rnd * (Max - Min)) + Min

End Function

