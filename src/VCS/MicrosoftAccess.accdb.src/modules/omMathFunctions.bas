Attribute VB_Name = "omMathFunctions"
Option Compare Database
Option Explicit
'@lib 2026-11-05 sha=d29d1614

' Rounding the way people count money.
'
' VBA's Round() rounds half to even ("banker's rounding"): Round(2.5) = 2. And a value such as 1.005 is stored as
' 1.00499999999999989..., so Round(1.005, 2) gives 1 instead of 1.01. RoundHalfUp rounds a .5 away from zero, and it first
' adds a tiny correction for that binary representation error.
'
' Public, so a form or report expression can call it:  =RoundHalfUp([Total], 2)

Public Function RoundHalfUp(ByVal value As Double, ByVal decimals As Integer) As Double
    Dim factor As Double
    factor = 10 ^ decimals

    ' correct the floating point error (1.005 is stored as 1.004999...)
    Dim corrected As Double
    corrected = value * factor + 0.0000000001 * Sgn(value)

    ' round half up (away from zero)
    If corrected >= 0 Then
        RoundHalfUp = Int(corrected + 0.5) / factor
    Else
        RoundHalfUp = -Int(-corrected + 0.5) / factor
    End If
End Function
