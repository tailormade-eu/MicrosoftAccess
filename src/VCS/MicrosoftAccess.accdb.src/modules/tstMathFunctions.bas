Attribute VB_Name = "tstMathFunctions"
Option Compare Database
Option Explicit

' Tests for omMathFunctions.
' Run from the Immediate window:  ?tstMathFunctions.RunAll
' Returns "OK (n tests)" or one line per failing case.

Private mFailures As String
Private mTests As Long

Private Sub Check(value As Double, decimals As Integer, expected As String)
    Dim actual As String
    mTests = mTests + 1
    actual = Trim$(Str$(omMathFunctions.RoundHalfUp(value, decimals)))
    If actual <> expected Then
        mFailures = mFailures & "RoundHalfUp(" & Trim$(Str$(value)) & ", " & decimals & "): got [" & actual & "] expected [" & expected & "]" & vbCrLf
    End If
End Sub

Public Function RunAll() As String
    mFailures = ""
    mTests = 0

    Check 1.005, 2, "1.01"          ' stored as 1.00499999...: Round() gives 1
    Check 2.675, 2, "2.68"
    Check 1.255, 2, "1.26"
    Check 0.285, 2, ".29"           ' Str$ writes no leading zero
    Check 2.5, 0, "3"               ' Round() gives 2 (half to even)
    Check 3.5, 0, "4"
    Check -2.5, 0, "-3"             ' away from zero
    Check -1.005, 2, "-1.01"
    Check 123.4549, 2, "123.45"
    Check 123.455, 2, "123.46"
    Check 0, 2, "0"
    Check 1000000.005, 2, "1000000.01"

    If mFailures = "" Then
        RunAll = "OK (" & mTests & " tests)"
    Else
        RunAll = mFailures
    End If
End Function
