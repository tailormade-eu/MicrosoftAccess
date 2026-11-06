Attribute VB_Name = "tstControlFunctions"
Option Compare Database
Option Explicit

' Tests for omControlFunctions.ResolveDefaultValue / SetDefaultValue and omStringFunctions.NotIsNull.
' Run from the Immediate window:  ?tstControlFunctions.RunAll
' Returns "OK (n tests)" or one line per failing case.

Private mFailures As String
Private mTests As Long

Private Sub Check(ByVal condition As Boolean, ByVal what As String)
    mTests = mTests + 1
    If Not condition Then mFailures = mFailures & what & vbCrLf
End Sub

' Same text and same type
Private Function Same(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNull(a) Or IsNull(b) Then
        Same = (IsNull(a) And IsNull(b))
    Else
        Same = (VarType(a) = VarType(b)) And (CStr(a) = CStr(b))
    End If
End Function

Public Function RunAll() As String
Dim r As Variant

    mFailures = ""
    mTests = 0

    ' NotIsNull: only Null is Null
    Check Not omStringFunctions.NotIsNull(Null), "NotIsNull(Null) must be False"
    Check omStringFunctions.NotIsNull(""), "NotIsNull(zero-length string) must be True"
    Check omStringFunctions.NotIsNull("   "), "NotIsNull(spaces) must be True"
    Check omStringFunctions.NotIsNull(0), "NotIsNull(0) must be True"
    Check omStringFunctions.NotIsNull("abc"), "NotIsNull(text) must be True"

    ' a value that is there stays: Null is the only empty
    Check Same(omControlFunctions.ResolveDefaultValue("abc", "0"), "abc"), "a text value was replaced by the default"
    Check Same(omControlFunctions.ResolveDefaultValue("", "0"), ""), "a zero-length string was replaced by the default"
    Check Same(omControlFunctions.ResolveDefaultValue("   ", "0"), "   "), "spaces were replaced by the default"
    Check Same(omControlFunctions.ResolveDefaultValue(5, "0"), 5), "a number was replaced by the default"
    Check Same(omControlFunctions.ResolveDefaultValue(0, "1"), 0), "the number 0 was replaced by the default"

    ' Null gets the default
    Check Same(omControlFunctions.ResolveDefaultValue(Null, "30"), "30"), "Null did not get the default"
    Check Same(omControlFunctions.ResolveDefaultValue(Null, "-1"), "-1"), "Null did not get the default -1"

    ' a date default (the property gives it with # around it, or in quotes) comes back as a date
    r = omControlFunctions.ResolveDefaultValue(Null, "#2026-01-15#")
    Check VarType(r) = vbDate And r = DateSerial(2026, 1, 15), "a #date# default was not returned as a date"
    r = omControlFunctions.ResolveDefaultValue(Null, """2026-02-20""")
    Check VarType(r) = vbDate And r = DateSerial(2026, 2, 20), "a quoted date default was not returned as a date"

    ' no default at all
    Check IsNull(omControlFunctions.ResolveDefaultValue(Null, Null)), "Null with a Null default must stay Null"

    If Len(mFailures) = 0 Then
        RunAll = "OK (" & mTests & " tests)"
    Else
        RunAll = mFailures
    End If
End Function
