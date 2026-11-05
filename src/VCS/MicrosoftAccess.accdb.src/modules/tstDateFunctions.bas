Attribute VB_Name = "tstDateFunctions"
Option Compare Database
Option Explicit

' Tests for omDateFunctions (so far GetTimeStamp).
' Run from the Immediate window:  ?tstDateFunctions.RunAll
' Returns "OK (n tests)" or one line per failing case.

Private mFailures As String
Private mTests As Long

Private Sub Check(testName As String, actual As String, expected As String)
    mTests = mTests + 1
    If StrComp(actual, expected, vbBinaryCompare) <> 0 Then
        mFailures = mFailures & testName & ": got [" & actual & "] expected [" & expected & "]" & vbCrLf
    End If
End Sub

Public Function RunAll() As String
Dim s As String

    mFailures = ""
    mTests = 0

    ' the given moment is used, with leading zeros (hour 9 gives 09, not 9)
    Check "GetTimeStamp early morning", omDateFunctions.GetTimeStamp(DateSerial(2026, 9, 5) + TimeSerial(9, 5, 2)), "20260905_090502"
    Check "GetTimeStamp afternoon", omDateFunctions.GetTimeStamp(DateSerial(2026, 11, 30) + TimeSerial(23, 59, 58)), "20261130_235958"
    Check "GetTimeStamp midnight", omDateFunctions.GetTimeStamp(DateSerial(2026, 1, 2) + TimeSerial(0, 0, 0)), "20260102_000000"

    ' without a parameter: now. Shape yyyymmdd_hhnnss, 15 characters, and not older than a minute
    s = omDateFunctions.GetTimeStamp()
    Check "GetTimeStamp now: length", CStr(Len(s)), "15"
    Check "GetTimeStamp now: underscore", Mid$(s, 9, 1), "_"
    Check "GetTimeStamp now: this year", Left$(s, 4), Format$(Now, "yyyy")

    If mFailures = "" Then
        RunAll = "OK (" & mTests & " tests)"
    Else
        RunAll = mFailures
    End If
End Function
