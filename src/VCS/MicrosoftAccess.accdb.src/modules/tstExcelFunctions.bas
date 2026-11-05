Attribute VB_Name = "tstExcelFunctions"
Option Compare Database
Option Explicit

' Tests for omExcelFunctions (so far ConvertToImage). Needs Excel.
' Run from the Immediate window:  ?tstExcelFunctions.RunAll
' Returns "OK (n tests)" or one line per failing case.

Private mFailures As String
Private mTests As Long

Private Sub Check(testName As String, actual As String, expected As String)
    mTests = mTests + 1
    ' Excel may write the function name in capitals; the rest (the file name!) is compared with case
    If UCase$(Left$(actual, 7)) = "=IMAGE(" Then actual = "=IMAGE(" & Mid$(actual, 8)
    If StrComp(actual, expected, vbBinaryCompare) <> 0 Then
        mFailures = mFailures & testName & ": got [" & actual & "] expected [" & expected & "]" & vbCrLf
    End If
End Sub

Public Function RunAll() As String
Dim xl As Object
Dim wb As Object
Dim ws As Object

    mFailures = ""
    mTests = 0
    On Error GoTo Fail
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)

    ' default: the text is lower-cased first (findString has to be lower case too)
    ws.Cells(1, 1).Value = "url"
    ws.Cells(2, 1).Value = "D:\Pics\ABC.PNG"
    ws.Cells(3, 1).Value = "D:\Pics\Def.png"
    omExcelFunctions.ConvertToImage ws, 1, "d:\pics\", "https://example.invalid/"
    Check "ConvertToImage default row 2", CStr(ws.Cells(2, 1).Formula), "=IMAGE(""https://example.invalid/abc.png"")"
    Check "ConvertToImage default row 3", CStr(ws.Cells(3, 1).Formula), "=IMAGE(""https://example.invalid/def.png"")"

    ' setLowerCase:=False keeps the case of the file name (and findString is matched as written)
    ws.Cells(2, 1).Value = "D:\Pics\ABC.PNG"
    ws.Cells(3, 1).Value = "D:\Pics\Def.png"
    omExcelFunctions.ConvertToImage ws, 1, "D:\Pics\", "https://example.invalid/", False
    Check "ConvertToImage keep case row 2", CStr(ws.Cells(2, 1).Formula), "=IMAGE(""https://example.invalid/ABC.PNG"")"
    Check "ConvertToImage keep case row 3", CStr(ws.Cells(3, 1).Formula), "=IMAGE(""https://example.invalid/Def.png"")"

Done:
    On Error Resume Next
    wb.Close False
    xl.Quit
    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing
    If mFailures = "" Then
        RunAll = "OK (" & mTests & " tests)"
    Else
        RunAll = mFailures
    End If
    Exit Function
Fail:
    mFailures = mFailures & "ERROR " & Err.Number & ": " & Err.Description & vbCrLf
    Resume Done
End Function
