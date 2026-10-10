Attribute VB_Name = "omExcelApplicationFunctions"
Option Explicit
'@lib 2026-10-10 sha=369439da

' Switch the Excel application settings off for a fast macro and back on afterwards, plus a full recalculation and SheetExists.

' Calculation manual, no screen updating, no alerts, no events. calculateBefore: recalculate everything first.
Public Sub DisableAllApplicationEvents(Optional calculateBefore As Boolean = True)
    Application.Cursor = xlWait
    If calculateBefore Then
        Application.CalculateFull
        DoEvents
    End If
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .CutCopyMode = False
    End With
    ResetStatusBar
    DoEvents
End Sub

' Undoes DisableAllApplicationEvents. calculateAfter: recalculate everything at the end.
Public Sub EnableAllApplicationEvents(Optional calculateAfter As Boolean = True)
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .CutCopyMode = False
    End With
    ResetStatusBar
    If calculateAfter Then
        Application.CalculateFull
    End If
    DoEvents
    Application.Cursor = xlDefault
    Application.StatusBar = False
End Sub

Public Sub ResetStatusBar()
    Application.DisplayStatusBar = True
    Application.StatusBar = False
End Sub

' Screen updating and events on (enabled = True) or off (enabled = False), nothing else; the light version of Disable/EnableAllApplicationEvents.
Public Sub ApplicationEventMode(enabled As Boolean)
    Application.ScreenUpdating = enabled
    Application.EnableEvents = enabled
End Sub

' Automatic calculation, waits for asynchronous queries (for example CUBEVALUE) and recalculates everything.
Public Sub FullRecalcWithCubeSupport()
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateUntilAsyncQueriesDone
    Application.CalculateFull
    Do Until Application.CalculationState = xlDone
        DoEvents
    Loop
End Sub

' True when the workbook (default: the workbook that holds this code) has a worksheet with this name (not case sensitive).
Public Function SheetExists(SheetName As String, Optional wb As Workbook = Nothing) As Boolean
Dim sh As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    For Each sh In wb.Worksheets
        If StrComp(sh.Name, SheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next sh
End Function
