Attribute VB_Name = "omWorksheetFunctions"
Option Explicit
'@lib 2026-10-10 sha=718ad6e7

' Helpers for worksheets: show/hide, and fill rows from a template row.
' Whether a sheet exists: omExcelApplicationFunctions.SheetExists.

' Hides the sheets that are visible and shows the ones that are hidden or very hidden. Excel keeps at least one sheet visible (error 1004).
Public Sub WorksheetsToggleVisible(sheetNames() As String, Optional wb As Workbook = Nothing)
Dim i As Long
Dim sh As Object
Dim screenUpdatingStatus As Boolean

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If Not IsInitialized(sheetNames) Then Exit Sub
    screenUpdatingStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Done
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set sh = wb.Sheets(sheetNames(i))
        If sh.Visible = xlSheetVisible Then
            sh.Visible = xlSheetHidden
        Else
            sh.Visible = xlSheetVisible
        End If
    Next i
Done:
    Application.ScreenUpdating = screenUpdatingStatus
    If Err.Number <> 0 Then Err.Raise Err.Number, "WorksheetsToggleVisible", Err.Description
End Sub

' Makes every worksheet of the workbook (default: the active workbook) visible, also the very hidden ones.
Public Sub SetAllWorksheetsVisible(Optional wb As Workbook = Nothing)
Dim ws As Worksheet

    If wb Is Nothing Then Set wb = ActiveWorkbook
    For Each ws In wb.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
End Sub

' Copies the formulas of the template row to one row (currentRow), column by column, for every column that has a name in the header row.
' Columns without a formula in the template row are left as they are.
Public Sub CopyTemplateByField(ws As Worksheet, currentRow As Long, Optional headerRow As Long = 2, Optional startColumn As Long = 1, Optional templateRow As Long = 1)
Dim col As Long

    col = startColumn
    Do While HasValue(ws.Cells(headerRow, col))
        If Len(ws.Cells(templateRow, col).FormulaR1C1) > 0 Then
            ws.Cells(currentRow, col).FormulaR1C1 = ws.Cells(templateRow, col).FormulaR1C1
        End If
        col = col + 1
    Loop
End Sub

' Copies the whole template row to numberOfRows rows, starting at startRow (formulas and formats).
Public Sub CopyTemplateByRow(ws As Worksheet, startRow As Long, Optional numberOfRows As Long = 100, Optional templateRow As Long = 1)
    ws.Rows(templateRow).Copy Destination:=ws.Rows(startRow & ":" & (startRow + numberOfRows - 1))
    Application.CutCopyMode = False
End Sub

Private Function HasValue(c As Range) As Boolean
    If IsError(c.Value) Then
        HasValue = True
    Else
        HasValue = (Len(CStr(c.Value)) > 0)
    End If
End Function

Private Function IsInitialized(arr() As String) As Boolean
    On Error Resume Next
    IsInitialized = (UBound(arr) >= LBound(arr))
    If Err.Number <> 0 Then IsInitialized = False
End Function
