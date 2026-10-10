Attribute VB_Name = "omListObjectFunctions"
Option Explicit
'@lib 2026-10-10 sha=d08630b5

' Helpers for Excel tables (ListObject). Needs omOnedriveFunctions and omExcelRecordsetFunctions (for OpenListObjectAsRecordset)
' and the reference "Microsoft ActiveX Data Objects 6.x Library".

' The largest value in a column of the table; Null when the table has no rows.
Public Function GetMaxInTableColumn(tbl As ListObject, columnName As String) As Variant
    If tbl.DataBodyRange Is Nothing Then
        GetMaxInTableColumn = Null
    Else
        GetMaxInTableColumn = WorksheetFunction.Max(tbl.ListColumns(columnName).DataBodyRange)
    End If
End Function

' The table as an ADO recordset (read only), with an optional WHERE / ORDER BY text (see OpenWorksheetRangeAsRecordset).
' It reads the saved file of the workbook that holds the table, so unsaved changes are not in it.
' Close it with omExcelRecordsetFunctions.CloseRecordset.
Public Function OpenListObjectAsRecordset(tbl As ListObject, Optional whereClause As String = "", Optional orderByClause As String = "") As ADODB.Recordset
Dim wb As Workbook

    Set wb = tbl.Parent.Parent
    Set OpenListObjectAsRecordset = omExcelRecordsetFunctions.OpenWorksheetRangeAsRecordset(omOnedriveFunctions.GetWorkbookLocalFilename(wb), tbl.Parent.Name, tbl.Range.Address(False, False), False, whereClause, orderByClause)
End Function

' The position of the column in the table (1 = first column), or -1 when there is no such column.
Public Function GetColumnNumberByName(tbl As ListObject, columnName As String) As Long
    On Error GoTo NotFound
    GetColumnNumberByName = tbl.ListColumns(columnName).Index
    Exit Function
NotFound:
    GetColumnNumberByName = -1
End Function

' The number of the first data row (1 = first row under the header) whose value in the column equals searchValue, or -1 when
' the column or the value is not there. Cells with an error value are skipped.
Public Function FindRowByValue(tbl As ListObject, columnName As String, searchValue As Variant) As Long
Dim columnNumber As Long
Dim values As Variant
Dim i As Long
Dim v As Variant

    FindRowByValue = -1
    columnNumber = GetColumnNumberByName(tbl, columnName)
    If columnNumber = -1 Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    values = tbl.ListColumns(columnNumber).DataBodyRange.Value
    If Not IsArray(values) Then
        If IsMatch(values, searchValue) Then FindRowByValue = 1
        Exit Function
    End If
    For i = 1 To UBound(values, 1)
        v = values(i, 1)
        If IsMatch(v, searchValue) Then
            FindRowByValue = i
            Exit Function
        End If
    Next i
End Function

Private Function IsMatch(v As Variant, searchValue As Variant) As Boolean
    If IsError(v) Then Exit Function
    On Error Resume Next
    IsMatch = (v = searchValue)
    If Err.Number <> 0 Then IsMatch = False
End Function

' Sorts the table on one or two columns. Raises error 5 when a column does not exist.
Public Sub SortListObjectByColumn(tbl As ListObject, column1Name As String, Optional column2Name As String = "", Optional sort1Order As XlSortOrder = xlAscending, Optional sort2Order As XlSortOrder = xlAscending)
Dim column1Number As Long
Dim column2Number As Long

    column1Number = GetColumnNumberByName(tbl, column1Name)
    If column1Number = -1 Then Err.Raise 5, "SortListObjectByColumn", "Column '" & column1Name & "' not found."
    If Len(column2Name) > 0 Then
        column2Number = GetColumnNumberByName(tbl, column2Name)
        If column2Number = -1 Then Err.Raise 5, "SortListObjectByColumn", "Column '" & column2Name & "' not found."
    End If
    tbl.Sort.SortFields.Clear
    tbl.Sort.SortFields.Add Key:=tbl.ListColumns(column1Number).Range, SortOn:=xlSortOnValues, Order:=sort1Order, DataOption:=xlSortNormal
    If column2Number > 0 Then
        tbl.Sort.SortFields.Add Key:=tbl.ListColumns(column2Number).Range, SortOn:=xlSortOnValues, Order:=sort2Order, DataOption:=xlSortNormal
    End If
    With tbl.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

' The number of the data row (1 = first row under the header) that holds the active cell, or -1 when the active cell is not in the table.
Public Function GetListObjectRowFromActiveCell(tbl As ListObject) As Long
    GetListObjectRowFromActiveCell = -1
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Not ActiveSheet Is tbl.Parent Then Exit Function
    If Intersect(ActiveCell, tbl.DataBodyRange) Is Nothing Then Exit Function
    GetListObjectRowFromActiveCell = ActiveCell.Row - tbl.HeaderRowRange.Row
End Function
