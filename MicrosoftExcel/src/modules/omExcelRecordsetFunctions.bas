Attribute VB_Name = "omExcelRecordsetFunctions"
Option Explicit
'@lib 2026-10-10 sha=867eba45

' Reads (and optionally writes) a sheet range of a closed .xlsx file as an ADO recordset.
' Needs the reference "Microsoft ActiveX Data Objects 6.x Library" and the ACE OLE DB provider (installed with Access or the Access Database Engine).

' The range of a sheet as a recordset; the first row of the range holds the field names.
' fileName: the .xlsx file (it must not be open in Excel); SheetName without the $; rangeAddress as in Excel, for example "A:C" or "A1:G200".
' updateable: True = rows can be changed and added; False = read only.
' whereClause / orderByClause: SQL text without the words WHERE / ORDER BY. It is placed in the statement as it is, so never
' build it from text that you did not write yourself.
' Close the recordset with CloseRecordset.
Public Function OpenWorksheetRangeAsRecordset(fileName As String, SheetName As String, rangeAddress As String, Optional updateable As Boolean = False, Optional whereClause As String = "", Optional orderByClause As String = "") As ADODB.Recordset
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim commandText As String

    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    commandText = "SELECT * FROM [" & Replace(SheetName, "]", "") & "$" & rangeAddress & "]"
    If Len(whereClause) > 0 Then commandText = commandText & " WHERE " & whereClause
    If Len(orderByClause) > 0 Then commandText = commandText & " ORDER BY " & orderByClause
    If updateable Then conn.CursorLocation = adUseClient
    conn.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & fileName & ";Extended Properties=""Excel 12.0;HDR=Yes" & IIf(updateable, "", ";IMEX=1") & """;"
    On Error GoTo Failed
    If updateable Then
        rs.Open commandText, conn, adOpenDynamic, adLockOptimistic
    Else
        rs.Open commandText, conn, adOpenDynamic, adLockReadOnly
    End If
    Set OpenWorksheetRangeAsRecordset = rs
    Exit Function
Failed:
    Dim errNumber As Long, errDescription As String
    errNumber = Err.Number
    errDescription = Err.Description
    conn.Close
    Err.Raise errNumber, "OpenWorksheetRangeAsRecordset", errDescription
End Function

' Closes the recordset and sets the variable to Nothing. closeConnection: also close the connection that the recordset was opened on
' (default True: OpenWorksheetRangeAsRecordset opened that connection for this recordset only).
Public Sub CloseRecordset(rs As ADODB.Recordset, Optional closeConnection As Boolean = True)
Dim conn As ADODB.Connection

    If rs Is Nothing Then Exit Sub
    If closeConnection Then Set conn = rs.ActiveConnection
    If rs.State <> adStateClosed Then rs.Close
    Set rs = Nothing
    If Not conn Is Nothing Then
        If conn.State <> adStateClosed Then conn.Close
    End If
End Sub
