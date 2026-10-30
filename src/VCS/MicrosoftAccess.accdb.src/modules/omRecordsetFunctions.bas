Attribute VB_Name = "omRecordsetFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-30 sha=f4ef88a7


Public Function CopyRecord(rsSrc As ADODB.Recordset, rsDst As ADODB.Recordset, Optional PrimaryKeyName As String = "", Optional LinkName As String = "", Optional LinkId As Long = 0) As Long
Dim fld As ADODB.Field

    CopyRecord = 0
    rsDst.AddNew
    For Each fld In rsSrc.Fields
        If fld.Name <> PrimaryKeyName Then
            If fld.Name = LinkName And LinkId <> 0 Then
                rsDst(fld.Name).Value = LinkId
            Else
                'On Error Resume Next
                If Left(fld.Name, 2) <> "s_" And fld.Type <> 204 And fld.Name <> "SSMA_TimeStamp" Then 'timestamp
                    rsDst(fld.Name).Value = rsSrc(fld.Name).Value
                End If

                'On Error GoTo 0
            End If
        End If
    Next
    rsDst.Update
    If Len(PrimaryKeyName) > 0 Then
        CopyRecord = rsDst(PrimaryKeyName).Value
    End If
End Function

Public Function GetList(SQL As String, Optional ColumnDelimiter As String = ", ", Optional RowDelimiter As String = vbCrLf) As String
'PURPOSE: to return a combined string from the passed query
'ARGS:
'   1. SQL is a valid Select statement
'   2. ColumnDelimiter is the character(s) that separate each column
'   3. RowDelimiter is the character(s) that separate each row
'RETURN VAL: Concatenated list
'DESIGN NOTES:
'EXAMPLE CALL: =GetList("Select Col1,Col2 From Table1 Where Table1.Key = " & OuterTable.Key)

Const PROCNAME = "GetList"
'Const adClipString = 2
Dim oConn As ADODB.connection
Dim oRS As ADODB.Recordset
Dim sResult As String

On Error GoTo ProcErr

Set oConn = CurrentProject.connection
Set oRS = oConn.Execute(SQL)

sResult = oRS.GetString(adClipString, -1, ColumnDelimiter, RowDelimiter)

If Right(sResult, Len(RowDelimiter)) = RowDelimiter Then
    sResult = Mid$(sResult, 1, Len(sResult) - Len(RowDelimiter))
End If

GetList = sResult
oRS.Close
' oConn is CurrentProject.Connection, the connection of the application itself: do not close it

CleanUp:
    Set oRS = Nothing
    Set oConn = Nothing

Exit Function
ProcErr:
    ' insert error handler
    Resume CleanUp

End Function

' Copies the values of all fields of one record to the same-named fields of another record (both open recordsets; the destination is
' already in AddNew or edit mode). ignoreFields: field names that are not copied, separated by commas (whole names, not case sensitive).
' It does not call Update, so the caller can set more fields first. See also ADODBCopyRecord (omADODBFunctions), which copies and updates.
Public Sub CopyFields(sourceRecord As Variant, destinationRecord As Variant, Optional ignoreFields As String = "")
Dim i As Long
Dim fieldName As String
Dim ignoreList As String

    ignoreList = "," & Replace(ignoreFields, " ", "") & ","
    For i = 0 To sourceRecord.Fields.Count - 1
        fieldName = sourceRecord.Fields(i).Name
        If InStr(1, ignoreList, "," & fieldName & ",", vbTextCompare) = 0 Then
            destinationRecord(fieldName) = sourceRecord(fieldName)
        End If
    Next i
End Sub
