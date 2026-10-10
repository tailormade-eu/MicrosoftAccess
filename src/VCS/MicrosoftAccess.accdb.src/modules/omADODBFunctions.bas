Attribute VB_Name = "omADODBFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=878adaa6


Public Function ADODBCreateCopyRecord(rsSrc As ADODB.Recordset, rsDst As ADODB.Recordset, Optional PrimaryKeyName As String = "", Optional LinkName As String = "", Optional LinkId As Long = 0, Optional ignoreFields As String) As Long
Dim fld As ADODB.Field

    ADODBCreateCopyRecord = 0
    rsDst.AddNew
        ADODBCopyRecord rsSrc, rsDst, PrimaryKeyName, LinkName, LinkId, ignoreFields
    If Len(PrimaryKeyName) > 0 Then
        ADODBCreateCopyRecord = rsDst(PrimaryKeyName).Value
    End If
End Function

Public Sub ADODBCopyRecord(rsSrc As ADODB.Recordset, rsDst As ADODB.Recordset, Optional PrimaryKeyName As String = "", Optional LinkName As String = "", Optional LinkId As Long = 0, Optional ignoreFields As String)
Dim fld As ADODB.Field

        ignoreFields = PrimaryKeyName & "," & Nz(ignoreFields)
    For Each fld In rsSrc.Fields
        If Not omStringFunctions.ContainsString(ignoreFields, fld.Name, ",") Then
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
End Sub

' Closes the recordset when it is open (also when it is at EOF) and sets the variable to Nothing; safe to call when it is already Nothing.
Public Sub SetRecordsetNothing(ByRef rs As ADODB.Recordset)
    If rs Is Nothing Then Exit Sub
    On Error Resume Next
    If rs.State <> adStateClosed Then rs.Close
    On Error GoTo 0
    Set rs = Nothing
End Sub
