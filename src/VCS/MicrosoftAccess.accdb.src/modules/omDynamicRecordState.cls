VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omDynamicRecordState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private rs As ADODB.Recordset
Private timestampField As String
Public TimestampChanged As Boolean

Public rsOriginal As ADODB.Recordset
'Public rsInput As ADODB.Recordset
Public rsDatabase As ADODB.Recordset

Public rsChangedFields As ADODB.Recordset

Public Sub SetupFieldsFromRecordset(rsSrc As ADODB.Recordset, fieldNames As String)
Dim fieldNamesArray() As String
Dim fieldName As Variant
Dim fld As ADODB.Field
Dim i_max As Long
Dim i As Long

    If (rs Is Nothing) = False Then
        Terminate
        Initialize
    End If


    fieldNamesArray = Split(fieldNames, ",")
    For Each fieldName In fieldNamesArray
        Set fld = Nothing
        On Error Resume Next
        Set fld = rsSrc.Fields(fieldName)
        If (fld Is Nothing) = False Then
            rs.Fields.Append fld.Name, fld.Type, fld.DefinedSize
            If fld.Type = adDBTimeStamp Then
                timestampField = fieldName
            End If
        End If
    Next
    rs.Open
    rs.AddNew
    rs("omDynamicRecordState") = "O"
    rs.Update
    Set rsOriginal = rs.Clone
    rsOriginal.Find "omDynamicRecordState='O'", 0, adSearchForward, 1

    'rs.AddNew
    'rs("omDynamicRecordState") = "I"
    'rs.Update
    'Set rsInput = rs.Clone(adLockOptimistic)

    rs.AddNew
    rs("omDynamicRecordState") = "D"
    rs.Update
    Set rsDatabase = rs.Clone
    rsDatabase.Find "omDynamicRecordState='D'", 0, adSearchForward, 1
    If Not rsSrc.EOF Then
        SetOriginalRecord rsSrc
    End If
End Sub

Public Sub SetupFields()
    If (rs Is Nothing) = False Then
        Terminate
        Initialize
    End If
    rs.Fields.Append "ID_Auto", adInteger
    rs.Fields.Append "DAflevering", adDate
    rs.Fields.Append "SAflevering", adInteger
    rs.Fields.Append "upsize_ts", adDBTimeStamp
    timestampField = "upsize_ts"
    rs.Open
    rs.AddNew
    rs("omDynamicRecordState") = "O"
    rs.Update
    Set rsOriginal = rs.Clone
    rsOriginal.Find "omDynamicRecordState='O'", 0, adSearchForward, 1

    'rs.AddNew
    'rs("omDynamicRecordState") = "I"
    'rs.Update
    'Set rsInput = rs.Clone(adLockOptimistic)

    rs.AddNew
    rs("omDynamicRecordState") = "D"
    rs.Update
    Set rsDatabase = rs.Clone
    rsDatabase.Find "omDynamicRecordState='D'", 0, adSearchForward, 1
End Sub

Private Sub Class_Initialize()
    Initialize
End Sub
Public Sub SetOriginalRecord(rsSrc As ADODB.Recordset)
    rsOriginal.Update
    SetRecord rsOriginal, rsSrc
End Sub
'Public Sub SetInputRecord(rsSrc As ADODB.Recordset)
'    rsInput.Update
'    SetRecord rsInput, rsSrc
'End Sub
Public Sub SetDatabaseRecord(rsSrc As ADODB.Recordset)
    rsDatabase.Update
    SetRecord rsDatabase, rsSrc
End Sub

Private Sub SetRecord(rec As ADODB.Recordset, rsSrc As ADODB.Recordset)
Dim fld As ADODB.Field
    rec.Update
    For Each fld In rec.Fields
        On Error Resume Next
        rec(fld.Name) = rsSrc(fld.Name)
    Next
    rec.Update
End Sub
Public Function Compare()
Dim fld As ADODB.Field
    TimestampChanged = False
    If timestampField <> "" Then
        TimestampChanged = (rsOriginal(timestampField) = rsDatabase(timestampField))
        If TimestampChanged = False Then
            Exit Function
        End If
    End If
    For Each fld In rsOriginal.Fields
        If fld.Name <> "omDynamicRecordState" And fld.Name <> timestampField And Nz(rsOriginal(fld.Name), "0") <> Nz(rsDatabase(fld.Name), "0") Then
            TimestampChanged = True
            rsChangedFields.AddNew
            rsChangedFields("FieldName") = fld.Name
            rsChangedFields.Update
        End If
    Next
End Function
Private Sub Initialize()
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Fields.Append "omDynamicRecordState", adVarChar, 1

    Set rsChangedFields = New ADODB.Recordset
    rsChangedFields.CursorLocation = adUseClient
    rsChangedFields.Fields.Append "FieldName", adVarChar, 255
    rsChangedFields.Open
End Sub

Private Sub Class_Terminate()
    Terminate
End Sub

Private Sub Terminate()
    If (rsOriginal Is Nothing) = False Then
        Set rsOriginal = Nothing
    End If

'   If (rsInput  Is Nothing) = False Then
'       Set rsInput = Nothing
'   End If
    If (rsDatabase Is Nothing) = False Then
        Set rsDatabase = Nothing
    End If
    If (rs Is Nothing) = False Then
        Set rs = Nothing
    End If
    If (rsChangedFields Is Nothing) = False Then
        Set rsChangedFields = Nothing
    End If
End Sub
