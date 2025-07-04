Attribute VB_Name = "omDAOFunctions"
Option Compare Database
Option Explicit

Public Sub SetQueryDefProperty(queryName As String, propertyName As String, Value As String, Optional propertyType As dao.DataTypeEnum = dbText)
Dim prp As dao.Property
    On Error Resume Next
    Set prp = CurrentDb.QueryDefs(queryName).CreateProperty(propertyName, propertyType, Value)
    CurrentDb.QueryDefs(queryName).Properties.Append prp
    If Err = 3367 Then
        CurrentDb.QueryDefs(queryName).Properties(propertyName).Value = Value
    End If
    On Error GoTo 0
End Sub
