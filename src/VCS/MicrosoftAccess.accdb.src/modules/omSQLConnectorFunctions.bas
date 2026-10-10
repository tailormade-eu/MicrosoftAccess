Attribute VB_Name = "omSQLConnectorFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=17c63d6d

Global gSQLConnector As New omSQLConnector


' Lives here, next to gSQLConnector, so omSSMAAConnector compiles in a database that has no gSQLConnector (e.g. a client project uses CurrentSQL).
Public Sub SetSQLConnectorConnectionString(tableName As String, Optional encryptType As EncryptTypes = EncryptTypes.EncryptOptional)
    gSQLConnector.ConnectionString = omSSMAAConnector.GetConnectionStringByProperty(tableName, ConnectionType:=SQLOLEDBProvider, encryptType:=encryptType)
End Sub


' Tests whether gSQLConnector can reach a server, with a statement that changes nothing (SELECT 1).
' The connection comes from explicitConnectionString when given, otherwise from the linked table linkedTableName
' (its server and login, the same way SetSQLConnectorConnectionString does it). Without either, the connection
' that gSQLConnector already has (or opens by itself) is tested.
' Returns "" when it works, otherwise the text of the error.
' Example: Debug.Print ConnectorTest("T_Auto", ConnectionTypes.SQLODBC, EncryptTypes.EncryptMandatory)
Public Function ConnectorTest(Optional linkedTableName As String = "", _
        Optional ConnectionType As ConnectionTypes = ConnectionTypes.SQLODBC, _
        Optional encryptType As EncryptTypes = EncryptTypes.EncryptOptional, _
        Optional explicitConnectionString As String = "") As String
    On Error GoTo ConnectorTest_Error
    If Len(explicitConnectionString) > 0 Then
        gSQLConnector.ConnectionString = explicitConnectionString
    ElseIf Len(linkedTableName) > 0 Then
        gSQLConnector.ConnectionString = omSSMAAConnector.GetConnectionStringByProperty(linkedTableName, ConnectionType:=ConnectionType, encryptType:=encryptType)
    End If
    gSQLConnector.ExecuteCommandTextNonQuery "SELECT 1"
    Exit Function
ConnectorTest_Error:
    ConnectorTest = Err.Description
End Function
