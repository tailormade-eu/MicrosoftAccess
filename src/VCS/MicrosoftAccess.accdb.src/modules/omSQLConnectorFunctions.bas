Attribute VB_Name = "omSQLConnectorFunctions"
Option Compare Database
Option Explicit

Global gSQLConnector As New omSQLConnector


' Lives here, next to gSQLConnector, so omSSMAAConnector compiles in a database that has no gSQLConnector (e.g. a client project uses CurrentSQL).
Public Sub SetSQLConnectorConnectionString(tableName As String, Optional encryptType As EncryptTypes = EncryptTypes.EncryptOptional)
    gSQLConnector.ConnectionString = omSSMAAConnector.GetConnectionStringByProperty(tableName, ConnectionType:=SQLOLEDBProvider, encryptType:=encryptType)
End Sub


Public Sub ConnectorTest()
    'gSQLConnector.ExecuteCommandTextNonQuery "update External_T_Auto set id_uitrsuting=?, ManualMatched=?,HideLine=? where TextClean=?", ID_Uitrsuting, ManualMatched, HideLine, TextClean
    ' Als tabel gelinked is kan de gewone access connectie gebruikt worden : dit is default
    gSQLConnector.ExecuteCommandTextNonQuery "UPDATE t_klantenquete SET ID_Taal=? WHERE ID_Enquete=?", 1, 1

    gSQLConnector.ConnectionString = omSSMAAConnector.GetConnectionStringByProperty("T_Auto", ConnectionType:=ConnectionTypes.SQLODBC, encryptType:=EncryptMandatory)
    gSQLConnector.ExecuteCommandTextNonQuery "UPDATE t_klantenquete SET ID_Taal=? WHERE ID_Enquete=?", 1, 2
End Sub
