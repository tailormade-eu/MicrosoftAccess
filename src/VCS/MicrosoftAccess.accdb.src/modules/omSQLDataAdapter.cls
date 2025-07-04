VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omSQLDataAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public connection As New ADODB.connection

Public Sub spExecuteNonQuery(storedProcedureName As String, Optional Parameter1 As Variant = Null, Optional Parameter2 As Variant = Null, Optional Parameter3 As Variant = Null, Optional Parameter4 As Variant = Null, Optional parameter5 As Variant = Null, Optional parameter6 As Variant = Null)
Dim cmd As New ADODB.Command
Dim i As Long

    cmd.CommandType = adCmdStoredProc
    cmd.commandText = storedProcedureName
    If connection.state = 0 Then
        connection.Open
    End If
    Set cmd.ActiveConnection = connection
    cmd.Parameters.Refresh
    ' Load Parameters
    For i = 1 To cmd.Parameters.Count - 1
        Select Case i
            Case 1
                cmd.Parameters(i) = Parameter1
            Case 2
                cmd.Parameters(i) = Parameter2
            Case 3
                cmd.Parameters(i) = Parameter3
            Case 4
                cmd.Parameters(i) = Parameter4
            Case 5
                cmd.Parameters(i) = parameter5
            Case 6
                cmd.Parameters(i) = parameter6
        End Select
    Next
    cmd.Execute
    connection.Close
    Set cmd = Nothing
End Sub
Public Function spExecuteScalar(storedProcedureName As String, Optional Parameter1 As Variant = Null, Optional Parameter2 As Variant = Null, Optional Parameter3 As Variant = Null, Optional Parameter4 As Variant = Null, Optional parameter5 As Variant = Null, Optional parameter6 As Variant = Null, Optional defaultValue As Variant = Null) As Variant
Dim cmd As New ADODB.Command
Dim rs As ADODB.Recordset
Dim i As Long

    cmd.CommandType = adCmdStoredProc
    cmd.commandText = storedProcedureName
    If connection.state = 0 Then
        connection.Open
    End If
    Set cmd.ActiveConnection = connection
    cmd.Parameters.Refresh
    ' Load Parameters
    For i = 1 To cmd.Parameters.Count - 1
        Select Case i
            Case 1
                cmd.Parameters(i) = Parameter1
            Case 2
                cmd.Parameters(i) = Parameter2
            Case 3
                cmd.Parameters(i) = Parameter3
            Case 4
                cmd.Parameters(i) = Parameter4
            Case 5
                cmd.Parameters(i) = parameter5
            Case 6
                cmd.Parameters(i) = parameter6
        End Select
    Next
    Set rs = cmd.Execute
    If Not rs.EOF Then
        spExecuteScalar = Nz(rs(0), defaultValue)
    Else
        spExecuteScalar = defaultValue
    End If
    rs.Close
    Set rs = Nothing
    connection.Close
    Set cmd = Nothing
End Function
Public Function sqlOpenRecordset(SQLStatement As String) As ADODB.Recordset
Dim cmd As New ADODB.Command

    cmd.CommandType = adCmdText
    cmd.commandText = SQLStatement
    If connection.state = 0 Then
        connection.Open
    End If
    Set cmd.ActiveConnection = connection
    Set sqlOpenRecordset = cmd.Execute
    Set cmd = Nothing
End Function
Public Sub sqlExecute(SQLStatement As String)
Dim cmd As New ADODB.Command

    cmd.CommandType = adCmdText
    cmd.commandText = SQLStatement
    If connection.state = 0 Then
        connection.Open
    End If
    Set cmd.ActiveConnection = connection
    cmd.Execute
    connection.Close
    Set cmd = Nothing
End Sub
Public Sub CloseConnection()
    connection.Close
End Sub
Private Sub Class_Initialize()
Dim strDatabaseName As String

    strDatabaseName = "Provider=MSDataShape.1;Data Provider=SQLOLEDB.1;Persist Security Info=True;SERVER=[SQLServer];DATABASE=[SQLDatabase];"
    strDatabaseName = Replace(strDatabaseName, "[SQLServer]", "vamib-srv\vamos")
    strDatabaseName = Replace(strDatabaseName, "[SQLDatabase]", "Vamos")
    If Len("jara") > 0 Then
        strDatabaseName = strDatabaseName & "UID=[SQLLogin];PWD=[SQLPassword];"
        strDatabaseName = Replace(strDatabaseName, "[SQLLogin]", "vamos")
        strDatabaseName = Replace(strDatabaseName, "[SQLPassword]", "79va20")
    Else
        strDatabaseName = strDatabaseName & "Trusted_Connection=yes;"
    End If
    connection.ConnectionString = strDatabaseName
    connection.ConnectionTimeout = 0
End Sub

Private Sub Class_Terminate()
    Set Me.connection = Nothing
End Sub
