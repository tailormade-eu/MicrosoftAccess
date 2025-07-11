﻿Attribute VB_Name = "omSSMAAConnector"
Option Compare Database
Option Explicit
Global cSQLNCLIVersion As Integer
Global cSQLNCLIFound As Boolean
Global cSQLODBCVersion As Integer
Global cSQLODBCFound As Boolean


'Global gomCS As New omConnectionString

Dim omCSTest As New omConnectionString
Dim omCSLastUsed As New omConnectionString
Dim m_ServerCon_ODBC As String
Dim m_ServerCon As String
Const cSSMAABAckup = "SSMAA_Backup_"

' Module Name   SSMAAConnector
'
' Author        Raoul Jacobs
' Company       Tailormade bv
' Email         jara@tailormade.eu
' Modify Date   2025-06-25
'
' Description
'   the function Link will relink all tables which are defined in the table SSMAA_ODBC_Tables
'   the function CreateTable will create the tabel SSMAA_ODBC_Tables

Public Function LocalCon() As ADODB.connection
  Set LocalCon = CurrentProject.connection
End Function

Public Function ServerCon(Optional ConnectionType As ConnectionTypes = ConnectionTypes.SQLNCLI, Optional ODBCConnection As Boolean = False, Optional encryptType As EncryptTypes = EncryptTypes.EncryptOptional) As String
  ' ODBCConnection = true => Direct connection to Database Server Using ODBC Provider
  ' ODBCConnection = False = Connection using MSAccess Objects OR ADODB Connection
  If ODBCConnection Then
    If IsNullOrEmpty(m_ServerCon_ODBC) Then
        m_ServerCon_ODBC = omSSMAAConnector.GetConnectionStringByProperty(ConnectionType:=ConnectionType, ODBCConnection:=ODBCConnection, encryptType:=encryptType)
    End If
    ServerCon = m_ServerCon_ODBC
  Else
    If IsNullOrEmpty(m_ServerCon) Then
        m_ServerCon = omSSMAAConnector.GetConnectionStringByProperty(ConnectionType:=ConnectionType, ODBCConnection:=ODBCConnection, encryptType:=encryptType)
    End If
    ServerCon = m_ServerCon
  End If
End Function


Public Function LinkUsingSSMA(Optional Group As String = "", Optional ConnectionType As ConnectionTypes = ConnectionTypes.SQLNCLI, Optional SavePassword As Boolean = False, Optional alwaysUpdate As Boolean = False, Optional encryptType As EncryptTypes = EncryptTypes.EncryptOptional) As Boolean
Dim rs As New ADODB.Recordset
Dim rsODBC As New ADODB.Recordset
Dim strSQLTable As String
Dim strDatabaseName As String

    'If IsNullOrEmptyOrZero(GetSQLNCLIVersion()) And ConnectionType = ConnectionTypes.SQLNCLI Then
    '    MsgBox "No SQL Server Native driver is installed"
    '    Exit Function
    'End If
    Debug.Print Now

    UpdateSSMAAGroups

    If alwaysUpdate Then
        'DeleteLinkTables Group
        LinkDeleteTables Group
    Else
        'DeleteLinkTables Group, True
        LinkDeleteTables Group, True
    End If

    UpdateSSMAConnectionString Group, ConnectionType, encryptType
    rs.Open "SELECT Name, Type FROM MSysObjects WHERE Type=1 OR Type=6 OR Type=4", LocalCon, adOpenStatic, adLockReadOnly
    rsODBC.Open "SELECT * FROM SSMAA_ODBC_Tables" & IIf(Group <> "", " WHERE [Groups] like " & Chr(34) & "%," & Group & ",%" & Chr(34), ""), LocalCon, adOpenForwardOnly, adLockOptimistic
    'rsODBC.Open "SELECT * FROM SSMAA_ODBC_Tables" & IIf(Group <> "", " WHERE [Group] like " & Chr(34) & "%," & Group & ",%" & Chr(34), ""), LocalCon, adOpenForwardOnly, adLockOptimistic
    While Not rsODBC.EOF
        'rs.Requery
        rs.filter = "Name='" & rsODBC("ODBCTable") & "'"
        If Not rs.EOF Then
            If rs("Type") = 1 Then
                DoCmd.Rename cSSMAABAckup & rsODBC("ODBCTable"), acTable, rsODBC("ODBCTable")
            ElseIf Not alwaysUpdate Then
                omCSTest.ParseByTableName rsODBC("ODBCTable")
                If Not (omCSTest.HasPasswordSaved Or omCSTest.HasTrustedConnection) Or (omCSTest.Server <> rsODBC("SQLServer")) Or (omCSTest.Database <> rsODBC("SQLDatabase")) Or (omCSTest.IsSQLNCLIConnection And omCSTest.SQLNCLIVersion <> GetSQLNCLIVersion) Then
                    DoCmd.DeleteObject acTable, rsODBC("ODBCTable")
                End If
            End If
        End If
        If Not omMSAccessFunctions.TableExists(rsODBC("ODBCTable")) Then
            strSQLTable = "[SQLTableOwner][SQLTable]"
            strSQLTable = Replace(strSQLTable, "[SQLTableOwner]", IIf(Len(rsODBC("SQLTableOwner")) > 0, rsODBC("SQLTableOwner") & ".", ""))
            strSQLTable = Replace(strSQLTable, "[SQLTable]", rsODBC("SQLTable"))

            'strDatabaseName = omCSLastUsed.GetConnectionString(rsODBC("SQLDatabase"), rsODBC("SQLServer"), True, ConnectionType, rsODBC("DSN"), rsODBC("SQLLogin"), rsODBC("SQLPassword"), Group)
            'strDatabaseName = "ODBC;" & rsODBC("ConnectionString")
            strDatabaseName = rsODBC("ConnectionString")
            'AttachedTable strDatabaseName, strSQLTable, rsODBC("ODBCTable"), SavePassword:=SavePassword
            If ConnectionType <> SQLOLEDBProvider Then
                strDatabaseName = "ODBC;" & strDatabaseName
            End If
            AttachTable strDatabaseName, strSQLTable, rsODBC("ODBCTable"), SavePassword:=SavePassword, primaryKey:=Nz(rsODBC("PrimaryKey"), "")
        End If
        'Debug.Print Now
        DoEvents
        rsODBC("Attach") = 0
        rsODBC.Update
        rsODBC.MoveNext
    Wend
    rsODBC.Close
    rs.Close
    Set rsODBC = Nothing
    Set rs = Nothing
    Debug.Print Now
End Function

Public Sub LinkTableViewUsingSSMA(existingTableName As String, toLinkName As String, Optional ConnectionType As ConnectionTypes = ConnectionTypes.SQLNCLI, Optional encryptType As EncryptTypes = EncryptTypes.EncryptOptional)
Dim cn As String
    cn = GetConnectionStringByProperty(existingTableName, ConnectionType:=ConnectionType, encryptType:=encryptType)
    AttachTable cn, toLinkName, toLinkName
End Sub

Sub UpdateSSMAConnectionString(Optional Group As String = "", Optional ConnectionType As ConnectionTypes = ConnectionTypes.SQLNCLI, Optional encryptType As EncryptTypes = EncryptTypes.EncryptOptional)
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command

    Set omCSLastUsed = New omConnectionString
    cmd.commandText = "UPDATE SSMAA_ODBC_Tables SET ConnectionString=? WHERE nz([Groups],'')=? AND nz(DSN,'')=? AND nz(SQLServer,'')=? AND nz(SQLServerPort,'')=? AND nz(SQLDatabase,'')=? AND nz(ConnectionTypeId,'')=? AND nz(SQLLogin,'')=? AND nz(SQLPassword,'')=?"
    'cmd.commandText = "UPDATE SSMAA_ODBC_Tables SET ConnectionString=? WHERE nz([Group],'')=? AND nz(DSN,'')=? AND nz(SQLServer,'')=? AND nz(SQLServerPort,'')=? AND nz(SQLDatabase,'')=? AND nz(ConnectionTypeId,'')=? AND nz(SQLLogin,'')=? AND nz(SQLPassword,'')=?"
    cmd.ActiveConnection = LocalCon
    cmd.Parameters.Append cmd.CreateParameter("p1", adVarWChar, adParamInput, 512)
    cmd.Parameters.Append cmd.CreateParameter("p2", adLongVarWChar, adParamInput, 256)
    cmd.Parameters.Append cmd.CreateParameter("p3", adLongVarWChar, adParamInput, 256)
    cmd.Parameters.Append cmd.CreateParameter("p4", adLongVarWChar, adParamInput, 256)
    cmd.Parameters.Append cmd.CreateParameter("p5", adLongVarWChar, adParamInput, 256)
    cmd.Parameters.Append cmd.CreateParameter("p6", adLongVarWChar, adParamInput, 256)
    cmd.Parameters.Append cmd.CreateParameter("p7", adLongVarWChar, adParamInput, 256)
    cmd.Parameters.Append cmd.CreateParameter("p8", adLongVarWChar, adParamInput, 256)
    cmd.Parameters.Append cmd.CreateParameter("p9", adLongVarWChar, adParamInput, 256)

    rs.Open "SELECT [Groups], DSN, SQLServer, SQLServerPort, SQLDatabase, ConnectionTypeId, SQLLogin, SQLPassword From SSMAA_ODBC_Tables " & IIf(Group <> "", " WHERE [Groups] like " & Chr(34) & "%," & Group & ",%" & Chr(34), "") & " GROUP BY [Groups], DSN, SQLServer, SQLServerPort, SQLDatabase, ConnectionTypeId, SQLLogin, SQLPassword", LocalCon, adOpenForwardOnly, adLockReadOnly
    'rs.Open "SELECT [Group], DSN, SQLServer, SQLServerPort, SQLDatabase, ConnectionTypeId, SQLLogin, SQLPassword From SSMAA_ODBC_Tables " & IIf(Group <> "", " WHERE [Group] like " & Chr(34) & "%," & Group & ",%" & Chr(34), "") & " GROUP BY [Group], DSN, SQLServer, SQLServerPort, SQLDatabase, ConnectionTypeId, SQLLogin, SQLPassword", LocalCon, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        cmd.Parameters(0) = omCSLastUsed.GetConnectionString(Nz(rs("sqlDatabase"), ""), Nz(rs("SQLServer"), ""), False, Nz(rs("ConnectionTypeId"), ConnectionType), Nz(rs("DSN"), ""), Nz(rs("SQLLogin"), ""), Nz(rs("SQLPassword"), ""), Nz(rs("Groups"), ""), Nz(rs("SQLServerPort"), ""), encryptType)
        'cmd.Parameters(0) = omCSLastUsed.GetConnectionString(Nz(rs("sqlDatabase"), ""), Nz(rs("SQLServer"), ""), False, Nz(rs("ConnectionTypeId"), ConnectionType), Nz(rs("DSN"), ""), Nz(rs("SQLLogin"), ""), Nz(rs("SQLPassword"), ""), Nz(rs("Group"), ""), Nz(rs("SQLServerPort"), ""))
        cmd.Parameters(1) = Nz(rs("Groups"), "")
        'cmd.Parameters(1) = Nz(rs("Group"), "")
        cmd.Parameters(2) = Nz(rs("DSN"), "")
        cmd.Parameters(3) = Nz(rs("SQLServer"), "")
        cmd.Parameters(4) = Nz(rs("SQLServerPort"), "")
        cmd.Parameters(5) = Nz(rs("SQLDatabase"), "")
        cmd.Parameters(6) = Nz(rs("ConnectionTypeId"), "")
        cmd.Parameters(7) = Nz(rs("SQLLogin"), "")
        cmd.Parameters(8) = Nz(rs("SQLPassword"), "")
        cmd.Execute
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
End Sub

'Sub AttachedTable(ConnectionString As String, sourceTable As String, DestinationTable As String, Optional SavePassword As Boolean = False)
Sub AttachTable(ConnectionString As String, SourceTable As String, DestinationTable As String, Optional SavePassword As Boolean = False, Optional primaryKey As String = "")
Dim tbl As dao.TableDef

    On Error Resume Next
    'Create a new TableDef object.
    Set tbl = CurrentDb.CreateTableDef(DestinationTable)
    'Set the properties to create the link
    tbl.Connect = ConnectionString
    tbl.SourceTableName = SourceTable

    If SavePassword And (tbl.Attributes And dbAttachSavePWD) = 0 Then
        tbl.Attributes = tbl.Attributes + dbAttachSavePWD
    End If
    'Add the new table to the database.
    CurrentDb.TableDefs.Append tbl
    Set tbl = Nothing

    If NotIsNullOrEmpty(primaryKey) Then
        CurrentDb.TableDefs.Refresh
        CurrentDb.Execute "CREATE UNIQUE INDEX PK_" & DestinationTable & " ON " & DestinationTable & " (" & primaryKey & ") WITH PRIMARY"
    End If
End Sub

Public Sub CreateTable()
Dim strSQL As String

    strSQL = "CREATE TABLE [SSMAA_ODBC_Tables]([SQLServer] TEXT(255),[SQLDatabase] TEXT(255),[SQLLogin] TEXT(100),[SQLPassword] TEXT(100),[SQLTableOwner] TEXT(100),[SQLTable] TEXT(255),[ODBCTable] TEXT(255),[ErrorMessage] TEXT(255)) "

    DoCmd.RunSQL strSQL

End Sub
Public Sub PopulateTable()
Dim strSQL As String

    strSQL = "INSERT INTO SSMAA_ODBC_Tables ( SQLTable, ODBCTable ) SELECT msysobjects.Name, msysobjects.Name FROM msysobjects LEFT JOIN SSMAA_ODBC_Tables ON msysobjects.Name = SSMAA_ODBC_Tables.SQLTable WHERE msysobjects.Type=6 AND SSMAA_ODBC_Tables.SQLTable Is Null"

    DoCmd.RunSQL strSQL
End Sub

Public Function GetSQLNCLIVersion() As String
Dim i As Long

    If cSQLNCLIFound Then
        GetSQLNCLIVersion = cSQLNCLIVersion
    Else
        i = 15
        While i > 9 And cSQLNCLIVersion = 0
            If gFso.FileExists("c:\windows\system32\sqlncli" & i & ".dll") Then
                cSQLNCLIVersion = i
                cSQLNCLIFound = True
            Else
                i = i - 1
            End If
        Wend
        GetSQLNCLIVersion = cSQLNCLIVersion
    End If
End Function
Public Function GetSQLODBCVersion() As String
Dim i As Long

    If cSQLODBCFound Then
        GetSQLODBCVersion = cSQLODBCVersion
    Else
        i = 25
        While i > 9 And cSQLODBCVersion = 0
            If gFso.FileExists("c:\windows\system32\msodbcsql" & i & ".dll") Then
                cSQLODBCVersion = i
                cSQLODBCFound = True
            Else
                i = i - 1
            End If
        Wend
        GetSQLODBCVersion = cSQLODBCVersion
    End If
End Function
Public Function GetConnectionStringByProperty(Optional tableName As String = "", Optional databaseName As String = "", Optional serverName As String = "", Optional dsnName As String, Optional GroupName As String, Optional ConnectionType As ConnectionTypes = ConnectionTypes.SQLNCLI, Optional ODBCConnection As Boolean = True, Optional encryptType As EncryptTypes = EncryptTypes.EncryptOptional) As String
Dim rs As New ADODB.Recordset
Dim strFilter As String

    rs.Open "SSMAA_ODBC_Tables", LocalCon, adOpenDynamic, adLockReadOnly
    If Not rs.EOF Then
        strFilter = strFilter & IIf(NotIsNullOrEmpty(tableName), " AND odbcTable='{0}'", "")
        strFilter = strFilter & IIf(NotIsNullOrEmpty(databaseName), " AND sqldatabase='{1}'", "")
        strFilter = strFilter & IIf(NotIsNullOrEmpty(serverName), " AND sqlserver='{2}'", "")
        strFilter = strFilter & IIf(NotIsNullOrEmpty(dsnName), " AND dsn='{3}'", "")
        'strFilter = strFilter & IIf(NotIsNullOrEmpty(groupName), " AND group like '%,{4},%'", "")
        strFilter = strFilter & IIf(NotIsNullOrEmpty(GroupName), " AND groups like '%,{4},%'", "")
        If NotIsNullOrEmpty(strFilter) Then
            rs.filter = StringFormat(Mid(strFilter, 6), tableName, databaseName, serverName, dsnName, GroupName)
            If rs.EOF Then
                rs.filter = ""
                rs.MoveFirst
            End If
        End If
        GetConnectionStringByProperty = Nz(rs("ConnectionString"), "")
        If IsNullOrEmpty(GetConnectionStringByProperty) Or Nz(rs("ConnectionTypeId"), 0) <> ConnectionType Then
            GetConnectionStringByProperty = omCSLastUsed.GetConnectionString(rs("SQLDatabase"), rs("SQLServer"), ODBCConnection, ConnectionType, rs("DSN"), rs("SQLLogin"), rs("SQLPassword"), encryptType:=encryptType)
        End If
        If ODBCConnection And Left(GetConnectionStringByProperty, 5) <> "ODBC;" Then
            GetConnectionStringByProperty = "ODBC;" & GetConnectionStringByProperty
        End If
    Else
        MsgBox "SSMAA Tables is empty!", vbExclamation
    End If
    rs.Close
    Set rs = Nothing
End Function
Public Function GetGroupByProperty(Optional tableName As String = "", Optional databaseName As String = "", Optional serverName As String = "", Optional dsnName As String) As String
Dim rs As New ADODB.Recordset
Dim strFilter As String

    rs.Open "SSMAA_ODBC_Tables", LocalCon, adOpenDynamic, adLockReadOnly
    If Not rs.EOF Then
        strFilter = strFilter & IIf(NotIsNullOrEmpty(tableName), " AND odbcTable='{0}'", "")
        strFilter = strFilter & IIf(NotIsNullOrEmpty(databaseName), " AND sqldatabase='{1}'", "")
        strFilter = strFilter & IIf(NotIsNullOrEmpty(serverName), " AND sqlserver='{2}'", "")
        strFilter = strFilter & IIf(NotIsNullOrEmpty(dsnName), " AND dsn='{3}'", "")
        If NotIsNullOrEmpty(strFilter) Then
            rs.filter = StringFormat(Mid(strFilter, 6), tableName, databaseName, serverName, dsnName)
            If Not rs.EOF Then
                'GetGroupByProperty = rs("Group")
                GetGroupByProperty = rs("Groups")
                If NotIsNullOrEmpty(GetGroupByProperty) Then
                    If Left(GetGroupByProperty, 1) = "," Then
                        GetGroupByProperty = Mid(GetGroupByProperty, 2)
                    End If
                    If Right(GetGroupByProperty, 1) = "," Then
                        GetGroupByProperty = Left(GetGroupByProperty, Len(GetGroupByProperty) - 1)
                    End If
                End If
            End If
        End If
    Else
        MsgBox "SSMAA Tables is empty!", vbExclamation
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Sub ImportTables(Optional storageTable As String = "SSMAA_ODBC_Tables", Optional overwrite As Boolean = False)
Dim rs As New ADODB.Recordset
Dim rsTable As New ADODB.Recordset
Dim typeFilter As String
Dim t As String

    rs.Open "SELECT * FROM MSysObjects", LocalCon, adOpenForwardOnly, adLockReadOnly
    rsTable.Open storageTable, CurrentProject.connection, adOpenDynamic, adLockOptimistic
    While Not rs.EOF

        typeFilter = "4" '",1,4,"
        If omStringFunctions.ContainsString(typeFilter, rs("Type"), ",") And rs("Name") <> storageTable Then
            Debug.Print rs("Name"), rs("Type")
            rsTable.filter = "[odbcTable]=" & "'" & rs("Name") & "'"
            'rsTable.Filter = "[odbcTable]=" & Chr(34) & rs("Name") & Chr(34)
            If rsTable.EOF Then
                rsTable.AddNew
            End If
            If rsTable.EditMode = adEditAdd Or overwrite Then
                rsTable("odbcTable") = rs("Name")
                rsTable("sqlTable") = rs("Name")
                If rs("Type") = 4 Then
                    omCSTest.ParseByTableName rs("Name")
                    rsTable("DSN") = omCSTest.DSN
                    rsTable("SQLServer") = omCSTest.Server
                    rsTable("SQLServerPort") = omCSTest.port
                    rsTable("SQLDatabase") = omCSTest.Database
                    rsTable("SQLLogin") = omCSTest.UID
                    rsTable("SQLPassword") = omCSTest.PWD
                    'rsTable("Group") = omCSTest.DSN
                    rsTable("Groups") = omCSTest.DSN
                    t = CurrentDb.TableDefs(rs("Name")).SourceTableName
                    If InStr(1, t, ".") = 0 Then
                        rsTable("SQLTableOwner") = ""
                        rsTable("SQLTable") = StringSplitGetByIndex(t, ".", 0)
                    Else
                        rsTable("SQLTableOwner") = StringSplitGetByIndex(t, ".", 0)
                        rsTable("SQLTable") = StringSplitGetByIndex(t, ".", 1)
                    End If
                End If
                rsTable.Update
            End If
        End If
        rs.MoveNext
    Wend
    rsTable.Close
    Set rsTable = Nothing
    rs.Close
    Set rs = Nothing
End Sub

Public Sub LinkPassthroughQueries(Optional ConnectionType As ConnectionTypes = ConnectionTypes.SQLNCLI, Optional encryptType As EncryptTypes = EncryptTypes.EncryptOptional)
Dim rs As New ADODB.Recordset
Dim t As String
Dim connString As String
Dim descriptionString As String

    rs.Open "SELECT * FROM MSysObjects WHERE Type=5 AND (flags=112 OR flags=144)", LocalCon, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        Debug.Print rs("Name"), rs("Type"), rs("flags")
        Debug.Print CurrentDb.QueryDefs(rs("Name")).Connect
        omCSTest.ParseByQueryName rs("Name")
        connString = ""
        t = ""
        t = omCSTest.Group
        If NotIsNullOrEmpty(t) Then
            connString = GetConnectionStringByProperty(GroupName:=t, ConnectionType:=ConnectionType, encryptType:=encryptType)
        Else
            t = omCSTest.DSN
            If NotIsNullOrEmpty(t) Then
                connString = GetConnectionStringByProperty(dsnName:=t, ConnectionType:=ConnectionType, encryptType:=encryptType)
                descriptionString = "Group=" & t
            Else
                t = omCSTest.Database
                If NotIsNullOrEmpty(t) Then
                    connString = GetConnectionStringByProperty(databaseName:=t, ConnectionType:=ConnectionType, encryptType:=encryptType)
                End If
            End If
        End If
        If NotIsNullOrEmpty(connString) Then
            CurrentDb.QueryDefs(rs("Name")).Connect = connString
            If IsNullOrEmpty(omCSTest.Group) Then
                omDAOFunctions.SetQueryDefProperty rs("Name"), "Description", descriptionString
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    CurrentDb.QueryDefs.Refresh
End Sub

Public Sub PassthroughQueriesReplaceDatabaseName(sourceDatabase As String, destinationDatabase As String, Optional schema As String = "dbo")
Dim rs As New ADODB.Recordset
Dim typeFilter As String
Dim t As String
Dim connString As String
Dim SQL As String

    Debug.Print Now
    rs.Open "SELECT * FROM MSysObjects WHERE Type=5 AND flags IN (112,144)", LocalCon, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        'Debug.Print rs("Name"), rs("Type"), rs("flags")
        'Debug.Print CurrentDb.QueryDefs(rs("Name")).Connect
        If InStr(1, CurrentDb.QueryDefs(rs("Name")).Connect, "database=" & sourceDatabase & ";") > 0 Then
                CurrentDb.QueryDefs(rs("Name")).Connect = Replace(CurrentDb.QueryDefs(rs("Name")).Connect, "database=" & sourceDatabase & ";", "database=" & destinationDatabase & ";")
        End If
        SQL = CurrentDb.QueryDefs(rs("Name")).SQL
        If InStr(1, SQL, "USE ") > 0 Then
            Debug.Print "USE => " & rs("Name"), rs("Type"), rs("flags")
            SQL = omSQLFunctions.ReplaceDatabaseInUseClause(SQL, sourceDatabase, destinationDatabase)
        End If
        If InStr(1, SQL, "." & schema & ".") > 0 Then
            Debug.Print "." & schema & "." & rs("Name"), rs("Type"), rs("flags")
            SQL = Replace(SQL, " " & sourceDatabase & "." & schema & ".", " " & destinationDatabase & "." & schema & ".")
            SQL = Replace(SQL, "," & sourceDatabase & "." & schema & ".", "," & destinationDatabase & "." & schema & ".")
            SQL = Replace(SQL, vbCrLf & sourceDatabase & "." & schema & ".", vbCrLf & destinationDatabase & "." & schema & ".")
            SQL = Replace(SQL, vbCr & sourceDatabase & "." & schema & ".", vbCr & destinationDatabase & "." & schema & ".")
            SQL = Replace(SQL, vbLf & sourceDatabase & "." & schema & ".", vbLf & destinationDatabase & "." & schema & ".")
        End If
        CurrentDb.QueryDefs(rs("Name")).SQL = SQL
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    CurrentDb.QueryDefs.Refresh
    Debug.Print Now
End Sub
Public Sub PassthroughQueriesUpdateDescriptionWithGroup()
Dim rs As New ADODB.Recordset
Dim typeFilter As String
Dim t As String
Dim connString As String
Dim descriptionString As String

    rs.Open "SELECT * FROM MSysObjects WHERE Type=5 AND flags IN (112,144)", LocalCon, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        Debug.Print rs("Name"), rs("Type"), rs("flags")
        Debug.Print CurrentDb.QueryDefs(rs("Name")).Connect
        omCSTest.ParseByQueryName rs("Name")
        t = omCSTest.Database
        If NotIsNullOrEmpty(t) And NotIsNullOrEmpty(GetGroupByProperty(databaseName:=t)) Then
            SetQueryDefProperty rs("Name"), "Description", "Group=" & GetGroupByProperty(databaseName:=t)
        End If
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    CurrentDb.QueryDefs.Refresh
End Sub

'Public Sub DeleteLinkTables(Optional Group As String = "", Optional AttachOnly As Boolean = False)
Public Sub LinkDeleteTables(Optional Group As String = "", Optional AttachOnly As Boolean = False)
Dim rsODBC As New ADODB.Recordset
Dim strWhere As String
Dim td As TableDef

    On Error Resume Next
    'Debug.Print Now
    strWhere = IIf(Group <> "", " WHERE ',' & [Groups] & ',' like '%," & Group & ",%'", "")
    strWhere = Replace(strWhere, "'", Chr(34))
    strWhere = omSQLFunctions.WhereAnd(strWhere, IIf(AttachOnly, "[Attach] = True", ""))
    rsODBC.Open omSQLFunctions.BuildSQL("SELECT * FROM SSMAA_ODBC_Tables", whereClause:=strWhere), LocalCon, adOpenForwardOnly
    Do Until rsODBC.EOF
        If CurrentDb.TableDefs(rsODBC("ODBCTable")).Connect <> "" Then
            CurrentDb.TableDefs.Delete rsODBC("ODBCTable")
        End If
        rsODBC.MoveNext
    Loop
    rsODBC.Close
    Set rsODBC = Nothing
    CurrentDb.TableDefs.Refresh
    'Debug.Print Now
End Sub


'Public Sub DeleteLinkPTs()
Public Sub LinkDeletePTs()
Dim rs As New ADODB.Recordset

  ' Does not work with LIKE "PT*"
  'rs.Open "SELECT Name, Type FROM MSysObjects AS O WHERE O.Type=5 AND O.[Name] LIKE " & SQLStringAcs("PT%"), LocalCon, adOpenStatic, adLockReadOnly

  rs.Open "SELECT Name, Type FROM MSysObjects WHERE Type=5 AND flags IN (112,144)", LocalCon, adOpenStatic, adLockReadOnly
  While Not rs.EOF
    DoCmd.DeleteObject acQuery, rs("Name")
    rs.MoveNext
  Wend
  rs.Close
  Set rs = Nothing

End Sub

Public Sub SetSQLConnectorConnectionString(tableName As String, Optional encryptType As EncryptTypes = EncryptTypes.EncryptOptional)
    gSQLConnector.ConnectionString = omSSMAAConnector.GetConnectionStringByProperty(tableName, ConnectionType:=SQLOLEDBProvider, encryptType:=encryptType)
End Sub


Public Sub DeleteSSMAABackupTables()
    omMSAccessFunctions.DeleteTables cSSMAABAckup
End Sub

Public Sub UpdateSSMAAGroups()
Dim rs As New ADODB.Recordset

    rs.Open "SELECT Groups FROM SSMAA_ODBC_Tables", LocalCon, adOpenForwardOnly, adLockOptimistic
    While Not rs.EOF
        If NotIsNullOrEmpty(rs("Groups")) Then
            rs("Groups") = Replace("," & rs("Groups") & ",", ",,", ",")
            rs.Update
        End If
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
End Sub

Public Function IsConnectingPossible(tableName As String, Optional connectionTimout As Long = 5, Optional encryptType As EncryptTypes = EncryptTypes.EncryptOptional) As Boolean
Dim lngStartTime As Long
Dim conn As New ADODB.connection
Dim cs As String

    cs = GetConnectionStringByProperty(tableName:=tableName, ODBCConnection:=False, encryptType:=encryptType)

    Set conn = New ADODB.connection
    conn.ConnectionString = cs
    conn.ConnectionTimeout = connectionTimout
    conn.Open Options:=adAsyncConnect

    lngStartTime = omKernalFunctions.GetTickCount()

    Do While ((omKernalFunctions.GetTickCount() - lngStartTime) < conn.ConnectionTimeout * 1000) And (Not conn.state = adStateOpen)
    Loop

    If conn.state = adStateOpen Then
        IsConnectingPossible = True
        conn.Close
    End If

    Set conn = Nothing
End Function

Public Sub DeleteSSMAATables(Optional Group As String = "")
Dim rsODBC As New ADODB.Recordset
Dim strWhere As String

    On Error Resume Next
    'Debug.Print Now
    strWhere = IIf(Group <> "", " WHERE [Groups] like " & Chr(34) & "%," & Group & ",%" & Chr(34), "")
    rsODBC.Open omSQLFunctions.BuildSQL("SELECT * FROM SSMAA_ODBC_Tables", whereClause:=strWhere), LocalCon, adOpenForwardOnly
    Do Until rsODBC.EOF
        If TableExists(rsODBC("ODBCTable")) Then
            DoCmd.DeleteObject acTable, rsODBC("ODBCTable")
        End If
        rsODBC.MoveNext
    Loop
    rsODBC.Close
    Set rsODBC = Nothing
    'Debug.Print Now
End Sub

Public Sub LinkMSAccess(Optional linkLocal As Boolean = False)
Dim tblCon As New omTableConnector
Dim DefaultPath As String
Dim dataPath As String
Dim filename As String


    filename = omFileFunctions.RemoveExtension(CurrentProject.Name)
    If InStrRev(filename, "_client") > 0 Then
        filename = Replace(filename, "_client", "")
    End If
    filename = filename & "_Data"
    gSystemDefaults.Mode = LocalMode
    gSystemDefaults.Development = gDevelopmentMode
    If linkLocal Then
        dataPath = CurrentProject.path
    Else
        dataPath = Nz(gSystemDefaults.Load("DataPath"), "")
    End If

    DefaultPath = omFileFunctions.BuildPathFileExists(dataPath, filename & ".mdb")
    If DefaultPath = "" Then
        DefaultPath = omFileFunctions.BuildPathFileExists(dataPath, filename & ".accdb")
    End If
    If DefaultPath = "" Then
        DefaultPath = omFileFunctions.BuildPathFileExists(CurrentProject.path, filename & ".mdb")
    End If
    If DefaultPath = "" Then
        DefaultPath = omFileFunctions.BuildPathFileExists(CurrentProject.path, filename & ".accdb")
    End If
    If Not gFso.FileExists(DefaultPath) Then
        MsgBox "No Data file was found at the location: " & DefaultPath & vbCrLf & "Application will now be closed.", vbCritical
        DoCmd.Quit acQuitSaveNone
    End If
    tblCon.DataFilename = DefaultPath
    tblCon.Connect omTableConnectionType.DatafileIsSource

End Sub
