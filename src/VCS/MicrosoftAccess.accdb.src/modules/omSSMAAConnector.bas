Attribute VB_Name = "omSSMAAConnector"
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
' Een database-object voor de hele koppelronde (LinkUsingSSMA). Elke CurrentDb maakt een nieuw
' object en leest de tabellenlijst opnieuw in; per tabel meermaals = bijna kwadratisch bij
' ~1000 tabellen. Idee overgenomen van client (omSSMAAConnector, Dirk 16/09/2026: 50 s -> 5 s).
Dim mDbKoppelen As dao.Database
' ==== BEGIN auto driver declarations ====
Private Const AUTO_DRIVER_ODBC_MAX As Long = 25      ' highest "ODBC Driver n for SQL Server" that is looked for
Private Const AUTO_DRIVER_ODBC_MIN As Long = 11
Private Const AUTO_DRIVER_PROBE_TIMEOUT As Long = 5  ' seconds for one probe connection

Private Enum SqlDriverProbeResult
    ProbeOk = 0
    ProbeDriverFailed = 1       ' the driver itself cannot be used: try the next candidate
    ProbeOtherFailure = 2       ' server unreachable, login refused, ...: another driver would not help
End Enum

Private m_AutoDrivers As Object  ' Scripting.Dictionary: "server|port|database|login" -> chosen driver name
Private m_SchemaOk As Boolean    ' EnsureSSMAASchema already ran in this session
' ==== END auto driver declarations ====


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

Private Function KoppelDb() As dao.Database
    If mDbKoppelen Is Nothing Then
        Set KoppelDb = CurrentDb
    Else
        Set KoppelDb = mDbKoppelen
    End If
End Function

Private Function KoppelTabelBestaat(tableName As String) As Boolean
    On Error Resume Next
    KoppelTabelBestaat = (KoppelDb.TableDefs(tableName).Name = tableName)
End Function

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
Dim sngStart As Single
Dim lngAantal As Long

    On Error GoTo Opruimen
    Set mDbKoppelen = CurrentDb
    sngStart = Timer
    ResetAutoDriver
    EnsureSSMAASchema True

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
                mDbKoppelen.TableDefs(rsODBC("ODBCTable")).Name = cSSMAABAckup & rsODBC("ODBCTable")
            ElseIf Not alwaysUpdate Then
                omCSTest.ParseByTableName rsODBC("ODBCTable")
                If Not (omCSTest.HasPasswordSaved Or omCSTest.HasTrustedConnection) Or (omCSTest.Server <> rsODBC("SQLServer")) Or (omCSTest.Database <> rsODBC("SQLDatabase")) Or (omCSTest.IsSQLNCLIConnection And omCSTest.SQLNCLIVersion <> GetSQLNCLIVersion) _
                        Or LinkedConnectionDiffers(omCSTest.Driver, omCSTest.port, omCSTest.UID, omCSTest.HasPasswordSaved, Nz(rsODBC("ConnectionString"), "")) _
                        Or Nz(rsODBC("Attach"), False) Then
                    mDbKoppelen.TableDefs.Delete rsODBC("ODBCTable")
                End If
            End If
        End If
        If Not KoppelTabelBestaat(rsODBC("ODBCTable")) Then
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
            lngAantal = lngAantal + 1
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
    Debug.Print Now; " LinkUsingSSMA: "; lngAantal; " tabellen gekoppeld in "; Format(Timer - sngStart, "0.0"); " s"
Opruimen:
    Set mDbKoppelen = Nothing
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
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
Dim db As dao.Database

    On Error Resume Next
    Set db = KoppelDb
    'Create a new TableDef object.
    Set tbl = db.CreateTableDef(DestinationTable)
    'Set the properties to create the link
    tbl.Connect = ConnectionString
    tbl.SourceTableName = SourceTable

    If SavePassword And (tbl.Attributes And dbAttachSavePWD) = 0 Then
        tbl.Attributes = tbl.Attributes + dbAttachSavePWD
    End If
    'Add the new table to the database.
    db.TableDefs.Append tbl
    Set tbl = Nothing

    If NotIsNullOrEmpty(primaryKey) Then
        db.TableDefs.Refresh
        db.Execute "CREATE UNIQUE INDEX PK_" & DestinationTable & " ON " & DestinationTable & " (" & primaryKey & ") WITH PRIMARY"
    End If
End Sub

Public Sub CreateTable()
    EnsureSSMAASchema True
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

    EnsureSSMAASchema
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

    EnsureSSMAASchema
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

    EnsureSSMAASchema
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

    EnsureSSMAASchema
    On Error Resume Next
    'Debug.Print Now
    strWhere = IIf(Group <> "", " WHERE ',' & [Groups] & ',' like '%," & Group & ",%'", "")
    strWhere = Replace(strWhere, "'", Chr(34))
    strWhere = omSQLFunctions.WhereAnd(strWhere, IIf(AttachOnly, "[Attach] = True", ""))
    rsODBC.Open omSQLFunctions.BuildSQL("SELECT * FROM SSMAA_ODBC_Tables", whereClause:=strWhere), LocalCon, adOpenForwardOnly
    Do Until rsODBC.EOF
        If KoppelDb.TableDefs(rsODBC("ODBCTable")).Connect <> "" Then
            KoppelDb.TableDefs.Delete rsODBC("ODBCTable")
        End If
        rsODBC.MoveNext
    Loop
    rsODBC.Close
    Set rsODBC = Nothing
    KoppelDb.TableDefs.Refresh
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

Public Sub DeleteSSMAABackupTables()
    omMSAccessFunctions.DeleteTables cSSMAABAckup
End Sub

Public Sub UpdateSSMAAGroups()
Dim rs As New ADODB.Recordset

    EnsureSSMAASchema
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

    lngStartTime = omKernalAPIFunctions.GetTickCount()

    Do While ((omKernalAPIFunctions.GetTickCount() - lngStartTime) < conn.ConnectionTimeout * 1000) And (Not conn.state = adStateOpen)
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

' ==== BEGIN auto driver functions ====
' SQLAuto = newest SQL Server ODBC driver that is installed on this PC, then older ones, then SQL Server
' Native Client, then the legacy Windows driver "SQL Server".
' IsOdbcDriverInstalled and GetDriverFromConnectionString come from client (omSSMAAConnector, 28/09/2026),
' translated to English. The registry check replaces the file check on c:\windows\system32.

Public Function IsOdbcDriverInstalled(ByVal driverName As String) As Boolean
    ' Registry first: read-only RegRead (BUILTIN\Users have ReadKey on the ODBC Drivers key). 32-bit Office reads the
    ' 32-bit view, so this lists exactly the drivers Access can load. If the registry cannot be read at all
    ' ("SQL Server" is registered on every Windows) the DLL on disk is checked instead.
    Dim wsh As Object
    Dim sentinel As Variant
    Dim registryReadable As Boolean

    If Len(driverName) = 0 Then Exit Function
    If StrComp(driverName, "SQL Server", vbTextCompare) = 0 Then
        IsOdbcDriverInstalled = True
        Exit Function
    End If

    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    sentinel = wsh.RegRead("HKLM\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers\SQL Server")
    registryReadable = (Err.Number = 0)
    Err.Clear
    If registryReadable Then
        IsOdbcDriverInstalled = (wsh.RegRead("HKLM\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers\" & driverName) = "Installed")
        Err.Clear
    Else
        IsOdbcDriverInstalled = DriverDllExists(driverName)
    End If
    Set wsh = Nothing
End Function

Private Function DriverDllExists(ByVal driverName As String) As Boolean
    Dim dllName As String
    Dim major As Long

    major = GetOdbcDriverMajor(driverName)
    If major > 0 Then
        dllName = "msodbcsql" & major & ".dll"
    ElseIf StrComp(driverName, "SQL Server Native Client 11.0", vbTextCompare) = 0 Then
        dllName = "sqlncli11.dll"
    ElseIf StrComp(driverName, "SQL Server Native Client 10.0", vbTextCompare) = 0 Then
        dllName = "sqlncli10.dll"
    Else
        Exit Function
    End If
    DriverDllExists = (Len(Dir(Environ$("SystemRoot") & "\System32\" & dllName)) > 0)
End Function

Public Function GetOdbcDriverMajor(ByVal driverName As String) As Long
    ' "ODBC Driver 18 for SQL Server" -> 18; every other name -> 0
    Dim s As String

    If StrComp(Left$(driverName, 12), "ODBC Driver ", vbTextCompare) <> 0 Then Exit Function
    If StrComp(Right$(driverName, 15), " for SQL Server", vbTextCompare) <> 0 Then Exit Function
    If Len(driverName) <= 27 Then Exit Function
    s = Mid$(driverName, 13, Len(driverName) - 27)
    If IsNumeric(s) Then GetOdbcDriverMajor = CLng(Val(s))
End Function

Public Function GetConnectionStringPart(ByVal connectionString As String, ByVal partName As String) As String
    ' ("ODBC;DRIVER={x};SERVER=s;UID=u;", "UID") -> "u"; braces removed; "" when the part is missing
    Dim parts() As String
    Dim p As Variant
    Dim prefix As String

    prefix = partName & "="
    parts = Split(connectionString, ";")
    For Each p In parts
        If StrComp(Left$(Trim$(p), Len(prefix)), prefix, vbTextCompare) = 0 Then
            GetConnectionStringPart = Replace(Replace(Mid$(Trim$(p), Len(prefix) + 1), "{", ""), "}", "")
            Exit Function
        End If
    Next p
End Function

Public Function GetDriverFromConnectionString(ByVal connectionString As String) As String
    ' "ODBC;DRIVER={ODBC Driver 18 for SQL Server};SERVER=..." -> "ODBC Driver 18 for SQL Server"; no DRIVER= -> ""
    GetDriverFromConnectionString = GetConnectionStringPart(connectionString, "DRIVER")
End Function

Private Function LinkedConnectionDiffers(ByVal linkedDriver As String, ByVal linkedPort As String, ByVal linkedUid As String, ByVal linkedHasPassword As Boolean, ByVal newConnectionString As String) As Boolean
    ' Compares what is linked now (parsed from the table's Connect property by the caller) with the string LinkUsingSSMA
    ' is about to use. Only parts Access keeps unchanged are compared: driver, port and, when the password is saved, the login.
    ' Server, database and password/trusted are already checked by the caller. DSN or OLE DB strings have no DRIVER= and are skipped.
    Dim wantedDriver As String
    Dim wantedServer As String
    Dim wantedPort As String
    Dim wantedUid As String

    If Len(newConnectionString) = 0 Then Exit Function
    wantedDriver = GetConnectionStringPart(newConnectionString, "DRIVER")
    wantedServer = GetConnectionStringPart(newConnectionString, "SERVER")
    If InStr(1, wantedServer, ",") > 0 Then
        wantedPort = Mid$(wantedServer, InStr(1, wantedServer, ",") + 1)
    Else
        wantedPort = GetConnectionStringPart(newConnectionString, "PORT")   ' MySQL only
    End If
    wantedUid = GetConnectionStringPart(newConnectionString, "UID")

    If Len(linkedDriver) > 0 And Len(wantedDriver) > 0 Then
        If StrComp(Replace(Replace(linkedDriver, "{", ""), "}", ""), wantedDriver, vbTextCompare) <> 0 Then LinkedConnectionDiffers = True: Exit Function
    End If
    If Len(wantedDriver) > 0 Then   ' only compare ports for driver strings (a DSN has no port of its own)
        If StrComp(linkedPort, wantedPort, vbTextCompare) <> 0 Then LinkedConnectionDiffers = True: Exit Function
    End If
    If linkedHasPassword And Len(wantedUid) > 0 Then
        If StrComp(linkedUid, wantedUid, vbTextCompare) <> 0 Then LinkedConnectionDiffers = True
    End If
End Function

Public Function EncryptTextForType(ByVal encryptType As EncryptTypes) As String
    ' The Encrypt keywords of ODBC Driver 18 and newer (the Provider=SQLOLEDB / ODBC 18 syntax): optional, yes + trust, strict.
    Select Case encryptType
        Case EncryptTypes.EncryptMandatory
            EncryptTextForType = "Encrypt=yes;TrustServerCertificate=yes;"
        Case EncryptTypes.EncryptStrict
            EncryptTextForType = "Encrypt=strict;"
        Case Else   ' optional and any unexpected value
            EncryptTextForType = "Encrypt=optional;"
    End Select
End Function

Public Function EncryptTextForDriver(ByVal driverName As String, ByVal encryptType As EncryptTypes) As String
    ' What this driver understands (measured on this pc, 2026-10-07):
    '  ODBC Driver 18+  optional / yes / strict.
    '  ODBC Driver 17-  ADO answers "Invalid value specified" to optional and strict (and an Access link hung on a dialog); only yes/no.
    '                   No keyword = does not require encryption = optional. Mandatory = yes + trust.
    '  Native Client    same as 17: yes/no only (not installed here, from the documentation, untested).
    '  "SQL Server"     no Encrypt keyword at all: it accepts optional/strict without complaint and then does NOT encrypt,
    '                   so GetAutoDriver never picks it for mandatory or strict (DriverSupportsEncrypt).
    If GetOdbcDriverMajor(driverName) >= 18 Then
        EncryptTextForDriver = EncryptTextForType(encryptType)
    ElseIf GetOdbcDriverMajor(driverName) > 0 Or InStr(1, driverName, "SQL Server Native Client", vbTextCompare) = 1 Then
        If encryptType = EncryptTypes.EncryptMandatory Then EncryptTextForDriver = "Encrypt=yes;TrustServerCertificate=yes;"
    End If
End Function

Public Function DriverSupportsEncrypt(ByVal driverName As String, ByVal encryptType As EncryptTypes) As Boolean
    ' Optional: every driver (one that knows no Encrypt keyword simply does not encrypt unless the server forces it).
    ' Mandatory: needs a driver that understands Encrypt=yes: any "ODBC Driver n" and SQL Server Native Client, not the legacy "SQL Server" driver.
    ' Strict (TDS 8): ODBC Driver 18 and newer only.
    ' Measured: legacy "SQL Server" fails on Encrypt=yes and silently ignores strict; ODBC 17 does not know strict. Native Client: from the docs.
    Select Case encryptType
        Case EncryptTypes.EncryptStrict
            DriverSupportsEncrypt = (GetOdbcDriverMajor(driverName) >= 18)
        Case EncryptTypes.EncryptMandatory
            DriverSupportsEncrypt = (GetOdbcDriverMajor(driverName) > 0) Or (InStr(1, driverName, "SQL Server Native Client", vbTextCompare) = 1)
        Case Else
            DriverSupportsEncrypt = True
    End Select
End Function

Private Function AutoDriverCandidate(ByVal index As Long) As String
    ' index 0.. = "ODBC Driver 25 for SQL Server" down to 11, then Native Client 11.0 and 10.0, then the legacy driver; "" = end of list
    Dim odbcCount As Long

    odbcCount = AUTO_DRIVER_ODBC_MAX - AUTO_DRIVER_ODBC_MIN + 1
    If index < 0 Then
        Exit Function
    ElseIf index < odbcCount Then
        AutoDriverCandidate = "ODBC Driver " & (AUTO_DRIVER_ODBC_MAX - index) & " for SQL Server"
    ElseIf index = odbcCount Then
        AutoDriverCandidate = "SQL Server Native Client 11.0"
    ElseIf index = odbcCount + 1 Then
        AutoDriverCandidate = "SQL Server Native Client 10.0"
    ElseIf index = odbcCount + 2 Then
        AutoDriverCandidate = "SQL Server"
    End If
End Function

Public Function ProbeSqlDriver(ByVal driverName As String, ByVal serverName As String, ByVal databaseName As String, ByVal userId As String, ByVal password As String, ByVal port As String, ByVal encryptType As EncryptTypes) As Long
    ' Opens one real connection with this driver. Returns a SqlDriverProbeResult as Long (a Public function cannot
    ' expose a Private Enum): 0 = works, 1 = the driver cannot be used, 2 = failed for another reason (server, login, network).
    Dim cn As Object
    Dim cs As String
    Dim msg As String

    cs = "DRIVER={" & driverName & "};SERVER=" & serverName & IIf(Len(port) > 0, "," & port, "") & ";DATABASE=" & databaseName & ";"
    cs = cs & EncryptTextForDriver(driverName, encryptType)   ' exactly the Encrypt text the real connection string will carry
    If Len(userId) > 0 Then
        cs = cs & "UID=" & userId & ";PWD=" & password & ";"
    Else
        cs = cs & "Trusted_Connection=yes;"
    End If

    On Error GoTo ProbeFailed
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = AUTO_DRIVER_PROBE_TIMEOUT
    cn.Open cs
    cn.Close
    ProbeSqlDriver = ProbeOk
    Exit Function

ProbeFailed:
    msg = LCase$(Err.Description)
    ' Driver problems: the driver is not registered/loadable, rejects a keyword, or cannot negotiate the TLS version of an
    ' old server (a newer driver needs TLS 1.2; an older driver or Native Client may still connect). From memory, untested.
    If InStr(msg, "data source name not found") > 0 _
       Or InStr(msg, "invalid connection string attribute") > 0 _
       Or InStr(msg, "invalid value specified") > 0 _
       Or InStr(msg, "specified driver could not be loaded") > 0 _
       Or InStr(msg, "ssl") > 0 _
       Or InStr(msg, "tls") > 0 Then
        ProbeSqlDriver = ProbeDriverFailed
    Else
        ProbeSqlDriver = ProbeOtherFailure
    End If
End Function

Public Function GetAutoDriver(ByVal serverName As String, ByVal databaseName As String, ByVal userId As String, ByVal password As String, ByVal port As String, ByVal encryptType As EncryptTypes) As String
    ' Walks the candidates from newest to oldest. A candidate counts when it is installed and its probe connection does
    ' not fail because of the driver itself. An unreachable server is not a driver problem, so it does not trigger N probes:
    ' the first installed driver is kept. Drivers that cannot do the requested encryption are skipped (see DriverSupportsEncrypt);
    ' when no installed driver can, the result is "" and the caller must refuse, not silently connect unencrypted.
    ' The choice is cached per server/port/database/login/encrypt type until ResetAutoDriver.
    Dim key As String
    Dim i As Long
    Dim candidate As String
    Dim chosen As String

    If m_AutoDrivers Is Nothing Then
        Set m_AutoDrivers = CreateObject("Scripting.Dictionary")
        m_AutoDrivers.CompareMode = 1   ' text compare
    End If
    key = serverName & "|" & port & "|" & databaseName & "|" & userId & "|" & encryptType
    If m_AutoDrivers.Exists(key) Then
        GetAutoDriver = m_AutoDrivers(key)
        Exit Function
    End If

    Do
        candidate = AutoDriverCandidate(i)
        If Len(candidate) = 0 Then Exit Do
        If IsOdbcDriverInstalled(candidate) And DriverSupportsEncrypt(candidate, encryptType) Then
            If StrComp(candidate, "SQL Server", vbTextCompare) = 0 Then
                chosen = candidate
                Exit Do
            End If
            If ProbeSqlDriver(candidate, serverName, databaseName, userId, password, port, encryptType) <> ProbeDriverFailed Then
                chosen = candidate
                Exit Do
            End If
        End If
        i = i + 1
    Loop
    m_AutoDrivers(key) = chosen    ' "" = nothing installed can do the requested encryption
    GetAutoDriver = chosen
End Function
' ==== END auto driver functions ====

' ==== BEGIN schema functions ====
Private Function SchemaColumns() As Variant
    ' "name|type" in the order a new SSMAA_ODBC_Tables is created. Older tables (a client project: 10 columns, "Group" instead of
    ' "Groups") get the missing ones added; nothing is dropped or renamed, so older code that still reads "Group" keeps working.
    SchemaColumns = Array("DSN|TEXT(255)", "SQLServer|TEXT(255)", "SQLServerPort|TEXT(255)", "SQLDatabase|TEXT(255)", _
        "SQLLogin|TEXT(100)", "SQLPassword|TEXT(100)", "SQLTableOwner|TEXT(100)", "SQLTable|TEXT(255)", "ODBCTable|TEXT(255)", _
        "Groups|TEXT(50)", "PrimaryKey|TEXT(255)", "Sort|DOUBLE", "ConnectionString|LONGTEXT", "ErrorMessage|TEXT(255)", _
        "Attach|BIT", "ConnectionTypeId|LONG")
End Function

Private Function SchemaTableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    Dim td As DAO.TableDef

    For Each td In db.TableDefs
        If StrComp(td.Name, tableName, vbTextCompare) = 0 Then
            SchemaTableExists = True
            Exit Function
        End If
    Next td
End Function

Private Function SchemaFieldExists(ByVal td As DAO.TableDef, ByVal fieldName As String) As Boolean
    Dim f As DAO.Field

    For Each f In td.Fields
        If StrComp(f.Name, fieldName, vbTextCompare) = 0 Then
            SchemaFieldExists = True
            Exit Function
        End If
    Next f
End Function

Public Sub EnsureSSMAASchema(Optional ByVal force As Boolean = False)
    ' Creates SSMAA_ODBC_Tables and SSMAA_ConnectionTypes when they are missing and completes them when they are old.
    ' Idempotent and non-destructive. Cheap after the first call in a session unless force is True (LinkUsingSSMA passes True).
    Dim db As DAO.Database

    If m_SchemaOk And Not force Then Exit Sub
    Set db = CurrentDb
    EnsureSSMAATable db
    EnsureSSMAAConnectionTypes db
    m_SchemaOk = True
End Sub

Private Sub EnsureSSMAATable(ByVal db As DAO.Database)
    Dim cols As Variant
    Dim parts() As String
    Dim sql As String
    Dim i As Long
    Dim groupsAdded As Boolean

    cols = SchemaColumns()
    If Not SchemaTableExists(db, "SSMAA_ODBC_Tables") Then
        For i = LBound(cols) To UBound(cols)
            parts = Split(cols(i), "|")
            sql = sql & IIf(Len(sql) > 0, ", ", "") & "[" & parts(0) & "] " & parts(1)
        Next i
        db.Execute "CREATE TABLE [SSMAA_ODBC_Tables] (" & sql & ")", dbFailOnError
        db.TableDefs.Refresh
        Exit Sub
    End If

    For i = LBound(cols) To UBound(cols)
        parts = Split(cols(i), "|")
        If Not SchemaFieldExists(db.TableDefs("SSMAA_ODBC_Tables"), parts(0)) Then
            db.Execute "ALTER TABLE [SSMAA_ODBC_Tables] ADD COLUMN [" & parts(0) & "] " & parts(1), dbFailOnError
            db.TableDefs.Refresh
            If parts(0) = "Groups" Then groupsAdded = True
        End If
    Next i

    ' The first time "Groups" appears, take over what the old "Group" column held. Later changes to "Group" are not copied again.
    If groupsAdded Then
        If SchemaFieldExists(db.TableDefs("SSMAA_ODBC_Tables"), "Group") Then
            db.Execute "UPDATE [SSMAA_ODBC_Tables] SET [Groups] = [Group] WHERE [Group] IS NOT NULL", dbFailOnError
        End If
    End If
End Sub

Private Sub EnsureSSMAAConnectionTypes(ByVal db As DAO.Database)
    ' Id = position in this list = the value of the ConnectionTypes enum. Missing ids are added; existing names stay as they are.
    Dim names As Variant
    Dim rs As DAO.Recordset
    Dim i As Long

    names = Array("SQLNCLI", "SQLOLEDBProvider", "LocalDSN", "SQLODBC_Deprecated", "MySQL", "SQLODBC", "SQLAuto")
    If Not SchemaTableExists(db, "SSMAA_ConnectionTypes") Then
        db.Execute "CREATE TABLE [SSMAA_ConnectionTypes] ([Id] LONG, [Name] TEXT(255))", dbFailOnError
        db.TableDefs.Refresh
    End If
    For i = LBound(names) To UBound(names)
        Set rs = db.OpenRecordset("SELECT [Id] FROM [SSMAA_ConnectionTypes] WHERE [Id] = " & i, dbOpenSnapshot)
        If rs.EOF Then
            db.Execute "INSERT INTO [SSMAA_ConnectionTypes] ([Id], [Name]) VALUES (" & i & ", '" & names(i) & "')", dbFailOnError
        End If
        rs.Close
    Next i
End Sub
' ==== END schema functions ====

Public Sub ResetAutoDriver()
    ' Forget the automatic driver choice and every cached connection string (also ServerCon), e.g. at the start of a link round.
    If Not m_AutoDrivers Is Nothing Then m_AutoDrivers.RemoveAll
    Set omCSLastUsed = New omConnectionString
    m_ServerCon_ODBC = vbNullString
    m_ServerCon = vbNullString
End Sub
