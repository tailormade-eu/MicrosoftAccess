﻿Attribute VB_Name = "omMSAccessFunctions"
Option Compare Database
Option Explicit

Public Function FormExists(formName As String) As Boolean
    FormExists = False
    On Error Resume Next
    FormExists = (CurrentProject.AllForms(formName).Name = formName)
End Function

Public Function ReportExists(ReportName As String) As Boolean
    ReportExists = False
    On Error Resume Next
    ReportExists = (CurrentProject.AllReports(ReportName).Name = ReportName)
End Function
Public Function TableExists(tableName As String) As Boolean
    TableExists = False
    On Error Resume Next
    TableExists = (CurrentDb.TableDefs(tableName).Name = tableName)
End Function
Public Function QueryExists(queryName As String) As Boolean
    QueryExists = False
    On Error Resume Next
    QueryExists = (CurrentDb.QueryDefs(queryName).Name = queryName)
End Function
'Public Function QueryExists(queryName As String) As Boolean
'  QueryExists = NotIsNullOrEmpty(GetQuerySQL(queryName))
'End Function

Public Sub DeleteQuery(queryName As String)
  If QueryExists(queryName) Then
    CurrentDb.QueryDefs.Delete queryName
  End If
End Sub
Public Sub CreatePassthroughQuery(queryName As String, SQL As String, connection As String)
Dim qd As QueryDef

  DeleteQuery queryName
  Set qd = CurrentDb.CreateQueryDef(queryName)
  qd.Connect = connection
  qd.SQL = SQL
  CurrentDb.QueryDefs.Refresh
End Sub

Public Function GetQuerySQL(queryName As String) As String
  On Error Resume Next
  GetQuerySQL = CurrentDb.QueryDefs(queryName).SQL
End Function

Public Sub HideNavigationPane()
    On Error Resume Next
    DoCmd.NavigateTo "acNavigationCategoryObjectType"           'Select Navigation Pane
    DoCmd.RunCommand acCmdWindowHide
End Sub

Public Sub MinimizeNavigationPane()
    On Error Resume Next
    DoCmd.NavigateTo "acNavigationCategoryObjectType"           'Select Navigation Pane
    DoCmd.Minimize
End Sub

Public Sub UnhideNavigationPane()
    On Error Resume Next
    DoCmd.SelectObject acTable, , True
End Sub

Public Sub ExportAllTables()
Dim accObj As AccessObject

    For Each accObj In CurrentData.AllTables
        On Error Resume Next
        omExportFunctions.ExportQuery accObj.Name
    Next
End Sub

Public Sub ConvertToLocalTables(Optional databaseType As String = ".accdb")
Dim rs As New ADODB.Recordset

    rs.Open "SELECT Name,Database,lv FROM MsysObjects WHERE [Type]=6 AND LV IS NOT NULL", CurrentProject.connection, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        If (omStringFunctions.IsNullOrEmpty(databaseType) Or InStr(1, rs("Database"), databaseType) > 0) And InStr(1, rs("Name"), "~") = 0 Then
            'On Error Resume Next
            Debug.Print rs("Name"), rs("database")
            DoCmd.SelectObject acTable, rs("Name"), True
            DoCmd.RunCommand acCmdConvertLinkedTableToLocal
        End If
        rs.MoveNext
        DoEvents
    Wend
    rs.Close
    Set rs = Nothing
End Sub

Public Sub MakeLinkedTablesLocal()
Dim i As Long
Dim tbl As TableDef

    For i = 0 To CurrentDb.TableDefs.Count - 1
        If CurrentDb.TableDefs(i).Connect <> "" And Left(CurrentDb.TableDefs(i).Name, 4) <> "MSys" And Left(CurrentDb.TableDefs(i).Name, 1) <> "~" Then
            If Right(CurrentProject.Name, 1) = "b" Then
                gLogging.WriteToFile Description:="MakeLinkedTables > Right(CurrentProject.Name, 1) = b"
                DoCmd.SelectObject acTable, CurrentDb.TableDefs(i).Name, True
                gLogging.WriteToFile Description:="MakeLinkedTables > DoCmd.SelectObject acTable"
                DoCmd.RunCommand acCmdConvertLinkedTableToLocal
                gLogging.WriteToFile Description:="MakeLinkedTables > DoCmd.RunCommand acCmdConvertLinkedTableToLocal"
            Else
                gLogging.WriteToFile Description:="MakeLinkedTables > Right(CurrentProject.Name, 1) <> b"
                MakeLinkedTableLocal CurrentDb.TableDefs(i).Name
                gLogging.WriteToFile Description:="MakeLinkedTables > MakeLinkedTableLocal CurrentDb.TableDefs(i).Name"
            End If
        End If
    Next
End Sub

Public Sub MakeLinkedTableLocal(tableName As String, Optional structureOnly As Boolean = False)
Dim tempTableName As String
Dim tbl As dao.TableDef
Dim idx As dao.index
'http://www.geeksengine.com/article/duplicate-access-table.html
'CurrentProject.Connection.Execute "SELECT * INTO T_Accounts FROM Accounts"
'DoCmd.CopyObject , "T_Accounts", acTable, "Accounts"
'DoCmd.TransferDatabase acExport, "Microsoft Access", CurrentDb.Name, acTable, "Accounts", "T_Accounts", StructureOnly:=True

    tempTableName = "TCOPY_" & tableName
    omMSAccessFunctions.DeleteTable tempTableName
    gLogging.WriteToFile Description:="MakeLinkedTable > DeleteTable tempTableName"

    CurrentProject.connection.Execute "SELECT * INTO [" & tempTableName & "] FROM [" & tableName & "]"
    gLogging.WriteToFile Description:="MakeLinkedTable > SELECT INTO FROM"
    ' copy existing indexes
    Workspaces(0).Databases(0).TableDefs.Refresh
    Set tbl = Workspaces(0).Databases(0).TableDefs(tableName)
    For Each idx In tbl.Indexes
        CreateIndexUsingDAO tempTableName, idx.Name, Mid(Replace(Replace(idx.Fields, ";+", "],["), "+", "][") & "]", 2), idx.Primary, idx.Unique, Not idx.IgnoreNulls ' not implemented,idx.Clustered ,idx.Foreign ,idx.Required
        gLogging.WriteToFile Description:="MakeLinkedTable > CreateIndex"
    Next
    DoCmd.DeleteObject acTable, tableName
    gLogging.WriteToFile Description:="MakeLinkedTable > DoCmd.DeleteObject acTable, tableName"
    'DoCmd.Rename tableName, acTable, tempTableName ' Does not work in ACCDR/ACCDE
    If structureOnly Then
        DoCmd.TransferDatabase acExport, "Microsoft Access", CurrentDb.Name, acTable, tempTableName, tableName, structureOnly:=True
    Else
        DoCmd.CopyObject , tableName, acTable, tempTableName
    End If

    DoCmd.DeleteObject acTable, tempTableName
    gLogging.WriteToFile Description:="MakeLinkedTable > DoCmd.Rename tableName, acTable, tempTableName"
    CurrentDb.TableDefs.Refresh
End Sub

Public Sub DropIndexUsingDAO(tableName As String, indexName As String)
    On Error Resume Next
    Workspaces(0).Databases(0).TableDefs(tableName).Indexes.Delete indexName
End Sub

Public Sub CreateIndexUsingDAO(tableName As String, indexName As String, fieldNames As String, Optional setPrimary As Boolean = False, Optional setUnique As Boolean = False, Optional setDisallowNull As Boolean = False)
Dim tbl As dao.TableDef
Dim idx As dao.index
Dim fld As dao.Field
Dim fldNames() As String
Dim fldName As Variant

    Set tbl = Workspaces(0).Databases(0).TableDefs(tableName)
    Set idx = tbl.CreateIndex(indexName)
    idx.Primary = setPrimary
    idx.Unique = setUnique
    fldNames = omStringFunctions.StringSplit(fieldNames, ",")
    For Each fldName In fldNames
        idx.Fields.Append idx.CreateField(Replace(Replace(fldName, "[", ""), "]", ""))
    Next
    DropIndexUsingDAO tableName, indexName
    Workspaces(0).Databases(0).TableDefs(tableName).Indexes.Append idx
    Workspaces(0).Databases(0).TableDefs(tableName).Indexes.Refresh
End Sub

Public Sub DropIndexUsingSQL(tableName As String, indexName As String)
    On Error Resume Next
    CurrentProject.connection.Execute StringFormat("DROP INDEX [{0}] ON [{1}]", indexName, tableName)
End Sub

Public Sub CreateIndexUsingSQL(tableName As String, indexName As String, fieldNames As String, Optional setPrimary As Boolean = False, Optional setUnique As Boolean = False, Optional setDisallowNull As Boolean = False)
Dim SQL As String

    SQL = StringFormat("CREATE" & IIf(setPrimary Or setUnique, " UNIQUE", "") & " INDEX [{0}] ON {1} ({2})", indexName, tableName, fieldNames)
    If setPrimary Then
        SQL = SQL & " WITH PRIMARY"
    ElseIf setDisallowNull Then
        SQL = SQL & " WITH DISALLOW NULL"
    End If
    DropIndexUsingSQL tableName, indexName
    gLogging.WriteToFile Description:="CreateIndexUsingSQL > DropIndexUsingSQL tableName, indexName"
    CurrentProject.connection.Execute SQL
    gLogging.WriteToFile Description:="CreateIndexUsingSQL > CurrentProject.Connection.Execute sql"
End Sub

Public Sub DeleteTables(tablePrefix As String)
Dim i As Long

    For i = CurrentDb.TableDefs.Count - 1 To 0 Step -1
        If Left(CurrentDb.TableDefs(i).Name, Len(tablePrefix)) = tablePrefix Then
            DoCmd.DeleteObject acTable, CurrentDb.TableDefs(i).Name
        End If
    Next i
End Sub

Public Sub SetAccessProperty(propertyName As String, Value As Variant, Optional propertyType As dao.DataTypeEnum = dbText)
Dim prp As dao.Property
    On Error Resume Next
    Set prp = CurrentDb.CreateProperty(propertyName, propertyType, Value)
    CurrentDb.Properties.Append prp
    If Err = 3367 Then
        CurrentDb.Properties(propertyName) = Value
    End If
    On Error GoTo 0
End Sub

Public Sub DeleteTable(tableName As String)
    If TableExists(tableName) Then
        DoCmd.DeleteObject acTable, tableName
    End If
End Sub

Public Function IsTableLocal(tableName As String) As Boolean
    If TableExists(tableName) Then
        IsTableLocal = (CurrentDb.TableDefs(tableName).Connect = "")
    End If
End Function

Public Sub FormFields_Extract()
Dim rsMSysObject As New ADODB.Recordset
Dim rsFields As New ADODB.Recordset
Dim ObjectTemp As Object
Dim ctl As Control

    On Error GoTo FormFields_Extract_Error

    rsMSysObject.Open "SELECT Name, Type FROM msysobjects WHERE (((Type)=-32768 Or (Type)=-32764))", CurrentProject.connection, adOpenDynamic, adLockOptimistic
    rsFields.Open "Fields", CurrentProject.connection, adOpenDynamic, adLockOptimistic
    While Not rsMSysObject.EOF
        Select Case rsMSysObject("Type")
            Case -32768 ' Form
                DoCmd.OpenForm rsMSysObject("Name"), acDesign, , , , acHidden
                Set ObjectTemp = Forms(rsMSysObject("Name"))
            Case -32764 ' Report
                DoCmd.OpenReport rsMSysObject("Name"), acViewDesign
                DoCmd.Minimize
                Set ObjectTemp = Reports(rsMSysObject("Name"))
        End Select
        For Each ctl In ObjectTemp.Controls
            With ctl
                If .ControlType = acLabel Or .ControlType = acCommandButton Then
                    rsFields.AddNew
                    rsFields("Field_ID") = newField_ID
                    rsFields("Field_Name") = .Name
                    rsFields.Update
                End If
            End With
        Next
        rsFields.AddNew
        rsFields("Field_ID") = newField_ID
        rsFields("Field_Name") = ObjectTemp.Name
        rsFields.Update
        Select Case rsMSysObject("Type")
            Case -32768 ' Form
                DoCmd.Close acForm, ObjectTemp.Name
            Case -32764 ' Report
                DoCmd.Close acReport, ObjectTemp.Name
        End Select
        rsMSysObject.MoveNext
    Wend
    rsMSysObject.Close
    rsFields.Close
    Set rsMSysObject = Nothing
    Set rsFields = Nothing
    Exit Sub

FormFields_Extract_Error:

    Select Case Err
        Case 3022
            Resume Next
        Case Else
            Exit Sub
    End Select

End Sub
Public Sub cmdBarFields_Extract(strcmdBar As String, ctrl As Object)
Dim rsFields As New ADODB.Recordset
Dim ctl As Object

    On Error GoTo cmdBarFields_Extract_Error

    rsFields.Open "Fields", CurrentProject.connection, adOpenDynamic, adLockOptimistic
    If strcmdBar <> "" Then
        For Each ctl In CommandBars(strcmdBar).Controls
            With ctl
                rsFields.AddNew
                rsFields("Field_ID") = newField_ID
                rsFields("Field_Name") = .tag
                rsFields.Update
                If .Type = 10 Then
                    cmdBarFields_Extract "", ctl
                End If
            End With
        Next
    Else
        For Each ctl In ctrl.Controls
            With ctl
                rsFields.AddNew
                rsFields("Field_ID") = newField_ID
                rsFields("Field_Name") = .tag
                rsFields.Update
                If .Type = 10 Then
                    cmdBarFields_Extract "", ctl
                End If
            End With
        Next
    End If
    rsFields.Close
    Set rsFields = Nothing
    Exit Sub

cmdBarFields_Extract_Error:

    Select Case Err
        Case 3022
            Resume Next
        Case Else
            Exit Sub
    End Select

End Sub
Public Function newField_ID() As Long

    If IsNull(DMax("Field_ID", "Fields")) Then
        newField_ID = 1
    Else
        newField_ID = DMax("Field_ID", "Fields") + 1
    End If

End Function

Public Sub RegisterCurrentLocation()
'Dim c As New cRegistry
'    With c
'        .ClassKey = HKEY_CURRENT_USER
'        .SectionKey = "Software\Microsoft\Office\16.0\Access\Security\Trusted Locations\YourTrustedLocationName"
'        .ValueKey = "Path"
'        .ValueType = REG_DWORD
'        .value = CurrentProject.path
'    End With
End Sub

Public Function IsRuntimeMode() As Boolean
    IsRuntimeMode = SysCmd(acSysCmdRuntime)
End Function


Public Sub TruncateMSAccessTable(tableName As String)
Dim strSQL As String

    strSQL = "DELETE FROM [" & tableName & "]"
    CurrentProject.connection.Execute strSQL
    strSQL = "ALTER TABLE [" & tableName & "] ALTER COLUMN id COUNTER (1, 1);"
    CurrentProject.connection.Execute strSQL
End Sub
