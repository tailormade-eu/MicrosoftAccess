﻿Attribute VB_Name = "omTableFunctions"
Option Compare Database
Option Explicit

Public Sub TableAddId(tableName As String)
    AddField tableName, "ID", dbLong, True
End Sub

Public Sub AddField(tableName As String, fieldName As String, FieldType As dao.DataTypeEnum, Optional AutoIncrement As Boolean = False)
Dim fld As dao.Field

        Set fld = CurrentDb.TableDefs(tableName).CreateField(fieldName, FieldType)
        If AutoIncrement Then
            fld.Attributes = fld.Attributes + dbAutoIncrField
        End If
        CurrentDb.TableDefs(tableName).Fields.Append fld
        CurrentDb.TableDefs(tableName).Fields.Refresh
End Sub

Sub AttachedTable(ConnectionString As String, SourceTable As String, DestinationTable As String)
Dim tbl As dao.TableDef

    'Create a new TableDef object.
    Set tbl = CurrentDb.CreateTableDef(DestinationTable)
    'Set the properties to create the link
    tbl.Connect = ConnectionString
    tbl.SourceTableName = SourceTable
    'Add the new table to the database.
    CurrentDb.TableDefs.Append tbl
    Set tbl = Nothing
End Sub
Public Function GenerateSelect(tableName As String, Optional includeTableName As Boolean = False, Optional excludeFields As String) As String
Dim rs As New ADODB.Recordset
Dim fld As ADODB.Field
Dim strSelect As String
Dim strFrom As String
Dim exFlds() As String

    exFlds = Split(excludeFields, ",")

    rs.Open "[" & tableName & "]", CurrentProject.connection, adOpenForwardOnly, adLockReadOnly
    For Each fld In rs.Fields
        If omArrayFunctions.StringArrayFind(exFlds, fld.Name, False) = -1 Then
            If includeTableName Then
                strSelect = strSelect & "[" & tableName & "]."
            End If
            strSelect = strSelect & "[" & fld.Name & "],"
        End If
    Next
    strSelect = Left(strSelect, Len(strSelect) - 1)
    strFrom = "[" & tableName & "]"
    rs.Close
    Set rs = Nothing
    GenerateSelect = omSQLFunctions.BuildSQL(strSelect, strFrom)
End Function
Public Function GenerateInsert(tableName As String, Optional idFieldName As String = "Id", Optional excludeFields As String, Optional excludeValuesPart As Boolean = False) As String
Dim rs As New ADODB.Recordset
Dim fld As ADODB.Field
Dim strInsert As String
Dim strFrom As String
Dim strInsertFields
Dim strInsertValues
Dim exFlds() As String

    exFlds = Split(excludeFields, ",")
    rs.Open "[" & tableName & "]", CurrentProject.connection, adOpenForwardOnly, adLockReadOnly
    For Each fld In rs.Fields
        If fld.Name <> idFieldName Then
            If omArrayFunctions.StringArrayFind(exFlds, fld.Name, False) = -1 Then
                strInsertFields = strInsertFields & "[" & fld.Name & "],"
                strInsertValues = strInsertValues & "?,"
            End If
        End If
    Next
    strInsertFields = Left(strInsertFields, Len(strInsertFields) - 1)
    strInsertValues = Left(strInsertFields, Len(strInsertFields) - 1)
    rs.Close
    Set rs = Nothing
    GenerateInsert = "INSERT INTO [" & tableName & "] (" & strInsertFields & ")"
    If Not excludeValuesPart Then
        GenerateInsert = GenerateInsert & " VALUES (" & strInsertValues & ")"
    End If
End Function
Public Function GenerateUpdate(tableName As String, Optional idFieldName As String = "Id", Optional excludeFields As String) As String
Dim rs As New ADODB.Recordset
Dim fld As ADODB.Field
Dim strInsert As String
Dim strFrom As String
Dim strUpdateFields
Dim exFlds() As String

    exFlds = Split(excludeFields, ",")
    rs.Open "[" & tableName & "]", CurrentProject.connection, adOpenForwardOnly, adLockReadOnly
    For Each fld In rs.Fields
        If fld.Name <> idFieldName Then
            If omArrayFunctions.StringArrayFind(exFlds, fld.Name, False) = -1 Then
                strUpdateFields = strUpdateFields & "[" & fld.Name & "]=?,"
            End If
        End If
    Next
    strUpdateFields = Left(strUpdateFields, Len(strUpdateFields) - 1)
    rs.Close
    Set rs = Nothing
    GenerateUpdate = "UPDATE [" & tableName & "] SET " & strUpdateFields & " WHERE [" & idFieldName & "]=?"
End Function
