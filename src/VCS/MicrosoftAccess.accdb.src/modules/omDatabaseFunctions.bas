Attribute VB_Name = "omDatabaseFunctions"
Option Compare Database
Option Explicit

Public Sub LinkTable(databaseFilename As String, sourceName As String, destinationName As String)
    UnlinkTable destinationName
    DoCmd.TransferDatabase acLink, "Microsoft Access", databaseFilename, acTable, sourceName, destinationName
End Sub
Public Sub UnlinkTable(Name As String)
    On Error Resume Next
    DoCmd.DeleteObject acTable, Name
    On Error GoTo 0
End Sub
Public Function CreateDataBase(filename As String) As Boolean
Dim varReturn As Variant

    On Error GoTo CreateDB_Error
    CreateDataBase = True
    DBEngine.CreateDataBase filename, dbLangGeneral

    Exit Function

CreateDB_Error:
    Select Case Err
        Case Else
            CreateDataBase = False
    End Select
End Function

Public Function ExportTable(filename As String, sourceName As String, destinationName As String) As Boolean
    On Error GoTo ExportTable_Error
    ExportTable = True
    DoCmd.TransferDatabase acExport, "Microsoft Access", filename, acTable, sourceName, destinationName, False
    Exit Function
ExportTable_Error:
    ExportTable = False
End Function

Public Function CompactDataBase(filename As String)
Dim dstFilename As String

    On Error GoTo CompactDataBase_Error
    CompactDataBase = True
    dstFilename = filename & "_" & (((year(Date) * 100 + Month(Date)) * 100 + Day(Date)) * 100 + Hour(Time)) * 100 + Minute(Time)
    DBEngine.CompactDataBase filename, dstFilename
    DeleteFile filename
    RenameFile dstFilename, filename
    Exit Function
CompactDataBase_Error:
    CompactDataBase = False
End Function

Public Function CopyTable(databaseFilename As String, TableSource As String, TableDestination As String) As Variant

    On Error Resume Next
    DoCmd.DeleteObject acTable, "_" & TableSource
    On Error GoTo 0
    DoCmd.TransferDatabase acImport, "Microsoft Access", databaseFilename, acTable, TableSource, "_" & TableSource
    DoCmd.TransferDatabase acExport, "Microsoft Access", databaseFilename, acTable, "_" & TableSource, TableDestination
    DoCmd.DeleteObject acTable, "_" & TableSource
End Function
