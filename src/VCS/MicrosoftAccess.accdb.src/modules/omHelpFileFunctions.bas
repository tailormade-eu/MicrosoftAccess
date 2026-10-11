Attribute VB_Name = "omHelpFileFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-11 sha=4505c64b

' Help text per form or report, kept in the table omHelpFile and edited in the form omHelpFile_Edit.
' The table belongs to this module: EnsureHelpFileTable creates it when there is no table with that name (a linked one is left as it is).
'   Id = automatic number, Name = name of the form or report, Memo = the help text
' Use: a button on a form with the control source  =OpenHelp([Screen].[ActiveForm])
Public Const HELP_TABLE As String = "omHelpFile"

' Creates the table omHelpFile when it does not exist.
Public Sub EnsureHelpFileTable()
Dim db As DAO.Database
Dim td As DAO.TableDef

    Set db = CurrentDb
    On Error Resume Next
    Set td = db.TableDefs(HELP_TABLE)
    On Error GoTo 0
    If Not td Is Nothing Then Exit Sub
    db.Execute "CREATE TABLE [" & HELP_TABLE & "] ([Id] COUNTER CONSTRAINT PK_" & HELP_TABLE & " PRIMARY KEY, [Name] TEXT(255), [Memo] MEMO)"
    db.Execute "CREATE UNIQUE INDEX IX_" & HELP_TABLE & "_Name ON [" & HELP_TABLE & "] ([Name])"
End Sub

' The Id of the help text of this form or report; the row is added (with an empty text) when it is not there yet.
Public Function GetHelpFileId(objectName As String) As Long
Dim db As DAO.Database
Dim criteria As String

    EnsureHelpFileTable
    Set db = CurrentDb
    criteria = "Name='" & Replace(objectName, "'", "''") & "'"
    If DCount("*", HELP_TABLE, criteria) = 0 Then
        db.Execute "INSERT INTO [" & HELP_TABLE & "] ([Name]) VALUES ('" & Replace(objectName, "'", "''") & "')"
    End If
    GetHelpFileId = DLookup("Id", HELP_TABLE, criteria)
End Function

' Opens the help text of the form or report (obj) in the form omHelpFile_Edit. Returns True when it was opened.
Public Function OpenHelp(obj As Object) As Boolean
    DoCmd.OpenForm "omHelpFile_Edit", , , "Id=" & GetHelpFileId(obj.Name), acFormEdit, acDialog
    OpenHelp = True
End Function
