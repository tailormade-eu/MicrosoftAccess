﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Code written by Raoul Jacobs
' E. jara@opmaat.be
' Date Craeted : 200703
'
' this Class is used to get the value from the default table for any defined variable in this table

Public Enum DefaultModes
    LocalMode = 1
    serverMode = 2
End Enum

Dim cmdGetValue As New ADODB.Command
Dim cmdSetValue As New ADODB.Command

Public Name As String
Public Value As Variant
Public ModifyDate As Variant
Public Development As Boolean
Private gMode As DefaultModes
Public Initialized As Boolean
Public Function Load(defaultName As String) As Variant
Dim rs As ADODB.Recordset

    If Me.Name <> defaultName Then
        cmdGetValue.Parameters(0) = defaultName & IIf(Development, "_dev", "")
        Set rs = cmdGetValue.Execute
        If rs.EOF And Development Then
            cmdGetValue.Parameters(0) = defaultName
            Set rs = cmdGetValue.Execute
        End If
        If Not rs.EOF Then
            Me.Value = rs(0)
            Me.ModifyDate = rs(1)
        Else
            Me.Name = ""
            Me.Value = Null
            Me.ModifyDate = Null
        End If
    End If
    rs.Close
    Set rs = Nothing
    Load = Me.Value
End Function
Public Function Save(defaultName As String, Value As Variant)
    cmdSetValue.Parameters(0) = Value
    cmdSetValue.Parameters(2) = defaultName
    cmdSetValue.Parameters(1) = Now
    cmdSetValue.Execute
End Function
Private Sub Class_Initialize()
Dim pm As ADODB.Parameter
Dim tableName As String

    If gMode = 0 Then Exit Sub

    tableName = IIf(gMode = LocalMode, "omSysDefaults", "omDefaults")
    Initialized = False
    If omMSAccessFunctions.TableExists(tableName) Then
        cmdGetValue.commandText = "SELECT [Value],[ModifyDate] FROM " & tableName & " WHERE [Name]=?"
        cmdGetValue.ActiveConnection = CurrentProject.connection
        cmdGetValue.Parameters.Refresh ' Does Not Work if there is a mistake in the SQL statement EG not existing Table
        'cmdGetValue.Parameters.Append cmdGetValue.CreateParameter("Name", adVarWChar, adParamInput, 256)

        cmdSetValue.commandText = "UPDATE [" & IIf(gMode = LocalMode, "omSysDefaults", "omDefaults") & "] SET [Value]=?,[ModifyDate]=? WHERE [Name]=?"
        cmdSetValue.ActiveConnection = CurrentProject.connection
        cmdSetValue.Parameters.Refresh
        'cmdSetValue.Parameters.Append cmdGetValue.CreateParameter("Value", adVarChar, adParamInput, 256)
        'cmdSetValue.Parameters.Append cmdGetValue.CreateParameter("ModifyDate", adDate, adParamInput)
        'cmdSetValue.Parameters.Append cmdGetValue.CreateParameter("Name", adVarChar, adParamInput, 256)
        Initialized = True
    End If
End Sub

Private Sub Class_Terminate()
    Set cmdGetValue = Nothing
End Sub

Public Static Property Get Mode() As DefaultModes
    Mode = gMode
End Property

Public Static Property Let Mode(ByVal vNewValue As DefaultModes)
    gMode = vNewValue
    Class_Initialize
End Property
