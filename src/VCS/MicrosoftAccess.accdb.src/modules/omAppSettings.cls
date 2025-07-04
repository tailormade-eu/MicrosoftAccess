VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omAppSettings"
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

Dim cmdGetValue As New ADODB.Command
Public Name As String
Public Value As Variant

Public Function Load(SettingsName As String) As Variant
Dim rs As ADODB.Recordset

    If Me.Name <> SettingsName Then
        cmdGetValue.Parameters(0) = SettingsName
        Set rs = cmdGetValue.Execute
        If Not rs.EOF Then
            Me.Value = rs(0)
        Else
            Me.Name = ""
            Me.Value = ""
        End If
        rs.Close
    End If
    Set rs = Nothing
    Load = Me.Value
End Function
Public Sub Save(SettingsName As String, NewValue As String)
Dim rs As New ADODB.Recordset

    rs.Open "SELECT * FROM appSettings WHERE Name=" & Chr(34) & SettingsName & Chr(34), CurrentProject.connection, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        rs.AddNew
        rs("Name") = SettingsName
    End If
    rs("Value") = NewValue
    rs.Update
    rs.Close
    Set rs = Nothing
    Name = SettingsName
    Value = NewValue
End Sub

Private Sub Class_Initialize()
    cmdGetValue.commandText = "SELECT Value FROM appSettings WHERE Name=?"
    cmdGetValue.ActiveConnection = CurrentProject.connection
    cmdGetValue.Parameters.Refresh
End Sub

Private Sub Class_Terminate()
    Set cmdGetValue = Nothing
End Sub
