﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omUserRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private cmdLoadId As New ADODB.Command
Private Const defaultUserRole = "User"
Public Id As Variant
Private pCode As Variant
Private pName As Variant

Public Property Get Name() As Variant
    Name = Nz(pName, defaultUserRole)
End Property

Public Property Let Name(ByVal vNewValue As Variant)
    pName = vNewValue
End Property
Public Property Get Code() As Variant
    Code = Nz(pCode, defaultUserRole)
End Property

Public Property Let Code(ByVal vNewValue As Variant)
    pCode = vNewValue
End Property


Public Sub LoadById(Id As Long)
Dim rs As ADODB.Recordset

    cmdLoadId.Parameters(0) = Id
    Set rs = cmdLoadId.Execute
    SetFields rs
End Sub

Private Sub Class_Initialize()
    ClearFields
    cmdLoadId.commandText = "SELECT * FROM UserRoles WHERE Id=?"
    cmdLoadId.ActiveConnection = CurrentProject.connection
    cmdLoadId.Parameters.Refresh
End Sub

Private Sub Class_Terminate()
    Set cmdLoadId = Nothing
End Sub

Private Sub SetFields(rs As ADODB.Recordset)
    On Error Resume Next
    If Not rs.EOF Then
        Me.Id = rs("Id")
        pCode = rs("Code")
        pName = rs("Name")
    Else
        ClearFields
    End If
    rs.Close
    Set rs = Nothing
End Sub
Private Sub ClearFields()
    Me.Id = Null
    Me.Code = Null
    Me.Name = Null
End Sub
