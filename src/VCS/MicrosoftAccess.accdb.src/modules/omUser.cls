VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private cmdLoginByName As New ADODB.Command
Private cmdLoginById As New ADODB.Command
Private cmdLoginByNameOnly As New ADODB.Command
Private cmdUpdatePassword As New ADODB.Command

Public Id As Variant
Public Code As Variant
Public Name As Variant
Public Active As Boolean
Public UserRole As New omUserRole

Public Sub LoginByNameOnly(LoginName, Optional windowsLogon As Boolean = True, Optional addWindowsLogon As Boolean = False)
Dim rs As ADODB.Recordset

    cmdLoginByNameOnly.Parameters(0) = LoginName
    Set rs = cmdLoginByNameOnly.Execute
    SetFields rs
    If Nz(Me.Id, 0) = 0 Then
        If addWindowsLogon Then
            Set rs = New ADODB.Recordset
            rs.Open "Users", CurrentProject.connection, adOpenForwardOnly, adLockOptimistic
            rs.AddNew
            rs("Name") = LoginName
            rs("Login") = LoginName
            rs("Active") = True
            rs.Update
            Me.Id = rs("Id")
            rs.Close
            Set rs = Nothing
        ElseIf windowsLogon Then
            Me.Id = -1
        End If
        Me.Code = LoginName
        Me.Name = LoginName
        Me.Active = True
    End If
End Sub
Public Function LoginById(UserId As Long, Password As String) As Boolean
Dim rs As ADODB.Recordset

    cmdLoginById.Parameters(0) = UserId
    cmdLoginById.Parameters(1) = Password
    Set rs = cmdLoginById.Execute
    SetFields rs
    LoginById = Me.LoggedIn
End Function

Public Function LoginByName(LoginName As String, Password As String) As Boolean
Dim rs As ADODB.Recordset

    cmdLoginByName.Parameters(0) = LoginName
    cmdLoginByName.Parameters(1) = Password
    Set rs = cmdLoginByName.Execute
    SetFields rs
    LoginByName = Me.LoggedIn
End Function
Public Sub UpdatePassword(Password As String)
    cmdUpdatePassword.Parameters(0) = Password
    cmdUpdatePassword.Parameters(1) = Me.Id
    cmdUpdatePassword.Execute
End Sub
Private Sub Class_Initialize()
    ClearFields
    cmdLoginByName.commandText = "SELECT * FROM Users WHERE Login=? AND Password=?"
    cmdLoginByName.ActiveConnection = CurrentProject.connection
    cmdLoginByName.Parameters.Refresh
    cmdLoginById.commandText = "SELECT * FROM Users WHERE Id=? AND Password=?"
    cmdLoginById.ActiveConnection = CurrentProject.connection
    cmdLoginById.Parameters.Refresh
    cmdLoginByNameOnly.commandText = "SELECT * FROM Users WHERE Login=?"
    cmdLoginByNameOnly.ActiveConnection = CurrentProject.connection
    cmdLoginByNameOnly.Parameters.Refresh
    cmdUpdatePassword.commandText = "UPDATE Users SET [Password]=? WHERE Id=?"
    cmdUpdatePassword.ActiveConnection = CurrentProject.connection
    cmdUpdatePassword.Parameters.Refresh
End Sub

Private Sub Class_Terminate()
    Set cmdLoginByName = Nothing
    Set cmdLoginById = Nothing
    Set cmdLoginByNameOnly = Nothing
    Set cmdUpdatePassword = Nothing
End Sub

Private Sub SetFields(rs As ADODB.Recordset)
    On Error Resume Next
    If Not rs.EOF Then
        Me.Id = rs("Id")
        Me.Code = rs("Code")
        Me.Name = rs("Name")
        If omStringFunctions.IsNullOrEmpty(Me.Name) Then
            Me.Name = Nz(rs("FirstName"), "") & " " & Nz(rs("LastName"), "")
        End If
        If omStringFunctions.IsNullOrEmpty(Me.Name) Then
            Me.Name = rs("Login")
        End If
        Me.Active = Nz(rs("Active"), False)
        Me.UserRole.LoadById rs("UserRoleId")
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
    Me.Active = False
End Sub
Public Function LoggedIn() As Boolean
    LoggedIn = omStringFunctions.NotIsNullOrEmptyOrZero(Me.Id) Or Nz(Me.Id, 0) = -1
End Function
