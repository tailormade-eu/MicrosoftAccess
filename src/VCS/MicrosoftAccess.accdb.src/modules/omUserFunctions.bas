Attribute VB_Name = "omUserFunctions"
Option Compare Database
Option Explicit
'@lib 2026-11-06 sha=9b273945

Public gUser As New omUser
' How the user logs in. The client sets exactly one of the two to True in ClientSettings_LoadOverride (see
' omGlobalDeclarations). There is no standard: both False means "not set", and AuthenticateUser then says so
' instead of silently choosing a login screen.
Public UseWindowsAuthentication As Boolean
Public UseLoginForm As Boolean

' True when exactly one of the two login methods is set.
Public Function LoginMethodIsSet() As Boolean
    omGlobalDeclarations.EnsureClientSettings
    LoginMethodIsSet = (UseWindowsAuthentication <> UseLoginForm)
End Function

Public Function GetWindowsUserName() As String
    GetWindowsUserName = Environ$("UserName")
End Function
Public Function GetWindowsComputerName() As String
    GetWindowsComputerName = Environ$("computername")
End Function

Public Function LoginWithWindowsUser(Optional addWindowsLogon As Boolean = False)
    Set gUser = New omUser
    gUser.LoginByNameOnly GetWindowsUserName, addWindowsLogon:=True
End Function
Public Function UseWindowsUser()
    Set gUser = New omUser
    gUser.Name = GetWindowsUserName
End Function
Public Sub UserLock()
    If omMSAccessFunctions.FormExists("UserLock") Then
        DoCmd.OpenForm "UserLock", , , , acFormReadOnly, acHidden
    End If
End Sub

Public Sub AuthenticateUser()
    If gUser.LoggedIn Then
        Exit Sub
    End If
    If Not LoginMethodIsSet Then
        MsgBox "Set UseWindowsAuthentication or UseLoginForm (exactly one) in ClientSettings_LoadOverride.", vbExclamation
        Exit Sub
    End If
    If UseWindowsAuthentication Then
        omUserFunctions.LoginWithWindowsUser True
    ElseIf UseLoginForm Then
        DoCmd.OpenForm "Login", windowMode:=acDialog
    End If
End Sub

Public Sub RecordsetUpdateCreateTracking(rs As ADODB.Recordset, Optional UpdateModify = True)
    rs("CreateDate") = Now
    On Error Resume Next
    omUserFunctions.AuthenticateUser
    If gUser.LoggedIn Then
        rs("CreateUserName") = gUser.Name
        rs("CreateUserId") = gUser.Id
    Else
        rs("CreateUserName") = gUser.Name
        rs("CreateUserId") = gUser.Name
    End If
    If UpdateModify Then
        RecordsetUpdateModifyTracking rs
    End If
End Sub
Public Sub RecordsetUpdateModifyTracking(rs As ADODB.Recordset)
    rs("ModifyDate") = Now
    On Error Resume Next
    omUserFunctions.AuthenticateUser
    If gUser.LoggedIn Then
        rs("ModifyName") = gUser.Name
        rs("ModifyUserId") = gUser.Id
    Else
        rs("ModifyUserName") = gUser.Name
        rs("ModifyUserId") = gUser.Name
    End If
End Sub
Public Function GetUserId() As Long
    GetUserId = Nz(gUser.Id, 0)
End Function

Public Function HasRole(roleCode As String) As Boolean
    HasRole = (UCase(Nz(gUser.UserRole.Code, "")) = UCase(roleCode))
End Function
