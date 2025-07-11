﻿Attribute VB_Name = "omFormFunctions"
Option Compare Database
Option Explicit

Public Sub UpdateEnableByTag(Form As Form, tag As String, state As Boolean, Optional delimiter = ",")
Dim ctrl As Control

    For Each ctrl In Form.Controls
        If omStringFunctions.NotIsNullOrEmpty(ctrl.tag) And (InStr(1, delimiter & ctrl.tag & delimiter, delimiter & tag & delimiter) > 0 Or omStringFunctions.IsNullOrEmpty(tag)) Then
            ctrl.Enabled = state
        End If
    Next
End Sub

Public Sub UpdateVisibleByTag(Form As Form, tag As String, state As Boolean, Optional delimiter = ",")
Dim ctrl As Control

    For Each ctrl In Form.Controls
        If omStringFunctions.NotIsNullOrEmpty(ctrl.tag) And (InStr(1, delimiter & ctrl.tag & delimiter, delimiter & tag & delimiter) > 0 Or omStringFunctions.IsNullOrEmpty(tag)) Then
            ctrl.visible = state
        End If
    Next
End Sub
Public Function OpenForm(Name As String, Optional view As AcFormView = AcFormView.acFormDS) As Variant
    Name = Replace(Name, "cmd", "")
    If InStr(1, Name, "listsearch") <> 0 Then
        Name = Replace(Name, "listsearch", "_List_Search")
    ElseIf InStr(1, Name, "list") <> 0 Then
        Name = Replace(Name, "list", "_List")
        view = acNormal
    End If
    DoCmd.OpenForm Name, view
End Function

Public Sub CloseForms(Optional keepOpen As String = "flow")
Dim frm As Form
    For Each frm In Forms
        If frm.Name <> keepOpen Then
            DoCmd.Close acForm, frm.Name, acSaveNo
        End If
    Next
End Sub

Public Sub OpenEditScreen(parentForm As Form, formName As String, Optional keyName As String = "Id", Optional windowMode As AcWindowMode = acDialog, Optional Requery As Boolean = True)
Dim Id As String

    If parentForm.CurrentRecord <> -1 Then
        If parentForm.Recordset.Fields(keyName).Type = dbGUID Then
            Id = StringFromGUID(parentForm.Recordset.Fields(keyName).Value)
        Else
            Id = parentForm.Recordset.Fields(keyName).Value
        End If
        DoCmd.OpenForm formName, datamode:=acFormEdit, windowMode:=windowMode, whereCondition:=keyName & " =" & Id
        If Requery Then
            parentForm.Requery
        End If
    End If
End Sub

Public Sub UpdateModifyTracking(frm As Form)
    frm.ModifyDate = Now
    On Error Resume Next
    omUserFunctions.AuthenticateUser
    If gUser.LoggedIn Then
        frm.ModifyUserName = gUser.Name
        frm.ModifyUserId = gUser.Id
    Else
        frm.ModifyUserName = gUser.Name
        frm.ModifyUserId = gUser.Name
    End If
End Sub
Public Sub UpdateCreateTracking(frm As Form, Optional UpdateModify = True)
    frm.CreateDate = Now
    On Error Resume Next
    omUserFunctions.AuthenticateUser
    If gUser.LoggedIn Then
        frm.CreateUserName = gUser.Name
        frm.CreateUserId = gUser.Id
    Else
        frm.CreateUserName = gUser.Name
        frm.CreateUserId = gUser.Name
    End If
    If UpdateModify Then
        UpdateModifyTracking frm
    End If
End Sub
Public Sub ListFormFields(formName As String)
    Dim frm As Form
    Dim ctl As Control
    Dim intControlOrder As Integer
    Dim ts As Scripting.TextStream

    Set ts = gFso.CreateTextFile(formName, True)
    Set frm = Forms(formName)

    For Each ctl In frm.Controls
        Debug.Print "Control Name: " & ctl.Name
        ts.Write "Control Name|" & ctl.Name & "|"
        Debug.Print "Control Type: " & TypeName(ctl)
        ts.Write "Control Type|" & TypeName(ctl) & "|"
        On Error Resume Next
        If Not ctl.Properties("Caption") Is Nothing Then
            Debug.Print "Control Caption: " & ctl.Properties("Caption")
            ts.Write "Control Caption|" & ctl.Properties("Caption") & "|"
        End If
        If Not ctl.Properties("LabelName") Is Nothing Then
            Debug.Print "Related Label: " & ctl.Properties("LabelName"), frm.Controls(ctl.Properties("LabelName")).Caption
            ts.Write "Related Label|" & ctl.Properties("LabelName") & "|"
            ts.Write "Related Label Caption|" & frm.Controls(ctl.Properties("LabelName")).Caption & "|"
        End If
        intControlOrder = ctl.Properties("Order")
        Debug.Print "Order: " & intControlOrder
        Debug.Print "----------------------"
        ts.Write (vbCrLf)
    Next ctl
    ts.Close
    Set ts = Nothing
End Sub
