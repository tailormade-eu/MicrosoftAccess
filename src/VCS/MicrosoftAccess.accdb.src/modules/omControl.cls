VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim rs As New ADODB.Recordset
Public Id As Long
Public ControlTypeId As Long
Public Name As String
Public Default As String
Public HasNoCaption As Boolean

Public Sub Load(ctrl As Object, Optional autoRename As Boolean = False)
Dim cnt As Long
Dim newName As String

    On Error GoTo Load_Error

    Me.HasNoCaption = False
    Me.Name = ctrl.Name
    If Left(TypeName(ctrl), 4) = "Form" Or Left(TypeName(ctrl), 6) = "Report" Then
        Me.ControlTypeId = 0
    Else
        Me.ControlTypeId = ctrl.ControlType
    End If

    Me.Default = ctrl.Caption
    If omStringFunctions.IsNullOrEmpty(Me.Default) Then
      Me.Default = Me.Name
    End If
    If autoRename Then
        Select Case ControlTypeId
            Case acLabel
                If Left(Me.Name, 3) <> "lbl" Then
                    newName = "lbl"
                End If
            Case acCommandButton
                If Left(Me.Name, 3) <> "cmd" Then
                    newName = "cmd"
                End If
            Case acPage
                If Left(Me.Name, 3) <> "pag" Then
                    newName = "pag"
                End If
            Case acToggleButton
                If Left(Me.Name, 3) <> "tgl" Then
                    newName = "tgl"
                End If
        End Select
        If newName <> "" Then
            Me.Name = newName & omStringFunctions.KeepChars(Me.Default, "abcdefghijklmnopqrtsuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", "")
            cnt = 0
            If ctrl.Name <> Left(Me.Name, 60) & IIf(cnt > 0, cnt, "") Then
                ctrl.Name = Left(Me.Name, 60) & IIf(cnt > 0, cnt, "")
            End If
        End If
    End If
    rs.filter = "ControlTypeId=" & Me.ControlTypeId & " AND Name='" & Me.Name & "'"
    If rs.EOF Then
        rs.AddNew
        rs("ControlTypeId") = Me.ControlTypeId
        rs("Name") = Me.Name
        'rs("Default") = Me.Default
        rs("CreateDate") = Now
        rs.Update
    End If
    rs("LastUsedDate") = Now
    rs.Update
    Id = rs("Id")

    Exit Sub

Load_Error:
  If Err = 438 Then
    Me.HasNoCaption = True
    Exit Sub
  End If
  If Err = 2104 Then
      cnt = cnt + 1
      Resume
  End If
  MsgBox Error & " (" & Err & ")"

End Sub

Private Sub Class_Initialize()
    rs.Open "omControls", CurrentProject.connection, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Class_Terminate()
    rs.Close
    Set rs = Nothing
End Sub
