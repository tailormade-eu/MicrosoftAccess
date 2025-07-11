﻿Attribute VB_Name = "omCommandBarFunctions"
Option Compare Database
Option Explicit

Public Const omPrintContextMenu = "omPrintContextMenu"

Public Sub RegisterPrintContextMenu()

    On Error Resume Next
    Dim cbar As CommandBar
    Dim bt As CommandBarButton
    'delete first if already exists
    CommandBars.Item(omPrintContextMenu).Delete
    'recreate
    Set cbar = CommandBars.Add(omPrintContextMenu, msoBarPopup, , False)
    Set bt = cbar.Controls.Add
    With bt
        .Caption = "&Print"
        .OnAction = "=fnPrint()"
        .FaceId = 15948
    End With

End Sub

Public Function ListCommandBarObjects(Optional contains As String = "")
  'List all the CommandBar objects in the
  'application in the Immediate window.
  Dim cb As CommandBar
  For Each cb In CommandBars
    If IsNullOrEmpty(contains) Or omStringFunctions.ContainsString(cb.Name, contains) Then
        Debug.Print cb.Name, vbTab, cb.Type, vbTab, cb.index
    End If
  Next
End Function


Public Function fnPrint()

On Error Resume Next

    DoCmd.RunCommand acCmdPrint

End Function
