﻿Attribute VB_Name = "omSystemProcessFunctions"
Option Compare Database
Option Explicit

Function IsProcessRunning(strProcess As String, Optional strServer As String = "127.0.0.1") As Boolean
    Dim Process, strObject
    IsProcessRunning = False
    strObject = "winmgmts://" & strServer
    For Each Process In GetObject(strObject).InstancesOf("win32_process")
    If UCase(Process.Name) = UCase(strProcess) Then
            IsProcessRunning = True
            Exit Function
        End If
    Next
End Function
