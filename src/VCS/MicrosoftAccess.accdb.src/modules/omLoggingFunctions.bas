﻿Attribute VB_Name = "omLoggingFunctions"
Option Compare Database
Option Explicit

Public gLogging As omLogging


Public Sub StartLogging(Optional enable As Boolean = False)
    Set gLogging = New omLogging
    SetLoggingEnable enable
End Sub

Public Sub SetLoggingEnable(enable As Boolean)
    If omObjectFunctions.NotIsNothing(gLogging) Then
        gLogging.Enabled = enable
    End If
End Sub

Public Function WriteToLog(Optional diffOnly As String = "y", Optional updateCurrentDate As Boolean = True, Optional Description As String = "") As String
    If omObjectFunctions.NotIsNothing(gLogging) Then
        WriteToLog = gLogging.WriteToStream(diffOnly, updateCurrentDate, Description)
    End If
End Function
Public Sub StopLogging()
    Set gLogging = Nothing
End Sub
