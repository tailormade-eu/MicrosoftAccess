﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omLogging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Enabled As Boolean
Public StartDate As Date
Public EndDate As Date
Public PreviousDate As Date
Public PreviousDescription As String
Public CurrentDescription As String
Private dtCurrent As Date
Private startTicks As Long
Private previousTicks As Long
Private currentTicks As Long
Private endTicks As Long
Dim ts As TextStream

Private Sub Class_Initialize()
    Me.Reset
    Set ts = gFso.CreateTextFile(gFso.BuildPath(CurrentProject.path, "log_" & omDateFunctions.GetTimeStamp() & ".txt"))
End Sub

Public Property Get currentdate() As Date
    currentdate = dtCurrent
End Property

Public Sub SetCurrentDate(Optional dt As Date = 0, Optional Description As String = "")
    If dt = 0 Then
        dt = Now
    End If
    Me.currentdate = dt
    Me.CurrentDescription = Description
End Sub

Public Property Let currentdate(ByVal vNewValue As Date)
    PreviousDate = dtCurrent
    PreviousDescription = CurrentDescription
    previousTicks = currentTicks
    dtCurrent = vNewValue
    currentTicks = omKernalFunctions.GetTickCount()
    If StartDate = 0 Then
        StartDate = dtCurrent
        PreviousDate = dtCurrent
        startTicks = currentTicks
        previousTicks = currentTicks
    End If
    If EndDate < dtCurrent Then
        EndDate = dtCurrent
        endTicks = currentTicks
    End If
End Property

Public Function WriteToFile(Optional diffOnly As String = "y", Optional updateCurrentDate As Boolean = True, Optional Description As String = "") As String
'Dim ts As TextStream
    If Not Enabled Then
        Exit Function
    End If
    'Set ts = gFso.OpenTextFile(filename, ForAppending, True)
    ts.WriteLine ToString(diffOnly, updateCurrentDate, Description)
    ts.Close
    Set ts = Nothing
End Function

Public Function ToString(Optional diffOnly As String = "y", Optional updateCurrentDate As Boolean = True, Optional Description As String = "") As String
    If updateCurrentDate Then
        SetCurrentDate Description:=Description
    End If
    If Not Enabled Then
        Exit Function
    End If
    If Len(Description) > 0 Then
        ToString = "Description: " & Description & vbCrLf
    End If
    If diffOnly = "n" Then
        ToString = ToString & "StartDate: " & StartDate & vbCrLf
        ToString = ToString & "PreviousDate: " & PreviousDate & vbCrLf
        ToString = ToString & "CurrentDate: " & currentdate & vbCrLf
        ToString = ToString & "EndDate: " & EndDate & vbCrLf
        ToString = ToString & "StartTicks: " & startTicks & vbCrLf
        ToString = ToString & "PreviousTicks: " & previousTicks & vbCrLf
        ToString = ToString & "CurrentTicks: " & currentTicks & vbCrLf
        ToString = ToString & "EndTicks: " & endTicks & vbCrLf
    End If
    ToString = ToString & "CurrentTicks - StartTicks: " & (currentTicks - startTicks) & vbCrLf
    ToString = ToString & "CurrentTicks - PreviousTicks: " & (currentTicks - previousTicks) & vbCrLf
End Function
Public Function WriteToStream(Optional diffOnly As String = "y", Optional updateCurrentDate As Boolean = True, Optional Description As String = "") As String
    WriteToStream = ToString(diffOnly, updateCurrentDate, Description)
    If NotIsNullOrEmpty(WriteToStream) Then
        ts.WriteLine WriteToStream
    End If
End Function

Public Sub Reset()
    Me.Enabled = False
    StartDate = 0
    PreviousDate = 0
    currentdate = 0
    EndDate = 0
    startTicks = 0
    previousTicks = 0
    currentTicks = 0
    endTicks = 0
    PreviousDescription = ""
    CurrentDescription = ""
End Sub

Private Sub Class_Terminate()
  ts.Close
  Set ts = Nothing
End Sub
