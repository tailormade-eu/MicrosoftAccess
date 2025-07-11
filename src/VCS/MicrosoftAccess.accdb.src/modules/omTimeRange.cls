﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omTimeRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_StartTime As Date
Private m_EndTime As Date

Public Property Get StartTime() As Date
    StartTime = m_StartTime
End Property

Public Property Let StartTime(ByVal v As Date)
    m_StartTime = TimeValue(v)
End Property


Public Property Get EndTime() As Date
    EndTime = m_EndTime
End Property

Public Property Let EndTime(ByVal v As Date)
    m_EndTime = TimeValue(v)
End Property
Public Function GetStartTimeInSeconds() As Long
    GetStartTimeInSeconds = omDateFunctions.GetTimeInSeconds(Me.StartTime)
End Function

Public Function GetEndTimeInSeconds() As Long
    GetEndTimeInSeconds = omDateFunctions.GetTimeInSeconds(Me.EndTime, True)
End Function
