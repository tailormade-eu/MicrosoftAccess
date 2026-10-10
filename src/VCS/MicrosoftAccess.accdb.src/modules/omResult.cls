VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' A new object counts as failed until Ok is called.
Private m_Status As Boolean
Public SuccessReason As String
Public FailReason As String
Public ObjectValue As Object
Public Value As Variant
Public Sub Ok(Optional reason As String = "")
    m_Status = True
    SuccessReason = reason
End Sub
Public Sub Failed(Optional reason As String = "")
    m_Status = False
    FailReason = reason
End Sub
Public Function IsSuccess() As Boolean
    IsSuccess = (m_Status = True)
End Function
Public Function IsFailed() As Boolean
    IsFailed = (m_Status = False)
End Function
