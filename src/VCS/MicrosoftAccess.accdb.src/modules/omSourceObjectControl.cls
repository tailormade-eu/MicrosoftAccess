VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omSourceObjectControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim rs As New ADODB.Recordset
Public Id As Long

Public Sub Load(soId As Long, c As omControl)

    rs.filter = "SourceObjectId=" & soId & " AND ControlId=" & c.Id
    If rs.EOF Then
        rs.AddNew
        rs("SourceObjectId") = soId
        rs("ControlId") = c.Id
        rs("ControlName") = c.Name
        rs("ControlTypeId") = c.ControlTypeId
        rs("ControlDefault") = c.Default
        rs("CreateDate") = Now
    End If
    rs("LastUsedDate") = Now
    rs.Update
    Id = rs("Id")
End Sub

Private Sub Class_Initialize()
    rs.Open "omSourceObjectControls", CurrentProject.connection, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Class_Terminate()
    rs.Close
    Set rs = Nothing
End Sub
