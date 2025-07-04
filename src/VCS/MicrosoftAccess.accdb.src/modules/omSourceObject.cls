VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omSourceObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim rs As New ADODB.Recordset
Public Id As Long
Public Name As String
Public ObjectTypeId As AcObjectType

Public Sub Load(obj As Object, Optional objType As AcObjectType = 0)

    Me.Name = obj.Name
    If objType = 0 Then
        objType = IIf(Left(TypeName(obj), 5) = "Form_", AcObjectType.acForm, AcObjectType.acReport)
    End If
    Me.ObjectTypeId = objType
    rs.filter = "ObjectTypeId=" & objType & " AND Name='" & obj.Name & "'"
    If rs.EOF Then
        rs.AddNew
        rs("ObjectTypeId") = objType
        rs("Name") = obj.Name
        rs("CreateDate") = Now
        rs.Update
    End If
    rs("LastUsedDate") = Now
    rs.Update
    Id = rs("Id")
End Sub

Private Sub Class_Initialize()
    rs.Open "omSourceObjects", CurrentProject.connection, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Class_Terminate()
    rs.Close
    Set rs = Nothing
End Sub
