VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omBaseClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim cmdLoad As New ADODB.Command


Public Id As Long
Public Code As Variant
Public Name As Variant

Public Sub Load(CountryId As Long)
Dim rs As ADODB.Recordset

    If Id <> CountryId Then
        cmdLoad.Parameters(0) = CountryId
        Set rs = cmdLoad.Execute
        If Not rs.EOF Then
            Id = CountryId
            Code = rs("Code")
            Name = rs("Name")
        Else
            Me.Clear
        End If
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub Class_Initialize()
    Me.Clear
    cmdLoad.commandText = "SELECT * FROM Countries WHERE Id=?"
    cmdLoad.ActiveConnection = CurrentProject.connection
    cmdLoad.Parameters.Refresh
End Sub
Public Sub Clear()
    Id = 0
    Code = Null
    Name = Null
End Sub
