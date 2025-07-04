Attribute VB_Name = "omObjectFunctions"
Option Compare Database
Option Explicit

Public Function IsNothing(obj As Object) As Boolean
    IsNothing = (obj Is Nothing)
End Function

Public Function NotIsNothing(obj As Object) As Boolean
    NotIsNothing = Not IsNothing(obj)
End Function

Public Sub SetAsNothing(obj As Object)
    Set obj = Nothing
End Sub
