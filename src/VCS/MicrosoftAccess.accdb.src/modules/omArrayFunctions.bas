﻿Attribute VB_Name = "omArrayFunctions"
Option Compare Database
Option Explicit

Public Sub ObjectArrayAdd(objectArray() As Object, Value As Object, Optional Clear = False)
    On Error Resume Next
    ReDim Preserve objectArray(UBound(objectArray) + 1) As Object
    If Err = 9 Or Clear Then ReDim objectArray(0) As Object
    Set objectArray(UBound(objectArray)) = Value
End Sub


Public Sub StringArrayAdd(stringArray() As String, Value As String, Optional Clear = False)
    On Error Resume Next
    ReDim Preserve stringArray(UBound(stringArray) + 1) As String
    If Err = 9 Or Clear Then ReDim stringArray(0) As String
    stringArray(UBound(stringArray)) = Value
End Sub
Public Sub StringArrayClear(stringArray() As String)
Dim emptyArray() As String
    stringArray = emptyArray
End Sub
Public Sub TestAddString()
Dim strings() As String
Dim newStrings() As String
    StringArrayAdd strings, "hallo"
    StringArrayAdd newStrings, "Best"
    StringArrayAddArray strings, newStrings
End Sub
Public Sub StringArrayAddArray(stringArray() As String, valueArray() As String)
Dim i As Integer

    On Error Resume Next
    i = UBound(valueArray)
    If Err = 0 Then
        For i = 0 To UBound(valueArray)
            StringArrayAdd stringArray, valueArray(i)
        Next
    End If

End Sub
Public Function StringArrayCount(stringArray() As String) As Long
    StringArrayCount = 0
    On Error Resume Next
    StringArrayCount = UBound(stringArray) + 1
End Function
Public Function StringArrayFind(stringArray() As String, Value As String, includeContains As Boolean) As Long
Dim i As Long
Dim Length As Long
Dim containsArray() As String

    Length = StringArrayLength(stringArray)
    If Length = 0 Then
        Exit Function
    End If
    StringArrayFind = -1
    If includeContains Then
        While i < Length
            If InStr(1, stringArray(i), Value) > 0 Then
                StringArrayFind = i
                Exit Function
            End If
            i = i + 1
        Wend
    Else
        While i < Length
            If stringArray(i) = Value Then
                StringArrayFind = i
                Exit Function
            End If
            i = i + 1
        Wend
    End If
End Function
Public Function StringArrayContains(stringArray() As String, Value As String) As Boolean

    StringArrayContains = (UBound(filter(stringArray, Value, True)) > -1)

End Function
Public Function StringArrayDoesNotContain(stringArray() As String, Value As String) As Boolean

    StringArrayDoesNotContain = Not StringArrayContains(stringArray, Value)

End Function

Public Function StringArrayLength(stringArray() As String) As Long
    On Error Resume Next
    StringArrayLength = UBound(stringArray) + 1
End Function

Public Function StringArrayToString(stringArray() As String, Optional delimiter As String = vbCrLf, Optional skipNullBlankSpaces As Boolean = False) As String
Dim i As Long

    For i = 0 To StringArrayLength(stringArray) - 1
        If Not skipNullBlankSpaces Or (skipNullBlankSpaces And NotIsNullOrEmpty(stringArray(i))) Then
            StringArrayToString = StringArrayToString & stringArray(i) & delimiter
        End If
    Next i
End Function

Public Function StringArrayRemoveEmpty(stringArray() As String) As String()
Dim i As Long

    For i = 0 To StringArrayLength(stringArray) - 1
        If Len(stringArray(i)) > 0 Then
            StringArrayAdd StringArrayRemoveEmpty, stringArray(i)
        End If
    Next i
End Function
