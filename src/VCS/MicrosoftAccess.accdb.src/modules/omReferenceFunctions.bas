Attribute VB_Name = "omReferenceFunctions"
Option Compare Database
Option Explicit

Public Sub ListReferences()
Dim R As Reference
Dim n As String
Dim P As String

    On Error Resume Next

    For Each R In Application.References
        n = ""
        n = R.Name
        P = ""
        P = R.FullPath
        Debug.Print "omReferenceFunctions.AddReference " & Chr(34) & P & Chr(34) & "     ' -> " & n
    Next
End Sub
Public Sub AddReference(filename As String)
    RemoveReference filename
    Application.References.AddFromFile filename
End Sub

Public Function FindReference(Name As String) As Reference
Dim R As Reference

    Set R = Nothing
    For Each R In Application.References
        If InStr(1, R.Name, Name) > 0 Then
            Set FindReference = R
            Exit Function
        End If
        If InStr(1, R.FullPath, Name) > 0 Then
            Set FindReference = R
            Exit Function
        End If
    Next
End Function

Public Sub RemoveReference(filename As String)
Dim R As Reference

    Set R = FindReference(filename)
    If Not (R Is Nothing) Then
        Application.References.Remove R
    End If
End Sub

Public Sub AddVBIDEReference()
    AddReference "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
End Sub

Public Sub AddAdoDbReference()
    AddReference "C:\Program Files (x86)\Common Files\System\ado\msado28.tlb"
End Sub

Public Sub AddScriptingReference()
    AddReference "C:\Windows\SysWOW64\scrrun.dll"     ' -> Scripting
End Sub
