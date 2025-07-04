Attribute VB_Name = "omZipFunctions"
Option Compare Database
Option Explicit

Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

Public Function ZipFolder(ByVal folderName As String, Optional filenameZip As String = "") As String
    Dim oApp As Object

    If IsNullOrEmpty(filenameZip) Then
        While Right(folderName, 1) = "\"
            folderName = Left(folderName, Len(folderName) - 1)
        Wend
        filenameZip = folderName & ".zip"
    End If

    'Create empty Zip File
    NewZip (filenameZip)

    Set oApp = CreateObject("Shell.Application")
    'Copy the files to the compressed folder
    oApp.NameSpace(CVar(filenameZip)).CopyHere oApp.NameSpace(CVar(folderName)).Items

    'Keep script waiting until Compressing is done
    On Error Resume Next
    Do Until oApp.NameSpace(CVar(filenameZip)).Items.Count = oApp.NameSpace(CVar(folderName)).Items.Count
       omKernalFunctions.Sleep 1000
    Loop
    On Error GoTo 0

    ZipFolder = filenameZip
End Function

Public Function UnzipToFolder(ByVal zipFileName As String, Optional fileNameFolder As String = "") As String
    Dim oApp As Object

    If IsNullOrEmpty(fileNameFolder) Then
        fileNameFolder = Replace(zipFileName, ".zip", "")
    End If
    omFileFunctions.CreateFolderPath fileNameFolder

    Set oApp = CreateObject("Shell.Application")
    oApp.NameSpace(CVar(fileNameFolder)).CopyHere oApp.NameSpace(CVar(zipFileName)).Items

    On Error Resume Next
    gFso.DeleteFolder Environ("Temp") & "\Temporary Directory*", True

    UnzipToFolder = fileNameFolder
End Function
