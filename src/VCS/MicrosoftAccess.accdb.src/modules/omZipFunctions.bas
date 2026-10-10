Attribute VB_Name = "omZipFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=338a33c5
Private Const MAX_WAIT_SECONDS As Long = 120 ' longest wait for the background copy of the shell

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
    Dim waitStart As Single

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
    waitStart = Timer
    Do Until oApp.NameSpace(CVar(filenameZip)).Items.Count = oApp.NameSpace(CVar(folderName)).Items.Count
       omKernalAPIFunctions.Sleep 1000
       If Timer < waitStart Or Timer - waitStart > MAX_WAIT_SECONDS Then Exit Do ' locked file or error: do not wait forever
    Loop
    On Error GoTo 0

    ZipFolder = filenameZip
End Function

Public Function UnzipToFolder(ByVal zipFileName As String, Optional fileNameFolder As String = "") As String
    Dim oApp As Object
    Dim waitStart As Single

    If IsNullOrEmpty(fileNameFolder) Then
        fileNameFolder = Left(zipFileName, Len(zipFileName) - 4) ' remove the .zip at the end, not one inside a path
    End If
    omFileFunctions.CreateFolderPath fileNameFolder

    Set oApp = CreateObject("Shell.Application")
    oApp.NameSpace(CVar(fileNameFolder)).CopyHere oApp.NameSpace(CVar(zipFileName)).Items

    ' CopyHere works in the background: wait until all items are there before the temp folders are removed
    On Error Resume Next
    waitStart = Timer
    Do Until oApp.NameSpace(CVar(fileNameFolder)).Items.Count >= oApp.NameSpace(CVar(zipFileName)).Items.Count
        omKernalAPIFunctions.Sleep 250
        If Timer < waitStart Or Timer - waitStart > MAX_WAIT_SECONDS Then Exit Do
    Loop
    gFso.DeleteFolder Environ("Temp") & "\Temporary Directory*", True

    UnzipToFolder = fileNameFolder
End Function
