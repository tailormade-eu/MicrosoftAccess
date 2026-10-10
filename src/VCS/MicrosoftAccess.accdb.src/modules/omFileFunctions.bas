Attribute VB_Name = "omFileFunctions"
'@lib 2026-10-10 sha=c4a450fd
' Last updated by Raoul Jacobs on 20130619_1115

Option Compare Database
Option Explicit
' Enum to define window state options
Public Enum WindowState
    wsHide = 0
    wsNormal = 1
    wsMinimized = 2
    wsMaximized = 3
End Enum

' Private variable to store WScript.Shell instance
Private wsh As Object
Public gFso As New Scripting.FileSystemObject

Public Function BuildPathFileExists(path As String, filename As String) As String
    BuildPathFileExists = gFso.BuildPath(path, filename)
    If (Not gFso.FileExists(BuildPathFileExists)) Then
        BuildPathFileExists = ""
    End If
End Function
Public Function FolderExists(path As String) As Boolean
    FolderExists = gFso.FolderExists(path)
End Function
Public Function FileExists(path As String, filename As String) As String
    FileExists = gFso.BuildPath(path, filename)
    If (Not gFso.FileExists(FileExists)) Then
        FileExists = ""
    End If
End Function
Public Function OpenPathAndFile(AppPath As String, filename As String) As String
        OpenPathAndFile = FileExists(AppPath, filename)
        If OpenPathAndFile <> "" Then
            Shell "explorer.exe " & OpenPathAndFile, vbMaximizedFocus
        End If
End Function

Public Sub OpenFile(filename As String)
        If gFso.FileExists(filename) Then
            Shell "explorer.exe " & filename, vbMaximizedFocus
        End If
End Sub
Public Sub OpenUrl(url As String, Optional newTab As Boolean = True)
    If Not newTab Then
        Shell "explorer.exe " & Chr(34) & url & Chr(34), vbMaximizedFocus
        Exit Sub
    End If
    If wsh Is Nothing Then Set wsh = CreateObject("WScript.Shell")
    Dim progId As String
    On Error Resume Next
    progId = wsh.RegRead("HKCU\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice\ProgId")
    On Error GoTo 0
    If InStr(LCase(progId), "chrome") > 0 Then
        wsh.Run "chrome.exe --new-tab """ & url & """"
    ElseIf InStr(LCase(progId), "edge") > 0 Or InStr(LCase(progId), "msedge") > 0 Then
        wsh.Run "msedge.exe --new-tab """ & url & """"
    ElseIf InStr(LCase(progId), "firefox") > 0 Then
        wsh.Run "firefox.exe -new-tab """ & url & """"
    Else
        Shell "explorer.exe " & Chr(34) & url & Chr(34), vbMaximizedFocus
    End If
End Sub

Public Sub OpenFolder(path As String, Optional createPath As Boolean = False)
        If createPath Then
            CreateFolderPath path
        End If
        If gFso.FolderExists(path) Then
            Shell "explorer.exe " & path, vbMaximizedFocus
        End If
End Sub

Public Sub CreateDatebase(filename As String)
    If gFso.FileExists(filename) Then
        gFso.DeleteFile filename
    End If
    DBEngine.CreateDataBase filename, dbLangGeneral, dbVersion30
End Sub
Public Sub DeleteFile(filename As String, Optional force As Boolean = False)
    If gFso.FileExists(filename) = False Then
                Exit Sub
        End If
        gFso.DeleteFile filename, force
End Sub
Public Sub DeleteFolder(folderName As String, Optional force As Boolean = False)
    If omFileFunctions.FolderExists(folderName) Then
        gFso.DeleteFolder folderName, force
    End If
End Sub

Public Sub RenameFile(strOldName As String, strNewName As String, Optional strNewExtension As String = "")
Dim strBackupName As String
Dim lBackupCount As Long

    On Error GoTo FileRename_Error

    If IsNullOrEmpty(strOldName) Then
        GoTo FileRename_Exit
    End If
    If NotIsNullOrEmpty(strNewExtension) Then
        strNewName = Left(strOldName, InStrRev(strOldName, ".")) & strNewExtension
    End If
    If Not IsNull(strNewName) Then
        lBackupCount = 1
FileRename_Backup:
        strBackupName = Left(strNewName, Len(strNewName) - 2) & format(lBackupCount, "0#")
        Name strNewName As strBackupName
FileRename_Rename:
        Name strOldName As strNewName
    End If
FileRename_Exit:
    Exit Sub

FileRename_Error:
    Select Case Err
        Case 53
            Resume FileRename_Rename
        Case 58
            lBackupCount = lBackupCount + 1
            Resume FileRename_Backup
        Case Else
            MsgBox "Error"
            'DoCmd.SetWarnings True
            'ErrorBox.Module = "Module : File Functions"
            'ErrorBox.Procedure = "Procedure : FileRename"
            'ErrorBox.Number = "Error : " & Err
            'ErrorBox.Description = Error
            'ErrorBox.Execute
            'If gbErrorResume = True Then
            '    If gbDebugMode = True Then
            '        Stop
            '        Resume
            '    Else
            '        Resume Next
            '    End If
            'Else
            '    DoCmd.Quit acQuitSaveAll
            'End If
    End Select

End Sub
Public Function ExtractPath(strFileName As String) As String
Dim lPos As Long

    lPos = InStr(1, strFileName, "\")
    If lPos = 0 Then
        ExtractPath = ""
        Exit Function
    End If
    While InStr(lPos + 1, strFileName, "\") <> 0
        lPos = InStr(lPos + 1, strFileName, "\")
    Wend
    ExtractPath = Left(strFileName, lPos)
End Function
Public Function RemoveExtension(Name As String) As String
Dim lPos As Long
    lPos = InStrRev(Name, ".")
    If lPos > 0 Then
        RemoveExtension = Left(Name, lPos - 1)
    Else
        RemoveExtension = Name
    End If
End Function

Public Function GetExtension(filename As String) As String
Dim lPos As Long
    lPos = InStrRev(filename, ".")
    If lPos > 0 Then
        GetExtension = Right(filename, Len(filename) - lPos)
    End If
End Function
Public Function MakeFilenameSafe(filename As String) As String
    MakeFilenameSafe = Replace(Replace(Replace(Replace(Replace(Trim(filename), "/", "_"), ",", " "), Chr(34), " "), "'", " "), ":", "_")
End Function

Public Sub CreateFolderPath(strPath As String, Optional OpenFolderInExplorer As Boolean = False)

    If Not gFso.FolderExists(strPath) Then
        If Not gFso.FolderExists(gFso.GetParentFolderName(strPath)) Then
            CreateFolderPath (gFso.GetParentFolderName(strPath))
        End If
        gFso.createFolder strPath
    End If
    If OpenFolderInExplorer Then
        omFileFunctions.OpenFolder strPath
    End If
End Sub

Public Sub ReadFilesIntoTable(folderPath As String, tableName As String, Optional pattern As String = "")
Dim rs As New ADODB.Recordset
Dim F As Scripting.File

    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM [" & tableName & "]"
    DoCmd.SetWarnings True
    rs.Open tableName, CurrentProject.connection, adOpenForwardOnly, adLockOptimistic
    For Each F In gFso.GetFolder(folderPath).Files
        If InStr(1, F.Name, pattern) <> 0 And InStr(1, F.Name, "~") = 0 Then
            rs.AddNew
            rs("Name") = F.Name
            rs("Filename") = F.path
            rs("Extension") = omFileFunctions.GetExtension(F.Name)
            rs.Update
        End If
    Next
    rs.Close
    Set rs = Nothing

End Sub
Public Function FileLocked(strFileName As String, Optional displayError As Boolean = False) As Boolean
   On Error Resume Next
   ' If the file is already opened by another process,
   ' and the specified type of access is not allowed,
   ' the Open operation fails and an error occurs.
   Open strFileName For Binary Access Read Write Lock Read Write As #1
   Close #1
   ' If an error occurs, the document is currently open.
   If Err.number <> 0 Then
      ' Display the error number and description.
      If displayError Then
        MsgBox "Error #" & str(Err.number) & " - " & Err.Description
      End If
      FileLocked = True
      Err.Clear
   End If
End Function

Public Function GetDesktopFolder() As String
    GetDesktopFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop")
End Function
Public Function GetUserRootFolder() As String
    GetUserRootFolder = Environ("USERPROFILE")
End Function
Public Function GetTempFilenameInTempPath() As String
    GetTempFilenameInTempPath = gFso.BuildPath(gFso.GetSpecialFolder(2), gFso.GetTempName)
End Function
Public Function GetTempPathInTempPath(Optional autoCreate As Boolean = True) As String
Dim pos As Long
    GetTempPathInTempPath = GetTempFilenameInTempPath
    pos = InStrRev(GetTempPathInTempPath, ".")
    If pos > 0 Then
        GetTempPathInTempPath = Left(GetTempPathInTempPath, pos - 1)
    End If
    If autoCreate Then
        omFileFunctions.CreateFolderPath GetTempPathInTempPath
    End If
End Function
Public Function CopyFileToFolder(source As String, path As String, Optional overWriteFiles As Boolean = False) As String

    CopyFileToFolder = gFso.BuildPath(path, gFso.GetFileName(source))
    gFso.CopyFile source, CopyFileToFolder, overWriteFiles

End Function
' Copies a file to the temp folder (same file name, overwrites) and returns the path of the copy.
' Returns "" when a copy with that name is already there and locked by another program.
Public Function CopyToLocalTemp(filename As String) As String
Dim tmpFilename As String

    tmpFilename = gFso.BuildPath(gFso.GetSpecialFolder(2), gFso.GetFileName(filename))
    If gFso.FileExists(tmpFilename) Then
        If FileLocked(tmpFilename) Then
            Exit Function
        End If
    End If
    gFso.CopyFile filename, tmpFilename, True
    CopyToLocalTemp = tmpFilename
End Function
Public Function ReadFileToString(path As String) As String
Dim ts As TextStream

    Set ts = gFso.OpenTextFile(path, ForReading, False)
    ReadFileToString = ts.ReadAll
    ts.Close
    Set ts = Nothing
End Function

Public Sub PrintFile(filename As String)
    CreateObject("Shell.Application").NameSpace(0).ParseName(filename).InvokeVerb ("Print")
End Sub

Public Sub WriteStringToFile(filename As String, s As String, Optional unicode As Boolean = False)
Dim ts As TextStream

    Set ts = gFso.CreateTextFile(filename, True, unicode)
    ts.Write s
    ts.Close
    Set ts = Nothing
End Sub

' Function to execute an application with configurable options
Public Sub ExecuteShell(filename As String, Optional workingDir As String = "", Optional winState As WindowState = wsNormal, Optional wait As Boolean = True)
    ' Initialize WScript.Shell if not already created
    If wsh Is Nothing Then Set wsh = CreateObject("WScript.Shell")

    ' Determine working directory
    If workingDir = "" Then
        ' Extract directory from file path if none is provided
        workingDir = gFso.GetParentFolderName(filename)
    End If
    If gFso.FolderExists(workingDir) Then
        wsh.CurrentDirectory = workingDir
    End If

    ' Run the executable with specified parameters
    wsh.Run Chr(34) & filename & Chr(34), winState, wait
End Sub

' Starts an exe with arguments, waits until it finishes and returns the exit code.
' -1 = Access could not start the exe. ExecuteShell above does not return the exit code.
Public Function ExecuteShellExitCode(filename As String, Optional args As String = "", Optional winState As WindowState = wsHide) As Long
    On Error GoTo Fout
    If wsh Is Nothing Then Set wsh = CreateObject("WScript.Shell")
    If Not gFso.FileExists(filename) Then
        ExecuteShellExitCode = -1
        Exit Function
    End If
    wsh.CurrentDirectory = gFso.GetParentFolderName(filename)
    ExecuteShellExitCode = wsh.Run(Chr(34) & filename & Chr(34) & IIf(args = "", "", " " & args), winState, True)
    Exit Function
Fout:
    ExecuteShellExitCode = -1
End Function
