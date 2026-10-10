Attribute VB_Name = "omLockingFunctions"
Option Explicit
'@lib 2026-10-10 sha=9b11e39e

' A lock file next to a project file, so that two people do not work in it at the same time.
' For C:\data\Project.xlsm the lock file is C:\data\Project.lock; it holds two lines: the user name and the time.
' The lock is a file, so it only works on a folder that both people can reach.

Private Const LOCK_FILE_EXTENSION As String = ".lock"

Private Function Fso() As Object
    Set Fso = CreateObject("Scripting.FileSystemObject")
End Function

' Creates the lock file. Returns True when this call created it, False when a lock file already exists (also your own)
' or when it could not be created. A lock file without a valid user name and time counts as corrupted and is replaced.
' The file is created in one step (it fails when it exists), so two people cannot both succeed.
Public Function TryCreateProjectLock(projectFilename As String) As Boolean
Dim lockPath As String
Dim f As Object

    lockPath = GetLockFilename(projectFilename)
    If Len(lockPath) = 0 Then Exit Function
    CleanupCorruptedLock projectFilename
    On Error GoTo Failed
    Set f = Fso().CreateTextFile(lockPath, False)
    f.WriteLine Environ("Username")
    f.WriteLine Format(Now, "yyyy-mm-dd hh:nn:ss")
    f.Close
    TryCreateProjectLock = True
    Exit Function
Failed:
    TryCreateProjectLock = False
End Function

' Same as TryCreateProjectLock, with messages for the user: an invalid file name, an existing lock (and who has it).
Public Sub CreateProjectLock(projectFilename As String, Optional showSuccessMessage As Boolean = False)
    If Len(GetLockFilename(projectFilename)) = 0 Then
        MsgBox "Cannot create lock file: invalid path for " & projectFilename, vbExclamation
    ElseIf TryCreateProjectLock(projectFilename) Then
        If showSuccessMessage Then MsgBox "Lock file created for project.", vbInformation
    ElseIf IsProjectLocked(projectFilename) Then
        MsgBox "Project is already locked. " & GetLockInfo(projectFilename), vbExclamation
    Else
        MsgBox "Error creating lock file for " & projectFilename, vbCritical
    End If
End Sub

Public Function IsProjectLocked(projectFilename As String) As Boolean
Dim lockPath As String

    lockPath = GetLockFilename(projectFilename)
    If Len(lockPath) > 0 Then IsProjectLocked = Fso().FileExists(lockPath)
End Function

' "Locked by <user> since <time>", or "No lock file exists."
Public Function GetLockInfo(projectFilename As String) As String
Dim user As String
Dim timestamp As Date

    If Not IsProjectLocked(projectFilename) Then
        GetLockInfo = "No lock file exists."
        Exit Function
    End If
    ReadLockFileData GetLockFilename(projectFilename), user, timestamp
    GetLockInfo = "Locked by " & user & " since " & IIf(timestamp = 0, "(unknown)", Format(timestamp, "yyyy-mm-dd hh:nn:ss"))
End Function

' The user name in the lock file, or "" when there is no lock.
Public Function GetLockUser(projectFilename As String) As String
Dim user As String
Dim timestamp As Date

    If Not IsProjectLocked(projectFilename) Then Exit Function
    ReadLockFileData GetLockFilename(projectFilename), user, timestamp
    GetLockUser = user
End Function

Public Function IsLockedByCurrentUser(projectFilename As String) As Boolean
    If IsProjectLocked(projectFilename) Then IsLockedByCurrentUser = (GetLockUser(projectFilename) = Environ("Username"))
End Function

' When the lock is yours from an earlier session: asks whether to take it over (the time in the file is renewed).
' Returns True when you now hold the lock.
Public Function HandleSelfLock(projectFilename As String) As Boolean
    If Not IsLockedByCurrentUser(projectFilename) Then Exit Function
    If MsgBox("You have an existing lock on this project from a previous session." & vbCrLf & GetLockInfo(projectFilename) & vbCrLf & vbCrLf & _
              "Do you want to reclaim this lock and continue editing?", vbYesNo + vbQuestion, "Existing Lock Found") = vbYes Then
        UnlockProject projectFilename
        HandleSelfLock = TryCreateProjectLock(projectFilename)
    End If
End Function

' True when there is no lock file, or when it holds a user name and a valid time.
Public Function ValidateLockFile(projectFilename As String) As Boolean
Dim user As String
Dim timestamp As Date
Dim lockPath As String

    If Not IsProjectLocked(projectFilename) Then
        ValidateLockFile = True
        Exit Function
    End If
    lockPath = GetLockFilename(projectFilename)
    ReadLockFileData lockPath, user, timestamp
    ValidateLockFile = (Len(Trim(user)) > 0 And timestamp <> 0)
End Function

' Deletes the lock file when it is corrupted (see ValidateLockFile).
Public Sub CleanupCorruptedLock(projectFilename As String)
    If ValidateLockFile(projectFilename) Then Exit Sub
    On Error Resume Next
    Kill GetLockFilename(projectFilename)
End Sub

' Deletes the lock file (no check of who owns it).
Public Sub UnlockProject(projectFilename As String)
Dim lockPath As String

    If Not IsProjectLocked(projectFilename) Then Exit Sub
    lockPath = GetLockFilename(projectFilename)
    On Error GoTo Failed
    Kill lockPath
    Exit Sub
Failed:
    MsgBox "Error removing lock file: " & Err.Description, vbCritical
End Sub

' Deletes the lock file of someone else. askConfirmation: show who has the lock and ask first (default True).
Public Sub ForceUnlockProject(projectFilename As String, Optional askConfirmation As Boolean = True)
    If Not IsProjectLocked(projectFilename) Then Exit Sub
    If askConfirmation Then
        If MsgBox("ADMINISTRATIVE UNLOCK" & vbCrLf & vbCrLf & "Project: " & Fso().GetBaseName(projectFilename) & vbCrLf & GetLockInfo(projectFilename) & vbCrLf & vbCrLf & _
                  "Are you sure you want to force unlock this project?", vbYesNo + vbExclamation, "Force Unlock Project") <> vbYes Then Exit Sub
    End If
    UnlockProject projectFilename
End Sub

' The lock files (full path) in one folder.
Public Function GetLockFilesInDirectory(directoryPath As String) As Collection
Dim result As New Collection
Dim f As Object

    If Fso().FolderExists(directoryPath) Then
        For Each f In Fso().GetFolder(directoryPath).Files
            If LCase(Right(f.Name, Len(LOCK_FILE_EXTENSION))) = LOCK_FILE_EXTENSION Then result.Add f.Path
        Next f
    End If
    Set GetLockFilesInDirectory = result
End Function

' The lock files (full path) in a folder and all folders under it. Folders that cannot be read are skipped.
Public Function GetAllLockFiles(rootPath As String) As Collection
Dim result As New Collection

    If Len(rootPath) > 0 Then
        If Fso().FolderExists(rootPath) Then CollectLockFiles Fso().GetFolder(rootPath), result
    End If
    Set GetAllLockFiles = result
End Function

Private Sub CollectLockFiles(folder As Object, result As Collection)
Dim f As Object
Dim subFolder As Object

    On Error Resume Next
    For Each f In folder.Files
        If LCase(Right(f.Name, Len(LOCK_FILE_EXTENSION))) = LOCK_FILE_EXTENSION Then result.Add f.Path
    Next f
    For Each subFolder In folder.SubFolders
        CollectLockFiles subFolder, result
    Next subFolder
End Sub

' The user name and the time from a lock file. user = "" and timestamp = 0 for a missing, empty or unreadable file.
Public Sub ReadLockFileData(lockPath As String, ByRef user As String, ByRef timestamp As Date)
Dim f As Object
Dim text As String

    user = ""
    timestamp = 0
    If Len(lockPath) = 0 Then Exit Sub
    If Not Fso().FileExists(lockPath) Then Exit Sub
    On Error GoTo Done
    Set f = Fso().OpenTextFile(lockPath, 1)
    If Not f.AtEndOfStream Then user = f.ReadLine
    If Not f.AtEndOfStream Then
        text = f.ReadLine
        If IsDate(text) Then timestamp = CDate(text)
    End If
Done:
    If Not f Is Nothing Then f.Close
End Sub

' The full name of the lock file for a project file, or "" when the folder does not exist or the name is empty.
Private Function GetLockFilename(projectFilename As String) As String
Dim folder As String
Dim nameOnly As String

    On Error GoTo Failed
    folder = Fso().GetParentFolderName(projectFilename)
    If Len(folder) = 0 Then Exit Function
    If Not Fso().FolderExists(folder) Then Exit Function
    nameOnly = Fso().GetBaseName(projectFilename)
    If Len(Trim(nameOnly)) = 0 Then Exit Function
    GetLockFilename = folder & "\" & nameOnly & LOCK_FILE_EXTENSION
    Exit Function
Failed:
    GetLockFilename = ""
End Function
