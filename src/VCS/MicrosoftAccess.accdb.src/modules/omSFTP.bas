Attribute VB_Name = "omSFTP"
Option Compare Database
Option Explicit

' Upload a file over SFTP by starting an existing command line client. Nothing is stored in the
' database: the caller passes host, user and password/key (keep those in a secret store, not in code).
'
' Engines:
' - WinSCP (WinSCP.com, https://winscp.net): password or private key (.ppk/OpenSSH key).
'   Looked up in Program Files, %LOCALAPPDATA%\Programs and next to the database, or pass winscpPath.
' - OpenSSH sftp.exe (built into Windows 10/11): private key only (sftp has no password option).
'   Used when no password is given and WinSCP is not found.
'
' Host key: WinSCP checks the fingerprint in hostKey ("ssh-ed25519 256 xxxx..."); empty = trust the
' key on the first connection (WinSCP 6 or newer, "acceptnew"). OpenSSH uses known_hosts and trusts
' a new host on the first connection.
'
' A key with a passphrase is not supported (use a key without one, or Pageant with WinSCP).
' The password is only written to a temporary script file that is deleted right after the upload.
' It never appears on a command line and never in the error text that is returned.

' Uploads localFile to remotePath (a remote folder ending in "/" keeps the file name).
' Returns "" when the upload worked, otherwise the reason.
Public Function SftpPut(host As String, user As String, localFile As String, remotePath As String, _
                        Optional password As String = "", Optional privateKeyFile As String = "", _
                        Optional hostKey As String = "", Optional port As Long = 22, _
                        Optional winscpPath As String = "") As String
Dim engine As String
Dim exe As String
Dim scriptFile As String
Dim logFile As String
Dim args As String
Dim exitCode As Long

    If Len(host) = 0 Or Len(user) = 0 Then
        SftpPut = "host and user are required"
        Exit Function
    End If
    If Not gFso.FileExists(localFile) Then
        SftpPut = "local file not found: " & localFile
        Exit Function
    End If
    If Len(password) = 0 And Len(privateKeyFile) = 0 Then
        SftpPut = "give a password or a private key file"
        Exit Function
    End If
    If Len(privateKeyFile) > 0 Then
        If Not gFso.FileExists(privateKeyFile) Then
            SftpPut = "private key file not found: " & privateKeyFile
            Exit Function
        End If
    End If

    exe = SftpFindWinScp(winscpPath)
    If Len(exe) > 0 Then
        engine = "winscp"
    ElseIf Len(password) = 0 Then
        exe = Environ("SystemRoot") & "\System32\OpenSSH\sftp.exe"
        If gFso.FileExists(exe) Then
            engine = "openssh"
        End If
    End If
    If Len(engine) = 0 Then
        SftpPut = IIf(Len(password) > 0, "a password needs WinSCP (WinSCP.com not found; pass winscpPath)", "neither WinSCP nor OpenSSH sftp.exe found")
        Exit Function
    End If

    scriptFile = omFileFunctions.GetTempFilenameInTempPath
    logFile = omFileFunctions.GetTempFilenameInTempPath
    On Error GoTo Fout
    If engine = "winscp" Then
        WriteTextFile scriptFile, SftpBuildWinScpScript(host, user, localFile, remotePath, password, privateKeyFile, hostKey, port)
        args = "/ini=nul /script=" & Quote(scriptFile) & " /log=" & Quote(logFile)
        exitCode = omFileFunctions.ExecuteShellExitCode(exe, args)
    Else
        WriteTextFile scriptFile, "put " & Quote(localFile) & " " & Quote(remotePath) & vbCrLf
        args = Quote(exe) & " -b " & Quote(scriptFile) & " -i " & Quote(privateKeyFile) & " -P " & port & _
               " -o BatchMode=yes -o StrictHostKeyChecking=accept-new " & Quote(user & "@" & host) & _
               " > " & Quote(logFile) & " 2>&1"
        ' cmd /c wants ONE pair of outer quotes around the whole command line
        exitCode = omFileFunctions.ExecuteShellExitCode(Environ("ComSpec"), "/c " & Chr(34) & args & Chr(34))
    End If

    If exitCode <> 0 Then
        SftpPut = engine & " failed (exit code " & exitCode & "): " & LastLogLines(logFile, 5, password)
    End If
    GoTo Opruimen
Fout:
    SftpPut = Err.Description
Opruimen:
    On Error Resume Next
    gFso.DeleteFile scriptFile, True
    gFso.DeleteFile logFile, True
End Function

' Full path of WinSCP.com, or "" when it is not found.
Public Function SftpFindWinScp(Optional winscpPath As String = "") As String
Dim candidates As Variant
Dim i As Long

    If Len(winscpPath) > 0 Then
        If gFso.FileExists(winscpPath) Then
            SftpFindWinScp = winscpPath
        End If
        Exit Function
    End If
    candidates = Array(Environ("ProgramFiles") & "\WinSCP\WinSCP.com", _
                       Environ("ProgramFiles(x86)") & "\WinSCP\WinSCP.com", _
                       Environ("LOCALAPPDATA") & "\Programs\WinSCP\WinSCP.com", _
                       CurrentProject.Path & "\WinSCP\WinSCP.com")
    For i = LBound(candidates) To UBound(candidates)
        If gFso.FileExists(candidates(i)) Then
            SftpFindWinScp = candidates(i)
            Exit Function
        End If
    Next i
End Function

' The WinSCP script text (public so it can be tested; it contains the password).
Public Function SftpBuildWinScpScript(host As String, user As String, localFile As String, remotePath As String, _
                                      password As String, privateKeyFile As String, hostKey As String, port As Long) As String
Dim openCmd As String

    openCmd = "open sftp://" & host & ":" & port & "/ -username=" & Quote(user)
    If Len(password) > 0 Then
        openCmd = openCmd & " -password=" & Quote(password)
    End If
    If Len(privateKeyFile) > 0 Then
        openCmd = openCmd & " -privatekey=" & Quote(privateKeyFile)
    End If
    openCmd = openCmd & " -hostkey=" & Quote(IIf(Len(hostKey) > 0, hostKey, "acceptnew"))

    SftpBuildWinScpScript = "option batch abort" & vbCrLf & _
                            "option confirm off" & vbCrLf & _
                            openCmd & vbCrLf & _
                            "put " & Quote(localFile) & " " & Quote(remotePath) & vbCrLf & _
                            "exit" & vbCrLf
End Function

' Text in double quotes; a double quote inside the text is doubled (WinSCP and cmd.exe style).
Private Function Quote(s As String) As String
    Quote = Chr(34) & Replace(s, Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Private Sub WriteTextFile(filename As String, s As String)
Dim ts As Object

    Set ts = gFso.CreateTextFile(filename, True, False)
    ts.Write s
    ts.Close
End Sub

' The last lines of a log file with the password masked; "(no log)" when there is nothing to read.
Private Function LastLogLines(filename As String, lineCount As Long, password As String) As String
Dim ts As Object
Dim lines() As String
Dim text As String
Dim i As Long
Dim n As Long

    If Not gFso.FileExists(filename) Then
        LastLogLines = "(no log)"
        Exit Function
    End If
    Set ts = gFso.OpenTextFile(filename, 1, False)
    If Not ts.AtEndOfStream Then
        text = ts.ReadAll
    End If
    ts.Close
    If Len(password) > 0 Then
        text = Replace(text, password, "***")
    End If
    lines = Split(Replace(text, vbCr, ""), vbLf)
    For i = UBound(lines) To LBound(lines) Step -1
        If Len(Trim(lines(i))) > 0 Then
            LastLogLines = lines(i) & IIf(Len(LastLogLines) > 0, " | ", "") & LastLogLines
            n = n + 1
            If n >= lineCount Then Exit For
        End If
    Next i
    If Len(LastLogLines) = 0 Then
        LastLogLines = "(empty log)"
    End If
End Function
