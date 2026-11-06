Attribute VB_Name = "omGlobalDeclarations"
Option Compare Database
Option Explicit
'@lib 2026-11-06 sha=b722bc0b

' Settings of the client project. The library holds the standard value; a client changes it in its own procedure
' ClientSettings_LoadOverride (a Public Sub without arguments in a standard module of the client), for example:
'     Public Sub ClientSettings_LoadOverride()
'         gDevelopmentMode = False
'         omUserFunctions.UseWindowsAuthentication = True
'         omUserFunctions.UseLoginForm = False
'         omSeqFunctions.SeqServerUrl = "http://myserver:5341/"
'     End Sub
' A module that reads a setting calls omGlobalDeclarations.EnsureClientSettings first. After a reset of the VBA project
' (an unhandled error, End) the variables are empty again; the next EnsureClientSettings loads the settings again.
' Never put a setting of a client in an om* module as a Const: the next copy of the library overwrites it.
Public gDevelopmentMode As Boolean
Public ClientSettingsError As String      ' text of the last error in the override, "" when there was none
Private mClientSettingsLoaded As Boolean

Public gDictionary As New Scripting.Dictionary

' Loads the settings when they are not loaded (yet, or not anymore after a reset).
Public Sub EnsureClientSettings()
    If Not mClientSettingsLoaded Then ClientSettings_Load
End Sub

' Standard values first, then the override of the client. quiet = no message box on an error in the override (tests).
Public Sub ClientSettings_Load(Optional quiet As Boolean = False)
    mClientSettingsLoaded = True      ' first: the override may ask for a setting itself
    gDevelopmentMode = False          ' standard value
    ClientSettingsError = ""
    RunClientOverride "ClientSettings_LoadOverride", quiet
End Sub

' Forgets what was loaded, as a reset does. The next EnsureClientSettings loads again.
Public Sub ClientSettingsReset()
    mClientSettingsLoaded = False
    gDevelopmentMode = False
End Sub

' Runs a procedure of the client by name. Returns False when the client has no such procedure (error 2517: that is
' fine, the override is optional). Any other error that Application.Run gives back (a procedure that needs an argument,
' a name that is not a procedure) is not swallowed: it is kept in ClientSettingsError and shown. An error that is raised
' INSIDE the override is not given back to this procedure: Access shows its own run-time error dialog for it (measured,
' also with error trapping on "Break on Unhandled Errors"). Keep the override to plain assignments.
Public Function RunClientOverride(procName As String, Optional quiet As Boolean = False) As Boolean
Dim errNumber As Long
Dim errText As String

    On Error Resume Next
    Application.Run procName
    errNumber = Err.Number
    errText = Err.Description
    On Error GoTo 0
    If errNumber = 2517 Then Exit Function
    RunClientOverride = True
    If errNumber <> 0 Then
        ClientSettingsError = procName & ": error " & errNumber & " " & errText
        If Not quiet Then MsgBox "Error in " & ClientSettingsError, vbExclamation, "Client settings"
    End If
End Function

'Public gDefaults As New omDefaults => Obsolete declared in omDefaultFunctions
'Public gSysDefaults As New omDefaults => Obsolete declared in omDefaultFunctions
'Public gUser As New omUser => Obsolete declared in omUserFunctions
'Public gFso As New Scripting.FileSystemObject => Obsolete declared in omFileFunctions
