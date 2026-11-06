Attribute VB_Name = "tstClientSettings"
Option Compare Database
Option Explicit

' Tests for the client settings of omGlobalDeclarations (standard values, the override of the client, loading again after a reset).
' Run from the Immediate window:  ?tstClientSettings.RunAll
' Returns "OK (n tests)" or one line per failing case.
'
' This module plays the client: it defines ClientSettings_LoadOverride. Its mode decides what the override does, and it is 0
' (does nothing) except while a test runs, so that no other test or startup code sees a setting.

Public gOverrideMode As Long      ' 0 nothing, 1 sets UseLoginForm and gDevelopmentMode
Private mFailures As String
Private mTests As Long

Public Sub ClientSettings_LoadOverride()
    Select Case gOverrideMode
        Case 1
            omUserFunctions.UseLoginForm = True
            omGlobalDeclarations.gDevelopmentMode = True
    End Select
End Sub

' A procedure that Application.Run cannot call without an argument: Run gives the error back to the caller.
' (An error raised inside an override cannot be tested: Access shows its own dialog and waits for a click.)
Public Sub ClientSettings_TestNeedsArgument(ByVal x As Long)
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal what As String)
    mTests = mTests + 1
    If Not condition Then mFailures = mFailures & what & vbCrLf
End Sub

Public Function RunAll() As String
Dim keepWin As Boolean
Dim keepForm As Boolean

    mFailures = ""
    mTests = 0
    keepWin = omUserFunctions.UseWindowsAuthentication
    keepForm = omUserFunctions.UseLoginForm

    ' 1. the standard values, and an override that sets a value
    gOverrideMode = 1
    omUserFunctions.UseLoginForm = False
    omGlobalDeclarations.ClientSettingsReset
    omGlobalDeclarations.ClientSettings_Load quiet:=True
    Check omUserFunctions.UseLoginForm, "the override did not set UseLoginForm"
    Check omGlobalDeclarations.gDevelopmentMode, "the override did not set gDevelopmentMode"
    Check Len(omGlobalDeclarations.ClientSettingsError) = 0, "a good override reported an error: " & omGlobalDeclarations.ClientSettingsError

    ' 2. the standard value comes back when the override does not set it
    gOverrideMode = 0
    omGlobalDeclarations.ClientSettings_Load quiet:=True
    Check Not omGlobalDeclarations.gDevelopmentMode, "gDevelopmentMode is not the standard value (False) when the override leaves it alone"

    ' 3. a reset: the values are gone, the next EnsureClientSettings gets them back from the override
    gOverrideMode = 1
    omGlobalDeclarations.ClientSettingsReset
    omUserFunctions.UseLoginForm = False
    Check Not omGlobalDeclarations.gDevelopmentMode, "ClientSettingsReset did not clear gDevelopmentMode"
    omGlobalDeclarations.EnsureClientSettings
    Check omUserFunctions.UseLoginForm, "after a reset EnsureClientSettings did not load the override again"
    Check omGlobalDeclarations.gDevelopmentMode, "after a reset gDevelopmentMode did not come back"

    ' 4. EnsureClientSettings loads only once while nothing was reset
    omUserFunctions.UseLoginForm = False
    omGlobalDeclarations.EnsureClientSettings
    Check Not omUserFunctions.UseLoginForm, "EnsureClientSettings ran the override again without a reset"

    ' 5. an error that Application.Run gives back is kept and shown, not swallowed
    gOverrideMode = 0
    omGlobalDeclarations.ClientSettings_Load quiet:=True
    Check omGlobalDeclarations.RunClientOverride("ClientSettings_TestNeedsArgument", True), "a procedure that exists was reported as not found"
    Check InStr(omGlobalDeclarations.ClientSettingsError, "ClientSettings_TestNeedsArgument") > 0, "an error from Application.Run was swallowed: [" & omGlobalDeclarations.ClientSettingsError & "]"

    ' 6. a client without an override: procedure not found is not an error
    gOverrideMode = 0
    Check Not omGlobalDeclarations.RunClientOverride("NoSuchProcedure_ClientSettings_Test", True), "a missing procedure was reported as found"
    Check Len(omGlobalDeclarations.ClientSettingsError) = 0 Or InStr(omGlobalDeclarations.ClientSettingsError, "NoSuchProcedure") = 0, "a missing procedure was reported as an error"

    ' 7. the login has no standard: nothing set means not set
    omUserFunctions.UseWindowsAuthentication = False
    omUserFunctions.UseLoginForm = False
    gOverrideMode = 0
    omGlobalDeclarations.ClientSettingsReset
    omGlobalDeclarations.EnsureClientSettings
    Check Not omUserFunctions.LoginMethodIsSet, "the login method counts as set while nothing sets it"
    gOverrideMode = 1
    omGlobalDeclarations.ClientSettingsReset
    omUserFunctions.UseWindowsAuthentication = False
    Check omUserFunctions.LoginMethodIsSet, "exactly one login method set was not recognised"
    omUserFunctions.UseWindowsAuthentication = True
    Check Not omUserFunctions.LoginMethodIsSet, "both login methods set were not refused"

    ' leave nothing behind
    gOverrideMode = 0
    omUserFunctions.UseWindowsAuthentication = keepWin
    omUserFunctions.UseLoginForm = keepForm
    omGlobalDeclarations.ClientSettingsReset
    omGlobalDeclarations.ClientSettings_Load quiet:=True

    If Len(mFailures) = 0 Then
        RunAll = "OK (" & mTests & " tests)"
    Else
        RunAll = mFailures
    End If
End Function
