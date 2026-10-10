VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omSeqItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=47848f0d

Public Action As String
Public Step As String
Public Parameters As String
Public SourceContext As String
Public EnvironmentUserName As String
Public MachineName As String
Public ProcessId As String
Public ProcessName As String
Public ThreadId As String
Public ApplicationName As String
Public ApplicationVersion As String
Public ErrorMessage As String

Public Level As omSeqItemLevel '@l
Public MessageTemplate As String '@mt
Public Exception As String '@e


Public Function ToJson(Optional formatted As Boolean = False) As String
Dim s As String

    AddTemplate s, formatted, "@t", DtAccessToIso(Now, GetUtcDateTime())
    AddTemplate s, formatted, "@mt", MessageTemplate
    AddTemplate s, formatted, "@l", omSeqItemLevelToString(Level)

    AddTemplate s, formatted, "Action", Action
    AddTemplate s, formatted, "Step", Step
    AddTemplate s, formatted, "Parameters", Parameters
    AddTemplate s, formatted, "SourceContext", SourceContext
    AddTemplate s, formatted, "EnvironmentUserName", EnvironmentUserName
    AddTemplate s, formatted, "MachineName", MachineName
    AddTemplate s, formatted, "ProcessId", ProcessId
    AddTemplate s, formatted, "ProcessName", ProcessName
    AddTemplate s, formatted, "ThreadId", ThreadId
    AddTemplate s, formatted, "ApplicationName", ApplicationName
    AddTemplate s, formatted, "ApplicationVersion", ApplicationVersion
    AddTemplate s, formatted, "ErrorMessage", ErrorMessage

    ToJson = "{" & IIf(formatted, vbCrLf, "") & s & IIf(formatted, vbCrLf, "") & "}"
End Function

' Adds  "name": "value"  to the JSON text in json (a comma between the pairs, none after the last one).
' An empty value is skipped.
Private Sub AddTemplate(ByRef json As String, formatted As Boolean, Name As String, Value As String)
    If Trim(Nz(Value, "")) = "" Then Exit Sub
    If Len(json) > 0 Then json = json & "," & IIf(formatted, vbCrLf, "")
    json = json & Chr(34) & Name & Chr(34) & ": " & Chr(34) & JsonEscape(Value) & Chr(34)
End Sub

' Escapes backslash, double quote and the control characters that are not allowed in a JSON string.
Private Function JsonEscape(ByVal Value As String) As String
Dim i As Long
Dim c As Long
Dim result As String
Dim bs As String

    bs = Chr(92)
    Value = Replace(Value, bs, bs & bs)
    Value = Replace(Value, Chr(34), bs & Chr(34))
    Value = Replace(Value, vbCrLf, bs & "n")
    Value = Replace(Value, vbCr, bs & "n")
    Value = Replace(Value, vbLf, bs & "n")
    Value = Replace(Value, vbTab, bs & "t")
    For i = 1 To Len(Value)
        c = AscW(Mid(Value, i, 1))
        If c >= 0 And c < 32 Then
            result = result & bs & "u" & Right("0000" & Hex(c), 4)
        Else
            result = result & Mid(Value, i, 1)
        End If
    Next i
    JsonEscape = result
End Function


Private Sub Class_Initialize()
    Me.Level = NotDefinedSeqLevel
    Me.MessageTemplate = "{Action} {Step} - {Parameters}"
End Sub
