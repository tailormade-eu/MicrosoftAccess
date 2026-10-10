Attribute VB_Name = "omSeqFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=c580750b

Const seqUrl = "http://seq.domainmmvii.local:5341/"

Public Enum omSeqItemLevel
    NotDefinedSeqLevel = 0
    DebugSeqLevel = 1
    WarningSeqLevel = 2
    ErrorSeqLevel = 3
End Enum

Public Function omSeqItemLevelToString(Value As omSeqItemLevel) As String
    Select Case Value
        Case omSeqItemLevel.DebugSeqLevel
            omSeqItemLevelToString = "Debug"
        Case omSeqItemLevel.WarningSeqLevel
            omSeqItemLevelToString = "Warning"
            Case omSeqItemLevel.ErrorSeqLevel
            omSeqItemLevelToString = "Error"
    End Select
End Function

Public Sub SendSeqItem(Optional seqItem As omSeqItem) 'item As omSeqItem)
Dim result As String

    If omObjectFunctions.IsNothing(seqItem) Then
        Set seqItem = New omSeqItem
        With seqItem
            .ApplicationName = "MSAccess-Test"
            .Level = ErrorSeqLevel
            .Action = "omSeqFunctions.SendSeqItem"
            .Step = "self test"
        End With
    End If
    result = HttpPostText(seqUrl & "api/events/raw?clef", seqItem.ToJson())
    seqItem.Step = ""
    seqItem.Level = DebugSeqLevel
    seqItem.Exception = ""
    seqItem.Parameters = ""
End Sub

' POST a text body (JSON) to an address and return the answer as text. The error of the request
' (unknown host, no connection) is raised to the caller.
Public Function HttpPostText(url As String, body As String) As String
Dim xml As Object

    Set xml = CreateObject("MSXML2.XMLHTTP")
    xml.Open "POST", url, False
    xml.setRequestHeader "Content-Type", "application/json"
    xml.setRequestHeader "Accept", "application/json"
    xml.Send body
    HttpPostText = xml.ResponseText
End Function

Public Sub OpenSeq(Optional filter As String = "")
    omFileFunctions.OpenUrl seqUrl & IIf(NotIsNullOrEmpty(filter), "#/events?filter=" & filter, "")
End Sub

Public Function GetUtcDateTime()
Dim dt As Object
    Set dt = CreateObject("WbemScripting.SWbemDateTime")
    dt.SetVarDate Now
    GetUtcDateTime = dt.GetVarDate(False)
    Set dt = Nothing
End Function

'ISO to Access
Public Function DtIsoToAccess(myisodate As String) As Date
DtIsoToAccess = CDate(Replace(myisodate, "T", " "))
End Function

'Access to ISO
Public Function DtAccessToIso(myaccdate As Date, Optional myUtcDate As Date = 0) As String
Dim dtDiff As Long
    myUtcDate = IIf(myUtcDate = 0, myaccdate, myUtcDate)
    dtDiff = DateDiff("n", myUtcDate, myaccdate) ' offset to UTC in minutes
    DtAccessToIso = format(myaccdate, "yyyy-mm-dd\Thh:nn:ss.000") & IIf(dtDiff = 0, "z", IIf(dtDiff < 0, "-", "+") & format(Int(Abs(dtDiff) / 60), "00") & ":" & format(Abs(dtDiff) Mod 60, "00"))
End Function
