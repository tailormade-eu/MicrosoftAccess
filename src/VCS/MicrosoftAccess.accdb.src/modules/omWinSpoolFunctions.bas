Attribute VB_Name = "omWinSpoolFunctions"
Option Compare Database
Option Explicit

' Paper sizes of the default printer, via winspool.drv DeviceCapabilities.
' https://docs.microsoft.com/en-us/office/vba/access/concepts/printing/programmatically-retrieve-printer-capabilities

#If VBA7 Then
Private Declare PtrSafe Function DeviceCapabilities Lib "winspool.drv" _
    Alias "DeviceCapabilitiesA" (ByVal lpsDeviceName As String, _
    ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, _
    ByVal lpDevMode As LongPtr) As Long
#Else
Private Declare Function DeviceCapabilities Lib "winspool.drv" _
    Alias "DeviceCapabilitiesA" (ByVal lpsDeviceName As String, _
    ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, _
    ByVal lpDevMode As Long) As Long
#End If

Private Const DC_PAPERS = 2
Private Const DC_PAPERNAMES = 16
Private Const PAPERNAME_LENGTH = 64 ' bytes per paper name in the DC_PAPERNAMES buffer

' Fills names() and numbers() (both 1-based) with the paper names and paper numbers
' of the default printer. Returns the number of papers (0 = none, or no printer).
Private Function GetPapers(names() As String, numbers() As Integer) As Long
Dim deviceName As String
Dim devicePort As String
Dim count As Long
Dim buffer As String
Dim i As Long
Dim pos As Long
Dim paperName As String

    deviceName = Application.Printer.DeviceName
    devicePort = Application.Printer.Port

    count = DeviceCapabilities(deviceName, devicePort, DC_PAPERNAMES, ByVal vbNullString, 0)
    If count <= 0 Then
        Exit Function
    End If

    ReDim numbers(1 To count)
    ReDim names(1 To count)
    buffer = String(PAPERNAME_LENGTH * count, vbNullChar)
    count = DeviceCapabilities(deviceName, devicePort, DC_PAPERNAMES, ByVal buffer, 0)
    If count <= 0 Then
        Exit Function
    End If
    If DeviceCapabilities(deviceName, devicePort, DC_PAPERS, numbers(1), 0) <= 0 Then
        Exit Function
    End If

    For i = 1 To count
        paperName = Mid(buffer, PAPERNAME_LENGTH * (i - 1) + 1, PAPERNAME_LENGTH)
        pos = InStr(1, paperName, vbNullChar)
        If pos > 0 Then
            paperName = Left(paperName, pos - 1)
        End If
        names(i) = paperName
    Next i
    GetPapers = count
End Function

' One line per paper: number, tab, name. Empty string when there is no printer or no papers.
Public Function GetPaperListText() As String
Dim names() As String
Dim numbers() As Integer
Dim count As Long
Dim i As Long

    On Error GoTo ErrHandler
    count = GetPapers(names, numbers)
    For i = 1 To count
        GetPaperListText = GetPaperListText & numbers(i) & vbTab & names(i) & vbCrLf
    Next i
    Exit Function
ErrHandler:
    GetPaperListText = ""
End Function

Public Sub GetPaperList()
    MsgBox "Papers available for " & Application.Printer.DeviceName & vbCrLf & vbCrLf & GetPaperListText()
End Sub

' Paper number (DMPAPER_*) of the paper with this name on the default printer (not case
' sensitive). -1 when the name is not found or there is no printer.
Public Function GetPaperSize(paperName As String) As Long
Dim names() As String
Dim numbers() As Integer
Dim count As Long
Dim i As Long

    GetPaperSize = -1
    On Error GoTo ErrHandler
    count = GetPapers(names, numbers)
    For i = 1 To count
        If StrComp(names(i), paperName, vbTextCompare) = 0 Then
            GetPaperSize = numbers(i)
            Exit Function
        End If
    Next i
    Exit Function
ErrHandler:
    GetPaperSize = -1
End Function
