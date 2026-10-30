Attribute VB_Name = "omReportFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-30 sha=30e3c539

Public Sub ExportToPDF(ReportName As String, filename As String, Optional whereCondition As String, Optional windowMode As AcWindowMode = acWindowNormal, Optional outputQuality As AcExportQuality = AcExportQuality.acExportQualityPrint)

    DoCmd.OpenReport ReportName, acViewPreview, whereCondition:=whereCondition, windowMode:=IIf(windowMode = acDialog, acHidden, windowMode)
    DoCmd.OutputTo acOutputReport, ReportName, acFormatPDF, filename, outputQuality:=outputQuality

    If windowMode = acHidden Or windowMode = acDialog Then
        DoCmd.Close acReport, ReportName, acSaveNo
        If windowMode = acDialog Then
            DoCmd.OpenReport ReportName, acViewPreview, whereCondition:=whereCondition, windowMode:=windowMode
        End If
    End If
End Sub

' The number of seconds that the report needs to open in print preview (hidden, closed again afterwards); use it to compare before and after a change.
' whereCondition (optional): the filter, as in OpenReport. A report that cancels itself because it has no data raises error 2501.
Public Function MeasureReportOpenTime(ReportName As String, Optional whereCondition As String = "") As Double
Dim started As Double
Dim errNumber As Long
Dim errDescription As String

    started = Timer
    On Error GoTo Failed
    DoCmd.OpenReport ReportName, acViewPreview, , whereCondition, acHidden
    MeasureReportOpenTime = Timer - started
    If MeasureReportOpenTime < 0 Then MeasureReportOpenTime = MeasureReportOpenTime + 86400 'the clock passed midnight
    DoCmd.Close acReport, ReportName, acSaveNo
    Exit Function
Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If SysCmd(acSysCmdGetObjectState, acReport, ReportName) <> 0 Then DoCmd.Close acReport, ReportName, acSaveNo
    On Error GoTo 0
    Err.Raise errNumber, "MeasureReportOpenTime", errDescription
End Function

' Shows or hides the controls of a REPORT whose tag is exactly this text. A procedure with the same name
' exists in omFormFunctions (for forms, with a list of tags): call them as
' omReportFunctions.UpdateVisibleByTag and omFormFunctions.UpdateVisibleByTag.
Public Sub UpdateVisibleByTag(Report As Report, tag As String, state As Boolean)
Dim ctrl As Control

    For Each ctrl In Report.Controls
        If ctrl.tag = tag Then
            ctrl.visible = state
        End If
    Next
End Sub

' Sets the paper (by name, as the default printer lists it: see GetPaperSize in omWinSpoolFunctions) and the margins of a REPORT and saves the report.
' Nothing happens when the printer has no paper with that name. The margins are in the unit of the report (twips).
Public Sub SetPaperSize(ReportName As String, paperSize As String, Optional bottomMargin As Long = 0, Optional leftMargin As Long = 0, Optional rightMargin As Long = 0, Optional topMargin As Long = 0)
Dim rep As Report
Dim reportSizeId As Long

    reportSizeId = GetPaperSize(paperSize)
    If reportSizeId = -1 Then Exit Sub
    DoCmd.OpenReport ReportName, acViewDesign, windowMode:=acHidden
    Set rep = Reports(ReportName)
    With rep.Printer
        .paperSize = reportSizeId
        .bottomMargin = bottomMargin
        .leftMargin = leftMargin
        .rightMargin = rightMargin
        .topMargin = topMargin
    End With
    'Flip the default view twice: this makes Access keep the printer settings when the report is saved
    rep.DefaultView = IIf(rep.DefaultView = 0, 1, 0)
    rep.DefaultView = IIf(rep.DefaultView = 0, 1, 0)
    DoCmd.Close acReport, ReportName, acSaveYes
    Set rep = Nothing
End Sub
