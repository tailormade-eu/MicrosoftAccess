Attribute VB_Name = "omReportFunctions"
Option Compare Database
Option Explicit

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

Public Sub UpdateVisibleByTag(Report As Report, tag As String, state As Boolean)
Dim ctrl As Control

    For Each ctrl In Report.Controls
        If ctrl.tag = tag Then
            ctrl.visible = state
        End If
    Next
End Sub
