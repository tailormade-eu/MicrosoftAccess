Attribute VB_Name = "omPDFFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=a181c183

Public Sub SaveReportAsPDF(ReportName As String, filename As String, Optional whereCondition As String = "")
    DoCmd.OpenReport ReportName, acViewPreview, whereCondition:=whereCondition
    DoCmd.OutputTo acOutputReport, ReportName, acFormatPDF, filename, False
    DoCmd.Close acReport, ReportName
End Sub
