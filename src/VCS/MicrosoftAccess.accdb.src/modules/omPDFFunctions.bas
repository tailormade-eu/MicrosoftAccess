Attribute VB_Name = "omPDFFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=2e5c2282

' Opens the report (visible), writes it to a PDF file and closes it. omReportFunctions.ExportToPDF does
' the same with more choices (hidden window, quality) but does not close a visible report, so the two
' are not interchangeable.
Public Sub SaveReportAsPDF(ReportName As String, filename As String, Optional whereCondition As String = "")
    DoCmd.OpenReport ReportName, acViewPreview, whereCondition:=whereCondition
    DoCmd.OutputTo acOutputReport, ReportName, acFormatPDF, filename, False
    DoCmd.Close acReport, ReportName
End Sub
