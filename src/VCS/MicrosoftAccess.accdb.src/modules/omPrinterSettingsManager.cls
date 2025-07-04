VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omPrinterSettingsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public DeviceName As String
Public PaperBin As Long

Public Sub GetSettings()
Dim ReportName As String
Dim rpt As Report

  ReportName = "omPrinterSettings"

  DoCmd.OpenReport ReportName, acViewPreview, windowMode:=acDialog
  'DoCmd.OpenReport reportName, acViewPreview

  'Set rpt = Reports(reportName)
  'DeviceName = rpt.Printer.DeviceName
  'PaperBin = rpt.Printer.PaperBin

  'DoCmd.Close acReport, reportName

End Sub
