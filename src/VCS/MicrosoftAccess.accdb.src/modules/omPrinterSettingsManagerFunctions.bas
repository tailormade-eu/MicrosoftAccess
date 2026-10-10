Attribute VB_Name = "omPrinterSettingsManagerFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=db1d8002

Public omPSM As New omPrinterSettingsManager

Public Sub GetPrinterSettings()


  omPSM.GetSettings
  Debug.Print omPSM.DeviceName
  Debug.Print omPSM.PaperBin

End Sub
