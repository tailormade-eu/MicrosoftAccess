Attribute VB_Name = "omPrinterSettingsManagerFunctions"
Option Compare Database
Option Explicit

Public omPSM As New omPrinterSettingsManager

Public Sub GetPrinterSettings()


  omPSM.GetSettings
  Debug.Print omPSM.DeviceName
  Debug.Print omPSM.PaperBin

End Sub
