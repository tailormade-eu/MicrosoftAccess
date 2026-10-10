Attribute VB_Name = "omTranslationEngineFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=321ecc26

Global omTE As New omTranslationEngine

Public Function GetCurrentLanguage() As Long
  GetCurrentLanguage = Nz(Forms("frmLicentie").Form.txtUserTaalId, 1)
End Function

Public Sub UpdateTranslations()
  omTE.IndexAll
End Sub

Public Sub ClearTranslations()
    omTE.ClearAll
End Sub
Public Sub InjectCodeInFormsAndReports()
  omTE.InsertTranslateCode "omTE"
End Sub
