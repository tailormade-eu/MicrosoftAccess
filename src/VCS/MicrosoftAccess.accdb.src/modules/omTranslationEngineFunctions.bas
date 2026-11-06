Attribute VB_Name = "omTranslationEngineFunctions"
Option Compare Database
Option Explicit
'@lib 2026-11-06 sha=1cd8532a

Global omTE As New omTranslationEngine

' The language id that the translate code in forms and reports uses (omTE.Translate Me, GetCurrentLanguage()).
' Set it at startup and whenever the user changes language:  omTranslationEngineFunctions.CurrentLanguageId = 2
' 0 (not set) means language 1.
Public CurrentLanguageId As Long

Public Function GetCurrentLanguage() As Long
    omGlobalDeclarations.EnsureClientSettings
    GetCurrentLanguage =IIf(CurrentLanguageId > 0, CurrentLanguageId, 1)
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
