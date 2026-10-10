Attribute VB_Name = "tstTesseractOCR"
Option Compare Database
Option Explicit

' Tests of omTesseractOCR. Run in the immediate window:  ?tstTesseractOCR.RunAll
' Answer: "OK (n tests)", "FAILED: <what>" or "SKIPPED: <why>" (no Tesseract, or the sample image is missing).
' The sample image is tests\ocr-sample.png of the library (the text HELLO OCR 123); give its path, or put a
' copy named ocr-sample.png next to the database.
Public Function RunAll(Optional imagePath As String = "") As String
Dim text As String
Dim again As String
Dim missing As String
Dim fails As String
Dim n As Long

    If Len(FindTesseractExe) = 0 Then
        RunAll = "SKIPPED: Tesseract not found"
        Exit Function
    End If
    If Len(imagePath) = 0 Then imagePath = gFso.BuildPath(CurrentProject.Path, "ocr-sample.png")
    If Not gFso.FileExists(imagePath) Then
        RunAll = "SKIPPED: sample image not found: " & imagePath
        Exit Function
    End If

    text = GetTxtFromImage(imagePath)
    Check n, fails, "reads the word", InStr(1, text, "HELLO", vbTextCompare) > 0
    Check n, fails, "reads the number", InStr(1, text, "123") > 0

    again = GetTxtFromImage(imagePath)
    Check n, fails, "a second read gives the same text", Trim(again) = Trim(text)

    ' an image that cannot be read must give "", not the text of the previous image
    missing = GetTxtFromImage(imagePath & ".does-not-exist")
    Check n, fails, "an unreadable image gives no text", Len(Trim(missing)) = 0

    RunAll = IIf(Len(fails) = 0, "OK (" & n & " tests)", "FAILED:" & fails)
End Function

Private Sub Check(ByRef n As Long, ByRef fails As String, label As String, ok As Boolean)
    n = n + 1
    If Not ok Then fails = fails & " [" & label & "]"
End Sub
