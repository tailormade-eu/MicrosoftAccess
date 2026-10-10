Attribute VB_Name = "omTesseractOCR"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=19c837a0

'reference Windows Script Host Object Model

' Full path of tesseract.exe, or "" when it is not installed. Looks in c:\tesseract-ocr (where this module
' always looked), in Program Files and in %LOCALAPPDATA%\Programs.
Public Function FindTesseractExe() As String
Dim candidates As Variant
Dim i As Long

    candidates = Array("c:\tesseract-ocr\tesseract.exe", _
                       Environ("ProgramFiles") & "\Tesseract-OCR\tesseract.exe", _
                       Environ("ProgramFiles(x86)") & "\Tesseract-OCR\tesseract.exe", _
                       Environ("LOCALAPPDATA") & "\Programs\Tesseract-OCR\tesseract.exe")
    For i = LBound(candidates) To UBound(candidates)
        If gFso.FileExists(candidates(i)) Then
            FindTesseractExe = candidates(i)
            Exit Function
        End If
    Next i
End Function

' The text Tesseract reads in an image (png, jpg, bmp, tif). Returns "" when the image cannot be read or
' Tesseract is not installed. The result file lives in the temp folder and is removed again, so a failed
' read can never return the text of an earlier image.
Public Function GetTxtFromImage(filename As String) As String
Dim O As New IWshRuntimeLibrary.WshShell
Dim exe As String
Dim cmd As String
Dim tmpBase As String
Dim tmpFile As String
Dim ts As Scripting.TextStream

    exe = FindTesseractExe
    If Len(exe) = 0 Then Exit Function
    tmpBase = omFileFunctions.GetTempFilenameInTempPath ' Tesseract adds .txt
    tmpFile = tmpBase & ".txt"
    cmd = Chr(34) & exe & Chr(34) & " " & Chr(34) & filename & Chr(34) & " " & Chr(34) & tmpBase & Chr(34)
    O.Run cmd, vbMinimizedNoFocus, 1
    If gFso.FileExists(tmpFile) Then
        Set ts = gFso.OpenTextFile(tmpFile, ForReading)
        If Not ts.AtEndOfStream Then GetTxtFromImage = ts.ReadAll
        ts.Close
        Set ts = Nothing
        gFso.DeleteFile tmpFile, True
    End If
    Set O = Nothing
End Function
