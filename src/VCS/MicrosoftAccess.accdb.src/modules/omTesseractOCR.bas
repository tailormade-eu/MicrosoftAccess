Attribute VB_Name = "omTesseractOCR"
Option Compare Database
Option Explicit

'reference Windows Script Host Object Model

Const ocrExe = "c:\tesseract-ocr\tesseract.exe"
Const ocrPath = "c:\tesseract-ocr\"
Public Function GetTxtFromImage(filename As String) As String
Dim O As New IWshRuntimeLibrary.WshShell
Dim cmd As String
Dim tmpFile As String
Dim ts As Scripting.TextStream

    tmpFile = ocrPath & "ocrtemp"
    cmd = Chr(34) & ocrExe & Chr(34) & " " & Chr(34) & filename & Chr(34) & " " & Chr(34) & tmpFile & Chr(34)
    O.Run cmd, vbMinimizedNoFocus, 1
    tmpFile = tmpFile + ".txt"
    If gFso.FileExists(tmpFile) Then
        Set ts = gFso.OpenTextFile(tmpFile, ForReading)
        GetTxtFromImage = ts.ReadAll
        ts.Close
        Set ts = Nothing
    End If
    Set O = Nothing
End Function

Public Sub GetTxtFromImageTest()
    Debug.Print GetTxtFromImage("C:\_d\Mvx.bmp")
End Sub
