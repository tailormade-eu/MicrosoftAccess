Attribute VB_Name = "omVATFunctions"
Option Compare Database
Option Explicit

' Belgian VAT numbers: "BE 0123.456.789" = prefix BE + 10 digits, the last 2 are a check number
' (97 - first 8 digits Mod 97).

Public Sub OpenBEVATNumberLookupWebsite(VATNumber As Variant)
    If omStringFunctions.IsNullOrEmpty(VATNumber) Then
        Exit Sub
    End If
    omFileFunctions.OpenUrl "https://kbopub.economie.fgov.be/kbopub/zoeknummerform.html?nummer=" & CleanBEVatNumber(VATNumber) & "&actionLu=Zoek"
End Sub

' Digits only: removes spaces, dots, commas and the BE prefix.
Public Function CleanBEVatNumber(VATNumber As Variant) As String
    CleanBEVatNumber = Nz(omStringFunctions.ReplaceCharPattern(VATNumber, " .,beBE", ""), "")
End Function

' True only when the number starts with BE (a number without prefix is not valid here),
' has 10 digits and passes the Mod 97 check.
Public Function IsValidBEVatNumber(VATNumber As Variant) As Boolean
Dim cleanVatNumber As String
Dim number As Long
Dim numberCheck As Long

    cleanVatNumber = CleanBEVatNumber(VATNumber)
    If Len(cleanVatNumber) <> 10 Or UCase(Left(Trim(Nz(VATNumber, "")), 2)) <> "BE" Then
        Exit Function
    End If
    If Not IsNumeric(cleanVatNumber) Then
        Exit Function
    End If
    number = Val(Left(cleanVatNumber, 8))
    numberCheck = Val(Right(cleanVatNumber, 2))
    IsValidBEVatNumber = ((number Mod 97) + numberCheck) = 97
End Function

' Returns the number in the given pattern (every digit in the pattern is an index 0-9 into the
' 10 digits); empty string when the number is not valid.
Public Function FormatBEVatNumber(VATNumber As Variant, Optional formatString As String = "BE 0123.456.789") As String
Dim cleanVatNumber As String
Dim i As Long
Dim c As String

    If omStringFunctions.IsNullOrEmpty(formatString) Then
        Exit Function
    End If
    If IsValidBEVatNumber(VATNumber) = False Then
        Exit Function
    End If
    cleanVatNumber = CleanBEVatNumber(VATNumber)

    For i = 1 To Len(formatString)
        c = Mid(formatString, i, 1)
        If IsNumeric(c) = False Then
            FormatBEVatNumber = FormatBEVatNumber & c
        Else
            FormatBEVatNumber = FormatBEVatNumber & Mid(cleanVatNumber, CLng(c) + 1, 1)
        End If
    Next i
End Function
