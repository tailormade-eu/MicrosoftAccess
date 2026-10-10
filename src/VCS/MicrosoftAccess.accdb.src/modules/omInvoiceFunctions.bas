Attribute VB_Name = "omInvoiceFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=015f119d

Public Function GetStructuredPaymentReference(number As Long, Optional alignLeft As Boolean = True) As String
Dim s As String
Dim rest As Long


    s = number
    If Len(s) < 10 And alignLeft Then
        s = Left(s & "0000000000", 10)
    End If
    rest = CLng(s) Mod 97
    If rest = 0 Then
        rest = 97
    End If
    If Len(s) < 10 Then
        s = Right("0000000000" & s, 10)
    End If
    s = s & Right("00" & rest, 2)
    GetStructuredPaymentReference = "+++" & Left(s, 3) & "/" & Mid(s, 4, 4) & "/" & Mid(s, 9) & "+++"

End Function
