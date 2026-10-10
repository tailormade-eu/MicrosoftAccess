Attribute VB_Name = "omExcelVBAFunctions"
Option Explicit
'@lib 2026-10-10 sha=058187aa

' Small helpers for VBA in Excel: Nz (Excel has none), text with a local number notation to a number, and a formula in local notation to the notation VBA wants.

' Value, or ValueIfNull when Value is Null or Empty.
Public Function Nz(Value As Variant, Optional ValueIfNull As Variant = "") As Variant
    If IsNull(Value) Or IsEmpty(Value) Then
        Nz = ValueIfNull
    Else
        Nz = Value
    End If
End Function

' "1.234,56" or "1,234.56" to a number. The separator that comes last is the decimal separator.
' With only one kind of separator, that one is the decimal separator ("1.234" is 1.234 and "1,234" is 1.234 as well).
' Returns the error value #VALUE! (CVErr(xlErrValue)) when the text is not a number.
Public Function ConvertLocalizedNumber(ByVal inputStr As String) As Variant
Dim decimalSep As String
Dim thousandSep As String
Dim dotPos As Long
Dim commaPos As Long
Dim cleaned As String
Dim sysSep As String
Dim d As Double

    dotPos = InStrRev(inputStr, ".")
    commaPos = InStrRev(inputStr, ",")
    If dotPos > commaPos Then
        decimalSep = "."
        thousandSep = ","
    ElseIf commaPos > dotPos Then
        decimalSep = ","
        thousandSep = "."
    Else
        decimalSep = Application.DecimalSeparator
        thousandSep = IIf(decimalSep = ".", ",", ".")
    End If
    sysSep = Application.DecimalSeparator
    cleaned = Replace(inputStr, thousandSep, "")
    cleaned = Replace(cleaned, decimalSep, sysSep)
    If IsNumeric(cleaned) Then
        d = CDbl(cleaned)
        ConvertLocalizedNumber = d
    Else
        ConvertLocalizedNumber = CVErr(xlErrValue)
    End If
End Function

' A formula typed in the local notation of Excel (for example =SOM(1,5;A1) with decimal comma and ; between arguments)
' to the notation that the VBA property .Formula wants (=SUM(1.5,A1)). Text between quotes is left alone.
' Only separators are converted; function names are not translated (use .FormulaLocal for that).
Public Function ConvertLocalizedFormula(ByVal inputStr As String) As String
    ConvertLocalizedFormula = ConvertFormulaSeparators(inputStr, Application.International(xlDecimalSeparator), Application.International(xlListSeparator))
End Function

' The conversion of ConvertLocalizedFormula for a given decimal separator and list separator.
Public Function ConvertFormulaSeparators(ByVal formula As String, decimalSep As String, listSep As String) As String
Dim i As Long
Dim ch As String
Dim inText As Boolean
Dim result As String

    If Len(formula) <= 1 Or Left(formula, 1) <> "=" Then
        ConvertFormulaSeparators = formula
        Exit Function
    End If
    For i = 1 To Len(formula)
        ch = Mid(formula, i, 1)
        If ch = """" Then
            inText = Not inText
        ElseIf Not inText Then
            If ch = decimalSep Then
                ch = "."
            ElseIf ch = listSep Then
                ch = ","
            End If
        End If
        result = result & ch
    Next i
    ConvertFormulaSeparators = result
End Function
