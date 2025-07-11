﻿Attribute VB_Name = "omStringFunctions"
Option Compare Database
Option Explicit

Public Enum StringPaddingMode
    PadLeft = 1
    PadRight = 2
End Enum


' Last Updated by Raoul Jacobs on 20130617_1305
'Public Function Nz(Value As Variant, Optional valueifnull = "")
'    Nz = IIf(IsNull(Value), valueifnull, Value)
'End Function
Public Function CleanString(Value As Variant) As String
    If IsNull(Value) Or IsEmpty(Value) Then
        CleanString = ""
    Else
        CleanString = Trim(Value)
    End If
End Function

Public Function IsNullOrEmpty(Value As Variant) As Boolean
    IsNullOrEmpty = (Len(CleanString(Value)) = 0)
End Function

Public Function IsNullOrEmptyOrZero(Value As Variant) As Boolean
    IsNullOrEmptyOrZero = IsNullOrEmpty(Value)
    If Not IsNullOrEmptyOrZero Then
        IsNullOrEmptyOrZero = (Value = 0)
        If Not IsNullOrEmptyOrZero And IsDate(Value) Then
            IsNullOrEmptyOrZero = (CDbl(CDate(Value)) = 0)
        End If
    End If
End Function

Public Function NotIsNullOrEmptyOrZero(Value As Variant) As Boolean
    NotIsNullOrEmptyOrZero = Not IsNullOrEmptyOrZero(Value)
End Function
Public Function NotIsNullOrEmpty(Value As Variant) As Boolean
    NotIsNullOrEmpty = Not IsNullOrEmpty(Value)
End Function

Public Function ParseValue(arguments As Variant, valueName As String, Optional assignChar = "=", Optional splitchar = ",", Optional notFoundValue As Variant = Null) As Variant
Dim lPos As Long
Dim lPosTerminator
Dim lPosEqual As Long

    ParseValue = notFoundValue
    If IsNullOrEmpty(arguments) Then
        Exit Function
    End If
    lPos = InStr(1, Nz(arguments, ""), valueName & assignChar)
    If lPos <> 0 Then
        lPosEqual = InStr(lPos, arguments, assignChar)
        If lPosEqual <> 0 Then
            lPosTerminator = InStr(lPosEqual, arguments, splitchar)
            ParseValue = Mid(arguments, lPosEqual + 1, IIf(lPosTerminator = 0, Len(arguments), lPosTerminator - 1) - lPosEqual)
        End If
    End If
End Function


Public Function StringSplit(Value As String, splitchar As String, Optional removeEmptyValues As Boolean = True) As String()
Dim startPos As Long
Dim endPos As Long
Dim result() As String
Dim Length As Long
Dim tempValue As String
    If Len(Value) > 0 Then
        omArrayFunctions.StringArrayClear result
        startPos = 1
        While startPos <> 0
            endPos = InStr(startPos, Value, splitchar)
            Length = IIf(endPos = 0, Len(Value) + 1, endPos) - startPos
            If (removeEmptyValues = False) Or (removeEmptyValues And Length > 0) Then
                tempValue = Mid(Value, startPos, Length)
                omArrayFunctions.StringArrayAdd result, tempValue
            End If
            startPos = IIf(endPos <> 0, endPos + 1, 0)
        Wend
    End If
    StringSplit = result
End Function

Public Function StringSplitGetByIndex(Value As String, splitchar As String, index As Long) As Variant
    StringSplitGetByIndex = Split(Value, splitchar)(index)
End Function
Public Function IsIdNullOrZero(Value As Variant) As Boolean
    IsIdNullOrZero = (Nz(Value, 0) = 0)
End Function
Public Function NotIsIdNullOrZero(Value As Variant) As Boolean
    NotIsIdNullOrZero = Not IsIdNullOrZero(Value)
End Function
Public Function IsIdNullOrEmptyOrZero(Value As Variant) As Boolean
    If IsNullOrEmpty(Value) Then
        Value = 0
    End If
    IsIdNullOrEmptyOrZero = (Nz(Value, 0) = 0)
End Function


Public Function NotIsIdNullOrEmptyOrZero(Value As Variant) As Boolean
    NotIsIdNullOrEmptyOrZero = Not IsIdNullOrEmptyOrZero(Value)
End Function

Public Function AreIdsEqual(id1 As Variant, id2 As Variant) As Boolean
    If omStringFunctions.IsIdNullOrZero(id1) And omStringFunctions.IsIdNullOrZero(id2) Then AreIdsEqual = True
    AreIdsEqual = Nz((id1 = id2), False)
End Function
Public Function AreStringsEqual(string1 As Variant, string2 As Variant) As Boolean
    If omStringFunctions.IsNullOrEmpty(string1) And omStringFunctions.IsNullOrEmpty(string2) Then AreStringsEqual = True
    AreStringsEqual = Nz((string1 = string2), False)
End Function

Public Function StringFormat(source As String, replace0 As String, Optional replace1 As Variant = Null, Optional replace2 As Variant = Null, Optional replace3 As Variant = Null, Optional replace4 As Variant = Null, Optional replace5 As Variant = Null, Optional replace6 As Variant = Null, Optional replace7 As Variant = Null, Optional replace8 As Variant = Null, Optional replace9 As Variant = Null, Optional replace10 As Variant = Null, Optional replace11 As Variant = Null, Optional replace12 As Variant = Null) As String
    StringFormat = Replace(source, "{0}", replace0)
    StringFormat = Replace(StringFormat, "{1}", Nz(replace1, ""))
    StringFormat = Replace(StringFormat, "{2}", Nz(replace2, ""))
    StringFormat = Replace(StringFormat, "{3}", Nz(replace3, ""))
    StringFormat = Replace(StringFormat, "{4}", Nz(replace4, ""))
    StringFormat = Replace(StringFormat, "{5}", Nz(replace5, ""))
    StringFormat = Replace(StringFormat, "{6}", Nz(replace6, ""))
    StringFormat = Replace(StringFormat, "{7}", Nz(replace7, ""))
    StringFormat = Replace(StringFormat, "{8}", Nz(replace8, ""))
    StringFormat = Replace(StringFormat, "{9}", Nz(replace9, ""))
    StringFormat = Replace(StringFormat, "{10}", Nz(replace10, ""))
    StringFormat = Replace(StringFormat, "{11}", Nz(replace11, ""))
    StringFormat = Replace(StringFormat, "{12}", Nz(replace12, ""))
End Function
Public Function RemoveChars(source As Variant, findPattern As String, replaceString As String)
Dim returnString As String
Dim i As Long

    returnString = Nz(source, "")
    For i = 1 To Len(findPattern)
        returnString = Replace(returnString, Mid(findPattern, i, 1), replaceString)
    Next i
    RemoveChars = returnString

End Function

Public Function KeepChars(source As Variant, keepPattern As String, replaceString As String)
Dim returnString As String
Dim i As Long
Dim currentChar As String

    returnString = Nz(source, "")
    i = 1
    While i <= Len(returnString)
        currentChar = Mid(returnString, i, 1)
        If InStr(1, keepPattern, currentChar) = 0 Then
            returnString = Replace(returnString, currentChar, replaceString)
        Else
            i = i + 1
        End If
    Wend
    KeepChars = returnString
End Function


Public Function ContainsString(source As Variant, containString As String, Optional useDelimiter As String = "") As Long
    ContainsString = InStr(1, useDelimiter & source & useDelimiter, useDelimiter & containString & useDelimiter)

End Function

Public Function GetEnglishPlural(Value As String) As String

    If Right(Value, 1) = "y" Then
        GetEnglishPlural = Left(Value, Len(Value) - 1) & "ies"
    ElseIf Right(Value, 1) = "s" Then
        GetEnglishPlural = Value & "es"
    Else
        GetEnglishPlural = Value & "s"
    End If
End Function

Public Function ReplaceCharPattern(source As Variant, findPattern As String, Replacement As String) As Variant
Dim returnString As String
Dim i As Long

    If IsNullOrEmpty(source) Then
        ReplaceCharPattern = Null
    Else
        returnString = source
        For i = 1 To Len(findPattern)
            returnString = replaceString(returnString, Mid(findPattern, i, 1), Replacement)
        Next i
        ReplaceCharPattern = returnString
    End If
End Function
Public Function CleanCommunication(Value As Variant) As Variant
    If IsNullOrEmpty(Value) Then
        CleanCommunication = Null
    Else
        CleanCommunication = replaceString(replaceString(replaceString(replaceString(replaceString(replaceString(replaceString(replaceString(replaceString(replaceString(replaceString(Value, " ", ""), "+", ""), ".", ""), ",", ""), "-", ""), "@", ""), ";", ""), "*", ""), "/", ""), "(", ""), ")", "")
    End If
End Function

Public Function ReverseCommunication(Value As Variant) As Variant
    Value = CleanCommunication(Value)
    If IsNullOrEmpty(Value) Then
        ReverseCommunication = Null
    Else
        ReverseCommunication = StrReverse(Value)
    End If
End Function

Function Proper(var As Variant) As Variant
' Purpose: Convert the case of var so that the first letter of each word capitalized.
   Dim strV As String, intChar As Integer, i As Integer
   Dim fWasSpace As Integer    'Flag: was previous char a space?

   If IsNull(var) Then Exit Function
   strV = var
   fWasSpace = True              'Initialize to capitalize first letter.
   For i = 1 To Len(strV)
      intChar = Asc(Mid$(strV, i, 1))
      Select Case intChar
      Case 65 To 90              ' A to Z
         If Not fWasSpace Then Mid$(strV, i, 1) = Chr$(intChar Or &H20)
      Case 97 To 122             ' a to z
         If fWasSpace Then Mid$(strV, i, 1) = Chr$(intChar And &HDF)
      End Select
      fWasSpace = (intChar = 32)
   Next
   Proper = strV
End Function


Public Function replaceString(ByVal Value As String, oldString As String, newString As String) As String
    While InStr(1, Value, oldString) > 0
        Value = Replace(Value, oldString, newString)
    Wend
    replaceString = Value
End Function
Public Function GetvbCrLf() As String
    GetvbCrLf = vbCrLf
End Function

Public Function GetDelimitedValue(Value As String, Optional Position As Long = 0, Optional delimiter As String = ";", Optional EmbraceChar As String)
Dim strings() As String

    strings = omStringFunctions.StringSplit(Value, delimiter & EmbraceChar, False)
    GetDelimitedValue = Replace(strings(Position), EmbraceChar, "")
End Function
Public Function GetScrambleNumbers(Length As Long) As String
Dim strTemp As String
Dim i As Long
    Randomize
    For i = 1 To Length
        strTemp = strTemp & CInt(Rnd(1) * 9)
    Next i
    GetScrambleNumbers = strTemp
End Function

Public Function IsStringInPattern(Value As String, pattern As String) As Boolean
    IsStringInPattern = (InStr(1, pattern, Value) > 0)
End Function

Public Function CleanStringUsingPattern(ByVal Value As Variant, Optional findPattern As String = vbCrLf & vbTab, Optional replaceString As String = " ", Optional replaceDoubleBlanks As Boolean = True, Optional trimResult As Boolean = True) As String
    If IsNull(Value) Then
        CleanStringUsingPattern = ""
        Exit Function
    End If

    Value = ReplaceCharPattern(Value, findPattern, replaceString)
    If replaceDoubleBlanks Then
        While InStr(1, Value, "  ") > 0
            Value = Replace(Value, "  ", " ")
        Wend
    End If
    If trimResult Then
        CleanStringUsingPattern = CleanString(Value)
    Else
        CleanStringUsingPattern = Value
    End If
End Function
Public Function StringPadLeft(s As String, totalLength As Long, paddingChar As String) As String
    StringPadLeft = StringPad(s, totalLength, paddingChar, PadLeft)
End Function
Public Function StringPadRight(s As String, totalLength As Long, paddingChar As String) As String
    StringPadRight = StringPad(s, totalLength, paddingChar, PadRight)
End Function

Public Function StringPad(s As String, totalLength As Long, paddingChar As String, paddingMode As StringPaddingMode) As String
Dim result As String

    result = strings.String(totalLength, paddingChar)
    If paddingMode = StringPaddingMode.PadLeft Then
        result = result & Nz(s, "")
        result = Right(result, totalLength)
    Else
        result = Nz(s, "") & result
        result = Left(result, totalLength)
    End If
    StringPad = result
End Function

Public Function ContainsCharFromPattern(s As String, pattern As String) As Boolean
Dim c As String
Dim i As Long
                For i = 1 To Len(pattern)
                        c = Mid(pattern, i, 1)
                        If InStr(1, s, c) > 0 Then
                                ContainsCharFromPattern = True
                                Exit Function
                        End If
                Next i
End Function

Public Function ExtractLong(text As Variant, findtext As String) As Long
Dim pos As Long
Dim posStart As Long

    text = Nz(text, "")
    While InStr(1, text, "  ") > 0
        text = Replace(text, "  ", " ")
    Wend
    pos = InStr(1, text, " " & findtext)
    If pos = 0 Then
        pos = InStr(1, text, findtext)
    End If
    If pos <> 0 Then
        posStart = InStrRev(text, " ", pos - 1) + 1
        'ExtractLong = CLng(Replace(Replace(Trim(Mid(text, posStart, pos - posStart)), ".", ""), ",", ""))
    End If
End Function

Public Function GetSubstringAfterChar(s As String, Optional c As String = ".", Optional autoTrim As Boolean = True) As String
    Dim pos As Integer
    Dim result As String

    ' Check if the input string is null or empty
    If omStringFunctions.IsNullOrEmpty(s) Then
        result = s
    Else
        ' Find the position of the last occurrence of the character
        pos = InStrRev(s, c)

        If pos > 0 Then
            ' Extract the substring after the character
            result = Mid(s, pos + Len(c))
        Else
            ' If the character is not found, return the original string
            result = s
        End If
    End If

    ' Trim the result if autoTrim is True
    If autoTrim Then
        result = Trim(result)
    End If

    GetSubstringAfterChar = result
End Function

Public Function ByteArrayToHexString(ba() As Byte) As String
Dim i As Long
Dim s As String

    For i = 0 To UBound(ba)
        s = s & Right("00" & Hex$(ba(i)), 2)
    Next
    ByteArrayToHexString = "0x" & s
End Function

Public Function ByteArrayToString(ba() As Byte) As String
Dim i As Long
Dim s As Double

    For i = 0 To UBound(ba)
        s = s * 256 + ba(i)
    Next
    ByteArrayToString = s
End Function

Public Function CheckIllegalCharInString(strCheck As String) As String
    'Check of the illegal characters include the following characters :    !#â‚¬$%()`'?Â§&Ã§[]@{}=+<>:;,*^|"  and
    'char(10) = LF - char(13) = CR - Char(34) = " - char(39) = ' - char(126) = tilde

    Dim intI As Integer
    Dim intPassedString As Integer
    Dim intCheckString As Integer
    Dim strChar As String
    Dim strIllegalChars As String
    Dim intReplaceLen As Integer

    If IsNull(strCheck) Then Exit Function

    CheckIllegalCharInString = ""
    'char(10) = LF - char(13) = CR - Char(34) = " - char(39) = ' - char(126) = tilde
    strIllegalChars = "Ã©Ã¨!#â‚¬$%`'?Â§&Ã§Â£Â²Â³@{}=+<>:;,*^|Âµ""" & Chr(34) & Chr(39) & Chr(13) & Chr(10) & Chr(126)  'add/remove characters you need removed to this string

    intPassedString = Len(strCheck)
    intCheckString = Len(strIllegalChars)

    For intI = 1 To intCheckString
        strChar = Mid(strIllegalChars, intI, 1)
        If InStr(strCheck, strChar) > 0 Then
            CheckIllegalCharInString = strChar
            Exit For
        End If
    Next intI
End Function
