VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omHashAlgorithms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private oD As Object
Private oT As Object
Private oMD5 As Object
Private oSHA1 As Object
Private oSHA256 As Object
Private oSHA384 As Object
Private oSHA512 As Object

Public Function MD5(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    'Late bound (CreateObject): no reference needed, uses the .NET Framework COM classes

    'Test with empty string input:
    'Hex:   d41d8cd98f00...etc
    'Base-64: 1B2M2Y8Asg...etc

    Dim TextToHash() As Byte
    Dim bytes() As Byte


    TextToHash = oT.GetBytes_4(sIn)
    bytes = oMD5.ComputeHash_2((TextToHash))

    If bB64 = True Then
       MD5 = ConvToBase64String(bytes)
    Else
       MD5 = ConvToHexString(bytes)
    End If

End Function

Public Function SHA1(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    'Late bound (CreateObject): no reference needed, uses the .NET Framework COM classes

    'Test with empty string input:
    '40 Hex:   da39a3ee5e6...etc
    '28 Base-64:   2jmj7l5rSw0yVb...etc


    Dim TextToHash() As Byte
    Dim bytes() As Byte


    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA1.ComputeHash_2((TextToHash))

    If bB64 = True Then
       SHA1 = ConvToBase64String(bytes)
    Else
       SHA1 = ConvToHexString(bytes)
    End If


End Function

Public Function SHA256(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    'Late bound (CreateObject): no reference needed, uses the .NET Framework COM classes

    'Test with empty string input:
    '64 Hex:   e3b0c44298f...etc
    '44 Base-64:   47DEQpj8HBSa+/...etc


    Dim TextToHash() As Byte, bytes() As Byte


    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA256.ComputeHash_2((TextToHash))

    If bB64 = True Then
       SHA256 = ConvToBase64String(bytes)
    Else
       SHA256 = ConvToHexString(bytes)
    End If


End Function

Public Function SHA384(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    'Late bound (CreateObject): no reference needed, uses the .NET Framework COM classes

    'Test with empty string input:
    '96 Hex:   38b060a751ac...etc
    '64 Base-64:   OLBgp1GsljhM2T...etc


    Dim TextToHash() As Byte, bytes() As Byte


    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA384.ComputeHash_2((TextToHash))

    If bB64 = True Then
       SHA384 = ConvToBase64String(bytes)
    Else
       SHA384 = ConvToHexString(bytes)
    End If


End Function

Public Function SHA512(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    'Late bound (CreateObject): no reference needed, uses the .NET Framework COM classes

    'Test with empty string input:
    '128 Hex:   cf83e1357eefb8bd...etc
    '88 Base-64:   z4PhNX7vuL3xVChQ...etc

    Dim TextToHash() As Byte, bytes() As Byte


    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA512.ComputeHash_2((TextToHash))

    If bB64 = True Then
       SHA512 = ConvToBase64String(bytes)
    Else
       SHA512 = ConvToHexString(bytes)
    End If


End Function

Public Function ConvToBase64String(vIn As Variant) As Variant

      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.text, vbLf, "")


End Function

Public Function ConvToHexString(vIn As Variant) As Variant


      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.text, vbLf, "")


End Function


Private Sub Class_Initialize()
    Set oD = CreateObject("MSXML2.DOMDocument")
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    Set oSHA1 = CreateObject("System.Security.Cryptography.SHA1Managed")
    Set oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    Set oSHA384 = CreateObject("System.Security.Cryptography.SHA384Managed")
    Set oSHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")
End Sub

Private Sub Class_Terminate()
    Set oD = Nothing
    Set oT = Nothing
    Set oMD5 = Nothing
    Set oSHA1 = Nothing
    Set oSHA256 = Nothing
    Set oSHA384 = Nothing
    Set oSHA512 = Nothing
End Sub
