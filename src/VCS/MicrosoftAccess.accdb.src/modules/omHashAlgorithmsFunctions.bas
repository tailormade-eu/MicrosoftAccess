Attribute VB_Name = "omHashAlgorithmsFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=392fa529
Private gHashAlgorithms As New omHashAlgorithms
Public Function MD5(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    MD5 = gHashAlgorithms.MD5(sIn, bB64)
End Function

Public Function SHA1(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    SHA1 = gHashAlgorithms.SHA1(sIn, bB64)
End Function

Public Function SHA256(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    SHA256 = gHashAlgorithms.SHA256(sIn, bB64)
End Function

Public Function SHA384(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    SHA384 = gHashAlgorithms.SHA384(sIn, bB64)
End Function

Public Function SHA512(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    SHA512 = gHashAlgorithms.SHA512(sIn, bB64)
End Function

Public Function ConvToBase64String(vIn As Variant) As Variant
    ConvToBase64String = gHashAlgorithms.ConvToBase64String(vIn)
End Function

Private Function ConvToHexString(vIn As Variant) As Variant
    ConvToHexString = gHashAlgorithms.ConvToHexString(vIn)
End Function
