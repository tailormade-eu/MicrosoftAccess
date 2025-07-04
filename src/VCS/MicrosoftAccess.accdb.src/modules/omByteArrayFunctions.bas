Attribute VB_Name = "omByteArrayFunctions"
Option Compare Database
Option Explicit

Public Function ByteArrayToHexString(ba() As Byte) As String
Dim i As Long
Dim s As String

    For i = 0 To UBound(ba)
        s = s & Right("00" & Hex$(ba(i)), 2)
    Next
    ByteArrayToHexString = "0x" & s
End Function
Public Function ByteArrayToString(ba() As Byte) As LongLong
Dim i As Long
Dim s As Double

    For i = 0 To UBound(ba)
        s = s * 256 + ba(i)
    Next
    ByteArrayToString = s
End Function

Public Sub ByteArrayToHexString_Test()
    Debug.Print ByteArrayToHexString(RandomByteArray(4))
End Sub
Public Sub ByteArrayToString_Test()
    Debug.Print ByteArrayToString(RandomByteArray(8))
End Sub

Private Function RandomByteArray(ByVal arrayLength As Integer) As Byte()
' Demo/helper function to create random byte array
    Dim retVal() As Byte, i As Integer
    Randomize
    ReDim retVal(arrayLength - 1)
    For i = 0 To arrayLength - 1
        retVal(i) = CByte(Rnd() * 255)
    Next i
    RandomByteArray = retVal
End Function
