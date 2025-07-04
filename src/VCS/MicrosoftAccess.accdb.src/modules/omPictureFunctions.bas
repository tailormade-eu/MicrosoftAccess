Attribute VB_Name = "omPictureFunctions"
Option Compare Database
Option Explicit

Public Function savePict(pImage As Access.Image)
    Dim fname As String 'The name of the file to save the picture to
    fname = Environ("Temp") + "\temp.jpg" ' Destination file path

    Dim iFileNum As Double
    iFileNum = FreeFile 'The next free file from the file system

    Dim pngImage As String 'Stores the image data as a string
    pngImage = StrConv(pImage.PictureData, vbUnicode) 'Convert the byte array to a string

    'Writes the string to the file
    Open fname For Binary Access Write As iFileNum
        Put #iFileNum, , pngImage
    Close #iFileNum
End Function
Public Function savePictEMF(pImage As Access.Image)
    Dim fname As String 'The name of the file to save the picture to
    Dim iFileNum As Double
    Dim bArray() As Byte, cArray() As Byte
    Dim lngRet As Long

    fname = Environ("Temp") + "\temp.emf" ' Destination file path
    iFileNum = FreeFile 'The next free file from the file system

    ' Resize to hold entire PictureData prop
    ReDim bArray(LenB(pImage.PictureData) - 1)
    ' Resize to hold the EMF wrapped in the PictureData prop
    ReDim cArray(LenB(pImage.PictureData) - (1 + 8))
    ' Copy to our array
    bArray = pImage.PictureData
    For lngRet = 8 To UBound(cArray)
        cArray(lngRet - 8) = bArray(lngRet)
    Next

    Open fname For Binary Access Write As iFileNum
    'Write the byte array to the file
    Put #iFileNum, , cArray
    Close #iFileNum
End Function
