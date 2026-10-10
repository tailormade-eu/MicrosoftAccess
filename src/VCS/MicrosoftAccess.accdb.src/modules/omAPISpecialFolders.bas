Attribute VB_Name = "omAPISpecialFolders"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=d461cf8e

#If VBA7 Then
    Private Declare PtrSafe Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwnd As LongPtr, ByVal nFolder As Long, Pidl As LongPtr) As Long
    Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" (ByVal Pidl As LongPtr, ByVal folderPath As String) As Long
#Else
    Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwnd As Long, ByVal nFolder As Long, Pidl As Long) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal Pidl As Long, ByVal folderPath As String) As Long
#End If

 Enum SpecialFolderConst
 sfAppData = 26
 sfCDBurning = 59
 sfCookies = 33
 sfDesktop = 0
 sfFavorites = 6
 sfFonts = 20
 sfHistory = 34
 sfLocalAppData = 28
 sfMyDocuments = 5
 sfMyMusic = 13
 sfMyPictures = 39
 sfMyVideo = 14
 sfNetHood = 19
 sfPrintHood = 27
 sfProfile = 40
 sfProgramFiles = 38
 sfRecent = 8
 sfSendTo = 9
 sfStartMenu = 11
 sfStartMenuPrograms = 2
 sfStartUp = 7
 sfSystem = 37
 sfTempInternet = 32
 sfTemplates = 21
 sfWindows = 36
 End Enum

 Public Function SpecialFolder(ByVal SFConst As SpecialFolderConst) As String
#If VBA7 Then
 Dim Pidl As LongPtr
#Else
 Dim Pidl As Long
#End If
 Dim s As String * 260
 Dim l As Long
 l = SHGetSpecialFolderLocation(0, SFConst, Pidl)

 If l = 0 Then
 l = SHGetPathFromIDList(Pidl, s)
 If l = 1 Then
 s = Left$(s, InStr(s, vbNullChar) - 1)
 SpecialFolder = Trim(s)
 End If
 End If
 End Function
