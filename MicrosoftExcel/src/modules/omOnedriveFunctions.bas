Attribute VB_Name = "omOnedriveFunctions"
Option Explicit
'@lib 2026-10-10 sha=9ecdfb19

' A workbook that lives in a synced OneDrive or SharePoint/Teams folder reports a web address (https://...) as its Path.
' These functions give the real folder on this pc, found through the sync settings in the registry
' (HKCU\SOFTWARE\SyncEngines\Providers\OneDrive: UrlNamespace = web address of a library, MountPoint = its local folder).

' The local folder of the workbook (default: the workbook that holds this code). A workbook that is not on the web is returned as is.
Public Function GetWorkbookLocalPath(Optional wb As Workbook = Nothing) As String
Dim reg As Object
Dim subKeys As Variant
Dim key As Variant
Dim urlNamespace As String
Dim mountPoint As String
Dim candidate As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    GetWorkbookLocalPath = wb.Path
    If InStr(1, wb.Path, "https://", vbTextCompare) <> 1 Then Exit Function

    On Error GoTo Done
    Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    reg.EnumKey &H80000001, "SOFTWARE\SyncEngines\Providers\OneDrive", subKeys
    If IsNull(subKeys) Then Exit Function
    For Each key In subKeys
        urlNamespace = ""
        mountPoint = ""
        reg.GetStringValue &H80000001, "SOFTWARE\SyncEngines\Providers\OneDrive\" & key, "UrlNamespace", urlNamespace
        reg.GetStringValue &H80000001, "SOFTWARE\SyncEngines\Providers\OneDrive\" & key, "MountPoint", mountPoint
        candidate = UrlToLocalPath(wb.Path, urlNamespace, mountPoint, wb.Name)
        If Len(candidate) > 0 Then
            GetWorkbookLocalPath = candidate
            Exit Function
        End If
    Next key
Done:
End Function

' The full local file name of the workbook (default: the workbook that holds this code).
Public Function GetWorkbookLocalFilename(Optional wb As Workbook = Nothing) As String
    If wb Is Nothing Then Set wb = ThisWorkbook
    GetWorkbookLocalFilename = GetWorkbookLocalPath(wb) & "\" & wb.Name
End Function

' Maps a web address to a local folder, given the web address (urlNamespace) and the local folder (mountPoint) of the library.
' Returns "" when the address is not inside that library or when the folder does not exist on this pc.
' A personal OneDrive has its account id in the web address but not in the local folder; both forms are tried.
' fileName (optional): when given, the folder only counts if that file is in it.
Public Function UrlToLocalPath(url As String, urlNamespace As String, mountPoint As String, Optional fileName As String = "") As String
Dim rest As String
Dim slash As Long
Dim variants(1) As String
Dim i As Long
Dim folder As String

    If Len(urlNamespace) = 0 Or Len(mountPoint) = 0 Then Exit Function
    url = UrlDecode(url)
    urlNamespace = UrlDecode(urlNamespace)
    If Right(urlNamespace, 1) = "/" Then urlNamespace = Left(urlNamespace, Len(urlNamespace) - 1)
    If Right(mountPoint, 1) = "\" Then mountPoint = Left(mountPoint, Len(mountPoint) - 1)
    If InStr(1, url, urlNamespace, vbTextCompare) <> 1 Then Exit Function
    rest = Mid(url, Len(urlNamespace) + 1)
    If Len(rest) > 0 And Left(rest, 1) <> "/" Then Exit Function
    variants(0) = rest
    slash = InStr(2, rest, "/")
    If slash > 0 Then variants(1) = Mid(rest, slash) Else variants(1) = ""
    For i = 0 To 1
        If i = 0 Or InStr(1, urlNamespace, "d.docs.live.net", vbTextCompare) > 0 Then
            folder = mountPoint & Replace(variants(i), "/", "\")
            If FolderHasFile(folder, fileName) Then
                UrlToLocalPath = folder
                Exit Function
            End If
        End If
    Next i
End Function

Private Function FolderHasFile(folder As String, fileName As String) As Boolean
    On Error Resume Next
    If Len(fileName) = 0 Then
        FolderHasFile = (Len(Dir(folder & "\", vbDirectory)) > 0)
    Else
        FolderHasFile = (Len(Dir(folder & "\" & fileName)) > 0)
    End If
End Function

Private Function UrlDecode(s As String) As String
Dim i As Long
Dim result As String

    i = 1
    Do While i <= Len(s)
        If Mid(s, i, 1) = "%" And i + 2 <= Len(s) Then
            If Mid(s, i + 1, 2) Like "[0-9A-Fa-f][0-9A-Fa-f]" Then
                result = result & Chr(CLng("&H" & Mid(s, i + 1, 2)))
                i = i + 3
                GoTo NextChar
            End If
        End If
        result = result & Mid(s, i, 1)
        i = i + 1
NextChar:
    Loop
    UrlDecode = result
End Function
