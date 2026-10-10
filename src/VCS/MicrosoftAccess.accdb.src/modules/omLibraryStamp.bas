Attribute VB_Name = "omLibraryStamp"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=fbb7efab

' Every library module carries a line  '@lib <date> sha=<8 hex>  with a hash of its own text.
' LibraryCheck recalculates that hash in the current database and shows which om* modules are
' in order, changed locally, or older than the library (compared with library-stamps.txt).
' The hash is calculated exactly as scripts/libstamp.py does (keep both in line):
' no VERSION/BEGIN/END class header, no Attribute lines, no stamp line, lower case, no trailing
' spaces or tabs, no leading or trailing blank lines, LF line ends, SHA-256 over UTF-8, first 8 hex.

Private Const STAMP_PREFIX As String = "'@lib "

' First 8 hex characters of the hash of a module of the current database ("" if it cannot be read).
Public Function GetModuleHash(moduleName As String) As String
    GetModuleHash = HashOfText(GetModuleText(moduleName))
End Function

' The stamp in the module: "yyyy-mm-dd sha=xxxxxxxx", or "" when the module has none.
Public Function GetModuleStamp(moduleName As String) As String
Dim lines() As String
Dim i As Long
Dim s As String

    lines = Split(NormalizeLineEnds(GetModuleText(moduleName)), vbLf)
    For i = LBound(lines) To UBound(lines)
        s = LTrim(lines(i))
        If Left(s, Len(STAMP_PREFIX)) = STAMP_PREFIX Then
            GetModuleStamp = Trim(Mid(s, Len(STAMP_PREFIX) + 1))
            Exit Function
        End If
    Next i
End Function

' One line per om* module that needs attention:
' "<module>;<state>" with state = NO STAMP, CHANGED LOCALLY, LIBRARY NEWER or LOCAL ONLY.
' manifestFile = library-stamps.txt of the library (optional; without it only the local check runs).
' Returns "" when everything is in order.
Public Function LibraryCheck(Optional manifestFile As String = "", Optional prefix As String = "om") As String
Dim obj As AccessObject
Dim manifest As Object
Dim stamp As String
Dim sha As String
Dim fileSha As String
Dim fileDate As String
Dim m As Variant

    If Len(manifestFile) > 0 Then
        Set manifest = ReadManifest(manifestFile)
    End If
    For Each obj In CurrentProject.AllModules
        If LCase(Left(obj.Name, Len(prefix))) = LCase(prefix) Then
            stamp = GetModuleStamp(obj.Name)
            sha = GetModuleHash(obj.Name)
            If Len(stamp) = 0 Then
                LibraryCheck = LibraryCheck & obj.Name & ";NO STAMP" & vbCrLf
            ElseIf Right(stamp, 8) <> sha Then
                LibraryCheck = LibraryCheck & obj.Name & ";CHANGED LOCALLY" & vbCrLf
            ElseIf Not manifest Is Nothing Then
                If manifest.Exists(LCase(obj.Name)) Then
                    m = Split(manifest(LCase(obj.Name)), ";")
                    fileDate = m(0)
                    fileSha = m(1)
                    If fileSha <> sha Then
                        LibraryCheck = LibraryCheck & obj.Name & ";LIBRARY NEWER" & vbCrLf
                    End If
                Else
                    LibraryCheck = LibraryCheck & obj.Name & ";LOCAL ONLY" & vbCrLf
                End If
            End If
        End If
    Next obj
End Function

' name -> "date;sha" (lower case names) from a library-stamps.txt file.
Private Function ReadManifest(manifestFile As String) As Object
Dim dict As Object
Dim ts As Object
Dim parts() As String
Dim line As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set ts = CreateObject("Scripting.FileSystemObject").OpenTextFile(manifestFile, 1, False)
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        parts = Split(line, ";")
        If UBound(parts) = 2 Then
            dict(LCase(parts(0))) = parts(1) & ";" & parts(2)
        End If
    Loop
    ts.Close
    Set ReadManifest = dict
End Function

' The text of a module as Access saves it (same layout as the exported file, without BOM).
Private Function GetModuleText(moduleName As String) As String
Dim tmp As String
Dim fso As Object
Dim ts As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    tmp = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName)
    On Error GoTo Fout
    Application.SaveAsText acModule, moduleName, tmp
    ' SaveAsText writes the code in the ANSI code page of the pc (no BOM)
    Set ts = fso.OpenTextFile(tmp, 1, False, 0)
    If Not ts.AtEndOfStream Then
        GetModuleText = ts.ReadAll
    End If
    ts.Close
    fso.DeleteFile tmp, True
    Exit Function
Fout:
    GetModuleText = ""
    On Error Resume Next
    fso.DeleteFile tmp, True
End Function

Private Function NormalizeLineEnds(ByVal text As String) As String
    NormalizeLineEnds = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
End Function

Private Function HashOfText(text As String) As String
Dim normalized As String

    If Len(text) = 0 Then
        Exit Function
    End If
    normalized = NormalizeModuleText(text)
    HashOfText = Left(omHashAlgorithmsFunctions.SHA256(normalized), 8)
End Function

' The text that is hashed (see the top of this module); public so it can be tested.
Public Function NormalizeModuleText(ByVal text As String) As String
Dim lines() As String
Dim i As Long
Dim first As Long
Dim line As String
Dim result As String

    lines = Split(NormalizeLineEnds(text), vbLf)
    first = LBound(lines)
    Do While first <= UBound(lines)
        If Len(Trim(lines(first))) > 0 Then Exit Do
        first = first + 1
    Loop
    If first <= UBound(lines) Then
        If Left(lines(first), 8) = "VERSION " Then
            Do While first <= UBound(lines)
                If Trim(lines(first)) = "END" Then Exit Do
                first = first + 1
            Loop
            first = first + 1
        End If
    End If
    For i = first To UBound(lines)
        line = lines(i)
        If Left(line, 10) = "Attribute " Then
            ' not part of the hashed text
        ElseIf Left(LTrim(line), Len(STAMP_PREFIX)) = STAMP_PREFIX Then
            ' the stamp itself
        Else
            line = LCase(line)
            Do While Right(line, 1) = " " Or Right(line, 1) = vbTab
                line = Left(line, Len(line) - 1)
            Loop
            result = result & line & vbLf
        End If
    Next i
    Do While Left(result, 1) = vbLf
        result = Mid(result, 2)
    Loop
    Do While Right(result, 1) = vbLf
        result = Left(result, Len(result) - 1)
    Loop
    NormalizeModuleText = result
End Function
