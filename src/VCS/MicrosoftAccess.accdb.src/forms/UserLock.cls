Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim filename As String


Private Sub Form_Load()
Dim ts As TextStream
    filename = gFso.BuildPath(CurrentProject.path, "program in use by " & gUser.Name & " - DO NOT DELETE THIS FILE")

    Set ts = gFso.CreateTextFile(filename)
    ts.Close
    Set ts = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gFso.FileExists(filename) Then
        gFso.DeleteFile filename
    End If
End Sub
