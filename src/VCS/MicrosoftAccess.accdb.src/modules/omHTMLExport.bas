Attribute VB_Name = "omHTMLExport"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=63d69791

' Writes the result of a SELECT as a simple HTML table to a UTF-8 file.
' Uses the database connection of the current project (ADODB) and late-bound ADODB.Stream.
Public Sub OutputHTMLRaw(ByVal sql As String, ByVal title As String, ByVal filename As String)
Dim rs As New ADODB.Recordset
Dim html As String
Dim i As Long
Dim stm As Object

    rs.Open sql, CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly
    html = "<!DOCTYPE html><html><head><meta charset=""utf-8""><title>" & HtmlEncode(title) & "</title></head>"
    html = html & "<body><table border=1 cellspacing=0><thead><tr>"
    For i = 0 To rs.Fields.Count - 1
        html = html & "<th>" & HtmlEncode(rs.Fields(i).Name) & "</th>"
    Next i
    html = html & "</tr></thead><tbody>"
    Do While Not rs.EOF
        html = html & "<tr>"
        For i = 0 To rs.Fields.Count - 1
            html = html & "<td>" & HtmlEncode(rs.Fields(i).Value & "") & "</td>"
        Next i
        html = html & "</tr>"
        rs.MoveNext
    Loop
    html = html & "</tbody></table></body></html>"
    rs.Close

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText html
    stm.SaveToFile filename, 2 ' adSaveCreateOverWrite
    stm.Close
End Sub

Private Function HtmlEncode(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, Chr(34), "&quot;")
    HtmlEncode = s
End Function
