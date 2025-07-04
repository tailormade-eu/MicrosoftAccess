VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omExcelEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub Merge(Template As String, SaveAs As String, rs As ADODB.Recordset, Optional show As Boolean = False)
Dim i As Integer
Dim xlApp As New Excel.Application
Dim xlWb As Excel.Workbook

    If Not rs.EOF Then
        Set xlWb = xlApp.Workbooks.Open(Template)
        MergeFromCurrentRecord xlApp.ActiveSheet, rs, "<<,>>,<,>,«,»,“,»"
        If Not show Then
            xlWb.SaveAs SaveAs
        Else
            omExcelFunctions.SetVisible xlApp, True
        End If
    End If

    Set xlWb = Nothing
    If Not show Then
        xlApp.Quit
    End If
End Sub

Public Sub MergeFromCurrentRecord(xlWs As Excel.Worksheet, rs As ADODB.Recordset, Optional delimiters As String = "<<,>>")
Dim k As Integer
Dim j As Integer
Dim del() As String
Dim findString As String

    If Not rs.EOF Then
        del = Split(delimiters, ",")
        For j = 0 To UBound(del) Step 2
            For k = 0 To rs.Fields.Count - 1
                findString = del(j) & rs.Fields(k).Name & del(j + 1)
                While omObjectFunctions.NotIsNothing(omExcelFunctions.FindFirst(xlWs, findString))
                    omExcelFunctions.ReplaceFirst xlWs, findString, Nz(rs.Fields(k), "")
                Wend
            Next k
        Next
    End If
End Sub
