VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public xlApp As New Excel.Application
Public xlWb As Excel.Workbook

Public Sub LoadWB(filename As String)
    Set xlWb = xlApp.Workbooks.Open(filename)
End Sub
Public Sub CloseWB(Optional saveChanges As Boolean = True, Optional filename As String = "")
    If IsNullOrEmpty(filename) Then
        xlWb.Close saveChanges
    Else
        xlWb.Close saveChanges, filename
    End If
    Set xlWb = Nothing
End Sub

Public Sub SetVisible(Optional visible As Boolean = True)
    omExcelFunctions.SetVisible xlApp, visible
End Sub
Public Sub SelectColumn(columnIndex As Long)
    omExcelFunctions.SelectColumn xlWb.ActiveSheet, columnIndex
End Sub
Public Sub SelectRow(rowIndex As Long)
    omExcelFunctions.SelectRow xlWb.ActiveSheet, rowIndex
End Sub
Public Sub SetSelectionBold(Optional boldFlag As Boolean = True)
    omExcelFunctions.SetSelectionBold xlWb.ActiveSheet, boldFlag
End Sub
Public Sub FormatColumn(columnIndex As Long, numberFormat As String)
    omExcelFunctions.FormatColumn xlWb.ActiveSheet, columnIndex, numberFormat
End Sub
Public Sub FormatCells(Optional rowIndexStart As Long = 0, Optional columnIndexStart As Long = 0, Optional rowIndexEnd As Long = 0, Optional columnIndexEnd As Long = 0, Optional numberFormat As Variant = Null, Optional horizontalAlignment As Excel.Constants = xlLeft, Optional verticalAlignment As Excel.Constants = xlTop, Optional MergeCells As Boolean = False, Optional WrapText As Boolean = False)
    omExcelFunctions.FormatCells xlWb.ActiveSheet, rowIndexStart, columnIndexStart, rowIndexEnd, columnIndexEnd, numberFormat, horizontalAlignment, verticalAlignment, MergeCells, WrapText
End Sub
Public Sub WriteCell(rowIndex As Long, columnIndex As Long, Value As Variant, Optional numberFormat As Variant = Null, Optional horizontalAlignment As Excel.Constants = xlLeft, Optional verticalAlignment As Excel.Constants = xlTop)
    omExcelFunctions.WriteCell xlWb.ActiveSheet, rowIndex, columnIndex, Value
    Me.FormatCells rowIndexStart:=rowIndex, columnIndexStart:=columnIndex, numberFormat:=numberFormat, horizontalAlignment:=horizontalAlignment, verticalAlignment:=verticalAlignment
End Sub
Public Sub InsertRows(rowIndex As Long, Optional numberOfRows As Long = 1)
Dim i As Long
    For i = 1 To numberOfRows
        omExcelFunctions.InsertRow xlWb.ActiveSheet, rowIndex
    Next
End Sub
Public Sub SetHeaderBold()
    xlWb.ActiveSheet.rows(1).Select
    xlWb.Application.Selection.font.bold = True
End Sub

Public Sub AutoFit()
    omExcelFunctions.AutoFitColumns xlWb.ActiveSheet
    omExcelFunctions.AutoFitRows xlWb.ActiveSheet
End Sub
Public Sub SelectionUpdateBold(bold As Boolean)
    UpdateFontOnSelection xlWb.ActiveSheet, bold
End Sub
Private Sub Class_Terminate()
    Set xlWb = Nothing
    xlApp.Quit
    Set xlApp = Nothing
End Sub
