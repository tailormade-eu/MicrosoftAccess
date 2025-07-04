Attribute VB_Name = "omExportFunctions"
Option Compare Database
Option Explicit



Public Function ExportQuery(queryName As String, Optional exportPath As String = "", Optional filename As String, Optional withTimestamp As Boolean = False) As String
Dim fullFilename As String
Dim omxl As New omExcel

    If IsNullOrEmpty(exportPath) Then
        exportPath = gFso.BuildPath(GetDesktopFolder, "ExportQuery")
    End If
    omFileFunctions.CreateFolderPath (exportPath)
    If omStringFunctions.IsNullOrEmpty(filename) Then
        fullFilename = gFso.BuildPath(exportPath, queryName & ".xlsx")
    Else
        fullFilename = gFso.BuildPath(exportPath, filename & ".xlsx")
    End If
    If withTimestamp Then
        fullFilename = Replace(fullFilename, ".xlsx", "_" & omDateFunctions.GetTimeStamp() & ".xlsx")
    End If
    If gFso.FileExists(fullFilename) Then
        gFso.DeleteFile fullFilename, True
    End If
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, queryName, fullFilename, True
    omxl.LoadWB fullFilename
    omxl.SetVisible True
    omxl.SetHeaderBold
    omxl.AutoFit
    omxl.CloseWB True
    Set omxl = Nothing
    ExportQuery = fullFilename

End Function
