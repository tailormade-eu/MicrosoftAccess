Attribute VB_Name = "omExcelVersionControl"
Option Explicit
'@lib 2026-10-10 sha=02202f7b

' Writes the sheets and the VBA code of a workbook to text files in a folder Src next to the workbook, so they can go in git, and reads them back:
'   Src\ExportedSheets\<sheet>.csv        the data of each worksheet
'   Src\ExportedVBA\<name>.bas/.cls/.frm  standard modules, class modules and userforms
'   Src\ExportedSheetVBA\<sheet>.cls      the code behind the sheets and ThisWorkbook
' Needs omOnedriveFunctions (a workbook in OneDrive) and omExcelApplicationFunctions (SheetExists), and in Excel the option
' "Trust access to the VBA project object model" (File > Options > Trust Center > Macro Settings).
' wb (optional, on every procedure): the workbook; default the workbook that holds this code.

Private Const MODULE_NAME As String = "omExcelVersionControl"
Private Const VBEXT_STD_MODULE As Long = 1
Private Const VBEXT_CLASS_MODULE As Long = 2
Private Const VBEXT_FORM As Long = 3
Private Const VBEXT_DOCUMENT As Long = 100

' Exports the sheet data, the VBA components and the code behind the sheets. showMessage: tell when it is done (False for automatic use).
Public Sub ExportAll(Optional objName As String = "", Optional showMessage As Boolean = True, Optional wb As Workbook = Nothing)
    ExportSheetData objName, wb
    ExportVBA objName, wb
    ExportSheetCode objName, wb
    If showMessage Then MsgBox "Export complete!", vbInformation
End Sub

' Reads everything back. Existing components with the same name are replaced; this module is never replaced while it runs.
Public Sub ImportAll(Optional objName As String = "", Optional showMessage As Boolean = True, Optional wb As Workbook = Nothing)
    ImportSheetData objName, wb
    ImportVBA objName, wb
    ImportSheetCode objName, wb
    If showMessage Then MsgBox "Import complete!", vbInformation
End Sub

' Each worksheet as a .csv file (objName: only that sheet).
Public Sub ExportSheetData(Optional objName As String = "", Optional wb As Workbook = Nothing)
Dim sh As Worksheet
Dim exportPath As String
Dim originalVisibility As XlSheetVisibility
Dim alerts As Boolean

    If wb Is Nothing Then Set wb = ThisWorkbook
    exportPath = SourceFolder(wb) & "\ExportedSheets"
    EnsureFolderExists exportPath
    alerts = Application.DisplayAlerts
    For Each sh In wb.Worksheets
        If Len(objName) = 0 Or SanitizeFilename(sh.Name) = SanitizeFilename(objName) Then
            originalVisibility = sh.Visible
            If sh.Visible <> xlSheetVisible Then sh.Visible = xlSheetVisible
            sh.Copy
            Application.DisplayAlerts = False
            With ActiveWorkbook
                .SaveAs fileName:=exportPath & "\" & SanitizeFilename(sh.Name) & ".csv", FileFormat:=xlCSV, CreateBackup:=False
                .Close saveChanges:=False
            End With
            Application.DisplayAlerts = alerts
            sh.Visible = originalVisibility
        End If
    Next sh
End Sub

' Each .csv file back into the worksheet with that name (the sheet is cleared first; a missing sheet is added at the end).
' A .csv file holds values only (no formulas or formats) and uses the list separator of this pc, so export and import on the same settings.
Public Sub ImportSheetData(Optional objName As String = "", Optional wb As Workbook = Nothing)
Dim importPath As String
Dim fileName As String
Dim wbTemp As Workbook
Dim targetSheet As Worksheet
Dim sheetName As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    importPath = SourceFolder(wb) & "\ExportedSheets\"
    EnsureFolderExists Left(importPath, Len(importPath) - 1)
    fileName = Dir(importPath & "*.csv")
    Do While Len(fileName) > 0
        sheetName = SanitizeFilename(Left(fileName, InStrRev(fileName, ".") - 1))
        If Len(objName) = 0 Or sheetName = SanitizeFilename(objName) Then
            Set wbTemp = Workbooks.Open(importPath & fileName)
            If omExcelApplicationFunctions.SheetExists(sheetName, wb) Then
                Set targetSheet = wb.Worksheets(sheetName)
                targetSheet.Cells.Clear
            Else
                Set targetSheet = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
                targetSheet.Name = sheetName
            End If
            wbTemp.Worksheets(1).UsedRange.Copy targetSheet.Range("A1")
            wbTemp.Close saveChanges:=False
        End If
        fileName = Dir()
    Loop
End Sub

' Each standard module, class module and userform as a file (objName: only that component).
Public Sub ExportVBA(Optional objName As String = "", Optional wb As Workbook = Nothing)
Dim comp As Object
Dim exportPath As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    exportPath = SourceFolder(wb) & "\ExportedVBA"
    EnsureFolderExists exportPath
    For Each comp In wb.VBProject.VBComponents
        Select Case comp.Type
            Case VBEXT_STD_MODULE, VBEXT_CLASS_MODULE, VBEXT_FORM
                If Len(objName) = 0 Or comp.Name = objName Then comp.Export exportPath & "\" & comp.Name & GetExtension(comp.Type)
        End Select
    Next comp
End Sub

' Each .bas, .cls and .frm file as a component. A standard or class module that exists already gets the new code in place (the module
' keeps its name, so code that calls it keeps working; the Attribute lines of the file, such as VB_PredeclaredId, are not carried over);
' a new module or a userform is added (a userform with the same name is removed first).
' This module itself is skipped (it cannot be replaced while it runs); the code behind sheets is not touched (see ImportSheetCode).
Public Sub ImportVBA(Optional objName As String = "", Optional wb As Workbook = Nothing)
Dim importPath As String
Dim pattern As Variant
Dim fileName As String
Dim compName As String
Dim existing As Object

    If wb Is Nothing Then Set wb = ThisWorkbook
    importPath = SourceFolder(wb) & "\ExportedVBA\"
    EnsureFolderExists Left(importPath, Len(importPath) - 1)
    For Each pattern In Array("*.bas", "*.cls", "*.frm")
        fileName = Dir(importPath & pattern)
        Do While Len(fileName) > 0
            compName = Left(fileName, InStrRev(fileName, ".") - 1)
            If (Len(objName) = 0 Or compName = objName) And StrComp(compName, MODULE_NAME, vbTextCompare) <> 0 Then
                Set existing = Nothing
                On Error Resume Next
                Set existing = wb.VBProject.VBComponents(compName)
                On Error GoTo 0
                If existing Is Nothing Then
                    wb.VBProject.VBComponents.Import importPath & fileName
                ElseIf existing.Type = VBEXT_STD_MODULE Or existing.Type = VBEXT_CLASS_MODULE Then
                    With existing.CodeModule
                        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
                        .AddFromString ReadSourceWithoutHeader(importPath & fileName)
                    End With
                ElseIf existing.Type <> VBEXT_DOCUMENT Then
                    wb.VBProject.VBComponents.Remove existing
                    wb.VBProject.VBComponents.Import importPath & fileName
                End If
            End If
            fileName = Dir()
        Loop
    Next pattern
End Sub

' The code behind each sheet and ThisWorkbook as a .cls file, named after the sheet (objName: only that one).
Public Sub ExportSheetCode(Optional objName As String = "", Optional wb As Workbook = Nothing)
Dim comp As Object
Dim exportPath As String
Dim fileBase As String
Dim fNum As Integer
Dim sh As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    exportPath = SourceFolder(wb) & "\ExportedSheetVBA"
    EnsureFolderExists exportPath
    For Each comp In wb.VBProject.VBComponents
        If comp.Type = VBEXT_DOCUMENT Then
            If Len(objName) = 0 Or comp.Name = objName Then
                fileBase = comp.Name
                For Each sh In wb.Worksheets
                    If sh.CodeName = comp.Name Then fileBase = sh.Name
                Next sh
                If comp.CodeModule.CountOfLines > 0 Then
                    fNum = FreeFile
                    Open exportPath & "\" & SanitizeFilename(fileBase) & ".cls" For Output As #fNum
                    Print #fNum, comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines);
                    Close #fNum
                End If
            End If
        End If
    Next comp
End Sub

' Each .cls file in ExportedSheetVBA as the code behind the sheet with that name (or ThisWorkbook); the old code is replaced.
Public Sub ImportSheetCode(Optional objName As String = "", Optional wb As Workbook = Nothing)
Dim importPath As String
Dim fileName As String
Dim baseName As String
Dim targetComp As Object
Dim sh As Worksheet
Dim fNum As Integer
Dim fileContent As String

    If wb Is Nothing Then Set wb = ThisWorkbook
    importPath = SourceFolder(wb) & "\ExportedSheetVBA\"
    EnsureFolderExists Left(importPath, Len(importPath) - 1)
    fileName = Dir(importPath & "*.cls")
    Do While Len(fileName) > 0
        baseName = Left(fileName, InStrRev(fileName, ".") - 1)
        If Len(objName) = 0 Or baseName = SanitizeFilename(objName) Then
            Set targetComp = Nothing
            If baseName = "ThisWorkbook" Then
                Set targetComp = wb.VBProject.VBComponents("ThisWorkbook")
            Else
                For Each sh In wb.Worksheets
                    If SanitizeFilename(sh.Name) = baseName Then
                        Set targetComp = wb.VBProject.VBComponents(sh.CodeName)
                        Exit For
                    End If
                Next sh
            End If
            If Not targetComp Is Nothing Then
                fNum = FreeFile
                Open importPath & fileName For Input As #fNum
                fileContent = Input$(LOF(fNum), fNum)
                Close #fNum
                With targetComp.CodeModule
                    If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
                    .AddFromString fileContent
                End With
            End If
        End If
        fileName = Dir()
    Loop
End Sub

' The extension of the exported file for a component type (1 = standard module, 2 = class module, 3 = userform).
Public Function GetExtension(compType As Long) As String
    Select Case compType
        Case VBEXT_STD_MODULE
            GetExtension = ".bas"
        Case VBEXT_CLASS_MODULE
            GetExtension = ".cls"
        Case VBEXT_FORM
            GetExtension = ".frm"
        Case Else
            GetExtension = ".txt"
    End Select
End Function

' Creates the folder, and the folders above it when needed.
Public Sub EnsureFolderExists(folderPath As String)
Dim fso As Object
Dim parentPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        parentPath = fso.GetParentFolderName(folderPath)
        If Not fso.FolderExists(parentPath) Then EnsureFolderExists parentPath
        fso.CreateFolder folderPath
    End If
End Sub

' The text with every character that is not allowed in a file name replaced by an underscore.
Public Function SanitizeFilename(ByVal name As String) As String
Dim ch As Variant

    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        name = Replace(name, ch, "_")
    Next ch
    SanitizeFilename = name
End Function

' The text of an exported module without the lines that only the file format needs: the class header (VERSION 1.0 CLASS ... END) and
' the Attribute lines. The VBA editor refuses them when the text is added to a module.
Private Function ReadSourceWithoutHeader(path As String) As String
Dim fNum As Integer
Dim content As String
Dim lines() As String
Dim i As Long
Dim result As String
Dim inHeader As Boolean

    fNum = FreeFile
    Open path For Input As #fNum
    content = Input$(LOF(fNum), fNum)
    Close #fNum
    content = Replace(Replace(content, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(content, vbLf)
    For i = 0 To UBound(lines)
        If i = 0 And Left(lines(i), 8) = "VERSION " Then
            inHeader = True
        ElseIf inHeader Then
            If Trim(lines(i)) = "END" Then inHeader = False
        ElseIf Left(lines(i), 10) <> "Attribute " Then
            result = result & lines(i) & vbCrLf
        End If
    Next i
    ReadSourceWithoutHeader = result
End Function

Private Function SourceFolder(wb As Workbook) As String
    If Len(wb.Path) = 0 Then Err.Raise 5, MODULE_NAME, "The workbook has not been saved yet; there is no folder for Src."
    SourceFolder = omOnedriveFunctions.GetWorkbookLocalPath(wb) & "\Src"
End Function
