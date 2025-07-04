Attribute VB_Name = "omLibraryFunctions"
Option Compare Database
Option Explicit

Private Function InternalGetTimeStamp(Optional prefix As String = "_") As String
    InternalGetTimeStamp = prefix & format(Now, "yyyymmdd_hhnnss")
End Function
Public Function ExportLibrary(Optional createFolder As Boolean = False, Optional startsWith As String = "om", Optional filenameWithTimestamp As Boolean = False) As String
Dim c As VBComponent
Dim sfx As String
Dim ts As String
Dim fn As String
Dim path As String

    ts = InternalGetTimeStamp("")
    path = CurrentProject.path
    If createFolder Then
        path = gFso.BuildPath(path, ts)
        If Not gFso.FolderExists(path) Then
            gFso.createFolder path
        End If
    End If
    ExportLibrary = path
    For Each c In Application.VBE.VBProjects(1).VBComponents
        Select Case c.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                sfx = ".cls"
            Case vbext_ct_MSForm
                sfx = ".frm"
            Case vbext_ct_StdModule
                sfx = ".bas"
            Case Else
                sfx = ""
        End Select
        If sfx <> "" And StrComp(Left(c.Name, Len(startsWith)), startsWith, vbBinaryCompare) = 0 Then
            fn = path & "\" & c.Name & IIf(filenameWithTimestamp, "_" & ts, "") & sfx
            c.Export filename:=fn
        End If
    Next c

End Function

Public Sub UpdateLibrary(Optional backupComponents As Boolean = True, Optional useUpdatesFolder As Boolean = True, Optional deleteAfterUpdate As Boolean = True)
Dim c As VBComponent
Dim cNew As VBComponent
Dim sfx As String
Dim fn As String
Dim path As String
Dim oName As String
Dim OType As AcObjectType

    ExportLibrary True
    path = CurrentProject.path
    If useUpdatesFolder Then
        path = gFso.BuildPath(path, "Updates")
        If Not gFso.FolderExists(path) Then
            MsgBox "Updates Folder does not exist: " & path
            Exit Sub
        End If
    End If
    For Each c In Application.VBE.VBProjects(1).VBComponents
        Select Case c.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                sfx = ".cls"
                OType = acModule
            Case vbext_ct_MSForm
                sfx = ".frm"
                OType = acForm
            Case vbext_ct_StdModule
                sfx = ".bas"
                OType = acModule
            Case Else
                sfx = ""
        End Select
        If sfx <> "" And StrComp(Left(c.Name, 2), "om", vbBinaryCompare) = 0 Then
            oName = c.Name
            fn = path & "\" & oName & sfx
            If gFso.FileExists(fn) Then
                Set cNew = Application.VBE.VBProjects(1).VBComponents.import(fn)
                'DoCmd.DeleteObject oType, oName
                'cNew.Name = oName
                'DoCmd.Save oType, objectname:=oName
                If deleteAfterUpdate Then
                    gFso.DeleteFile fn
                End If
            End If
        End If
    Next c

End Sub

' RemoveEqualFiles "C:\Users\JaRa\OneDrive\Shared\Kybucs\LibraryCompare\","C:\Users\JaRa\OneDrive\Shared\Kybucs\Library\"

Public Sub RemoveEqualFiles(sourcePath As String, destinationPath As String, Optional moveFile As Boolean = False)
Dim sFolder As Scripting.folder
Dim sf As File
Dim DF As File
Dim sourceFound As Boolean
Dim i As Long
Dim destinationFilename As String
Dim sString As String
Dim dString As String

    For Each DF In gFso.GetFolder(destinationPath).Files
        destinationFilename = omFileFunctions.RemoveExtension(DF.Name)
        i = 1
        sourceFound = False
        Set sFolder = gFso.GetFolder(sourcePath)

        For Each sf In sFolder.Files
            If InStr(1, sf.Name, destinationFilename) = 1 Then
                'dString = omFileFunctions.ReadFileToString(DF.path)
                'sString = omFileFunctions.ReadFileToString(sf.path)
                'sString = LCase(omStringFunctions.CleanStringUsingPattern(sString))
                'dString = LCase(omStringFunctions.CleanStringUsingPattern(dString))
                If dString = sString Then
                    sf.Delete
                Else
                    If moveFile Then
                        sf.Move gFso.BuildPath(destinationPath, sf.Name)
                    End If
                End If
                Exit For
            End If
        Next
    Next
End Sub

Public Sub LoadFromText(objectType As AcObjectType, objectName As String, filename As String)
    'Application.LoadFromText AcObjectType.acModule, "omLibraryFunctions", "\\sql01\data\ACenter\2007\ACenter_9002_JaRa_Updates\omLibraryFunctions.bas"
    Application.LoadFromText objectType, objectName, filename
End Sub
Public Sub LoadAsAXL(objectType As AcObjectType, objectName As String, filename As String)
    'Application.LoadFromText AcObjectType.acModule, "omLibraryFunctions", "\\sql01\data\ACenter\2007\ACenter_9002_JaRa_Updates\omLibraryFunctions.bas"
    MsgBox "Does not exist in version 2007"
    'Application.LoadAsAXL objectType, objectName, fileName
End Sub

Public Sub SaveAsText(objectType As AcObjectType, objectName As String, filename As String)
    'Application.SaveAsText  AcObjectType.acModule, "omLibraryFunctions", "\\sql01\data\ACenter\2007\ACenter_9002_JaRa_Updates\omLibraryFunctions.bas"
    Application.SaveAsText objectType, objectName, filename
End Sub
Public Sub SaveAsAXL(objectType As AcObjectType, objectName As String, filename As String)
    'Application.SaveAsText  AcObjectType.acModule, "omLibraryFunctions", "\\sql01\data\ACenter\2007\ACenter_9002_JaRa_Updates\omLibraryFunctions.bas"
    MsgBox "Does not exist in version 2007"
    'Application.SaveAsAXL objectType, objectName, fileName
End Sub

Public Sub ExportFormControlProperties(formName As String, Optional WriteToFile As Boolean = False, Optional controlEscaped As Boolean = False)
Dim F As Form
Dim c As Control
Dim S1 As String
Dim S2 As String
Dim cEscapedStart As String
Dim cEscapedEnd As String
Dim cName As String

    If controlEscaped Then
        cEscapedStart = "Controls(" & Chr(34)
        cEscapedEnd = Chr(34) & ")"
    End If
    DoCmd.OpenForm formName, acDesign, , , , acHidden
    Set F = Forms(formName)
    S1 = "Public Sub SetDefaultControlProperties()" & vbCrLf
    S2 = "Public Sub SetMinimumControlProperties(minTop as long, minLeft as long, minWidth as long, minHeight as long)" & vbCrLf
    For Each c In F.Controls
        S1 = S1 & "' ------ " & c.Name & vbCrLf
        cName = cEscapedStart & c.Name & cEscapedEnd
        S1 = S1 & "me." & cName & ".Visible=" & c.visible & vbCrLf
        S1 = S1 & "me." & cName & ".Top=" & c.Top & vbCrLf
        S1 = S1 & "me." & cName & ".Left=" & c.Left & vbCrLf
        S1 = S1 & "me." & cName & ".Width=" & c.width & vbCrLf
        S1 = S1 & "me." & cName & ".Height=" & c.height & vbCrLf

        S2 = S2 & "' ------ " & c.Name & vbCrLf
        S2 = S2 & "me." & cName & ".Visible=false" & vbCrLf
        S2 = S2 & "me." & cName & ".Top=minTop" & vbCrLf
        S2 = S2 & "me." & cName & ".Left=minLeft" & vbCrLf
        S2 = S2 & "me." & cName & ".Width=minWidth" & vbCrLf
        S2 = S2 & "me." & cName & ".Height=minHeight" & vbCrLf
    Next
    S1 = S1 & "End Sub" & vbCrLf
    S2 = S2 & "End Sub" & vbCrLf
    If WriteToFile Then
        omFileFunctions.WriteStringToFile gFso.BuildPath(CurrentProject.path, "ExportFormControlProperties_" & formName & InternalGetTimeStamp()) & ".txt", S1 & vbCrLf & S2
    Else
        Debug.Print S1 & vbCrLf & S2
    End If
    DoCmd.Close acForm, formName, acSaveNo
End Sub

Public Sub Dump(Optional objectType As AcObjectType = AcObjectType.acDefault, Optional silentMode As Boolean = True, Optional destinationPath As String, Optional addTimeStamp As Boolean = True)
Dim path As String
Dim O As Object
Dim currentPath As String
Dim rs As New ADODB.Recordset

    If Len(Trim(Nz(destinationPath, ""))) <> 0 Then
        path = destinationPath
    Else
        path = gFso.BuildPath(CurrentProject.path, "Dump")
    End If
    If addTimeStamp Then
        path = IIf(Right(path, 1) = "\", Left(path, Len(path) - 1), path) & InternalGetTimeStamp()
    End If

    omFileFunctions.CreateFolderPath path

    If objectType = acTable Or objectType = acDefault Then
        currentPath = gFso.BuildPath(path, "Tables")
        omFileFunctions.DeleteFolder currentPath
        omFileFunctions.CreateFolderPath currentPath
        rs.Open "SELECT T.Name FROM MSysObjects T WHERE T.Type=1 AND T.Flags=0", CurrentProject.connection, adOpenForwardOnly, adLockReadOnly
        While Not rs.EOF
            Application.ExportXML acExportTable, rs("Name"), gFso.BuildPath(currentPath, rs("Name") & ".xml"), gFso.BuildPath(currentPath, rs("Name") & "_Schema.xml"), , , , acExportAllTableAndFieldProperties + acEmbedSchema
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
    End If

    If objectType = acQuery Or objectType = acDefault Then
        currentPath = gFso.BuildPath(path, "Queries")
        omFileFunctions.DeleteFolder currentPath
        omFileFunctions.CreateFolderPath currentPath
        For Each O In Application.CodeData.AllQueries
            omLibraryFunctions.SaveAsText acQuery, O.Name, gFso.BuildPath(currentPath, O.Name & ".txt")
        Next
    End If

    If objectType = acForm Or objectType = acDefault Then
        currentPath = gFso.BuildPath(path, "Forms")
        omFileFunctions.DeleteFolder currentPath
        omFileFunctions.CreateFolderPath currentPath
        For Each O In Application.CodeProject.AllForms
            omLibraryFunctions.SaveAsText acForm, O.Name, gFso.BuildPath(currentPath, O.Name & ".txt")
        Next
    End If

    If objectType = acReport Or objectType = acDefault Then
        currentPath = gFso.BuildPath(path, "Reports")
        omFileFunctions.DeleteFolder currentPath
        omFileFunctions.CreateFolderPath currentPath
        For Each O In Application.CodeProject.AllReports
            omLibraryFunctions.SaveAsText acReport, O.Name, gFso.BuildPath(currentPath, O.Name & ".txt")
        Next
    End If

    If objectType = acModule Or objectType = acDefault Then
        currentPath = gFso.BuildPath(path, "Modules")
        omFileFunctions.DeleteFolder currentPath
        omFileFunctions.CreateFolderPath currentPath
        For Each O In Application.CodeProject.AllModules
            omLibraryFunctions.SaveAsText acModule, O.Name, gFso.BuildPath(currentPath, O.Name & ".txt")
        Next
    End If

    If objectType = acMacro Or objectType = acDefault Then
        currentPath = gFso.BuildPath(path, "Macros")
        omFileFunctions.DeleteFolder currentPath
        omFileFunctions.CreateFolderPath currentPath
        For Each O In Application.CodeProject.AllMacros
            omLibraryFunctions.SaveAsText acMacro, O.Name, gFso.BuildPath(currentPath, O.Name & ".txt")
        Next
    End If

    If Not silentMode Then
        MsgBox "Completed", vbOKOnly
    End If
End Sub

Public Sub DumpImport(objectType As AcObjectType, Optional silentMode As Boolean = True, Optional sourcePath As String)
Dim path As String
Dim O As Object
Dim currentPath As String
Dim fso As New FileSystemObject
Dim F As File
Dim i As Long
Dim tableName As String

    If Len(Trim(Nz(sourcePath, ""))) <> 0 Then
        path = sourcePath
    Else
        path = gFso.BuildPath(CurrentProject.path, "Dump")
    End If

    Select Case objectType
        Case acTable
            currentPath = fso.BuildPath(path, "Tables")
            For Each F In fso.GetFolder(currentPath).Files
                If InStr(1, F.Name, "_Schema") = 0 Then
                    If Not silentMode Then
                        Debug.Print "start : " & F.Name
                        DoEvents
                    End If
                    tableName = Left(F.Name, Len(F.Name) - 4)
                    For i = 0 To CurrentData.AllTables.Count - 1
                        If CurrentData.AllTables(i).Name = tableName Then
                            DoCmd.Rename tableName & "_" & format(Now, "yyyymmdd_hhnnss"), acTable, tableName
                        End If
                    Next
                    Application.ImportXML F.path, acStructureAndData
                End If
            Next
        Case acModule
            currentPath = fso.BuildPath(path, "Modules")
            For Each F In fso.GetFolder(currentPath).Files
                If F.Name <> "omLibraryFunctions.txt" Then
                    If Not silentMode Then
                        Debug.Print "start : " & F.Name
                        DoEvents
                    End If
                    omLibraryFunctions.LoadFromText acModule, Left(F.Name, Len(F.Name) - 4), F.path
                End If
            Next
        Case acQuery
            currentPath = fso.BuildPath(path, "Queries")
            For Each F In fso.GetFolder(currentPath).Files
                If Not silentMode Then
                    Debug.Print "start : " & F.Name
                    DoEvents
                End If
                omLibraryFunctions.LoadFromText acQuery, Left(F.Name, Len(F.Name) - 4), F.path
            Next
        Case acForm
            currentPath = fso.BuildPath(path, "Forms")
            For Each F In fso.GetFolder(currentPath).Files
                If Not silentMode Then
                    Debug.Print "start : " & F.Name
                    DoEvents
                End If
                omLibraryFunctions.LoadFromText acForm, Left(F.Name, Len(F.Name) - 4), F.path
            Next
        Case acReport
            currentPath = fso.BuildPath(path, "Reports")
            For Each F In fso.GetFolder(currentPath).Files
                If Not silentMode Then
                    Debug.Print "start : " & F.Name
                    DoEvents
                End If
                omLibraryFunctions.LoadFromText acReport, Left(F.Name, Len(F.Name) - 4), F.path
            Next
        Case acMacro
            currentPath = fso.BuildPath(path, "Macros")
            For Each F In fso.GetFolder(currentPath).Files
                If Not silentMode Then
                    Debug.Print "start : " & F.Name
                    DoEvents
                End If
                omLibraryFunctions.LoadFromText acMacro, Left(F.Name, Len(F.Name) - 4), F.path
            Next
    End Select
    If Not silentMode Then
        MsgBox "Completed", vbOKOnly
    End If
End Sub

Public Sub DeleteByObjectType(objectType As AcObjectType, Optional silentMode As Boolean = True, Optional maxNumberOfObjects = 0)
Dim cnt As Long
Dim cntTo As Long
Dim i As Long

    Select Case objectType
        Case acModule
            cnt = Application.CodeProject.AllModules.Count - 1
            For i = cnt To 0 Step -1
                DoCmd.DeleteObject objectType, Application.CodeProject.AllModules(i).Name
            Next
        Case acQuery
            cnt = Application.CodeData.AllQueries.Count - 1
            For i = cnt To 0 Step -1
                DoCmd.DeleteObject objectType, Application.CodeData.AllQueries(i).Name
            Next
        Case acForm
            cnt = Application.CodeProject.AllForms.Count - 1
            cntTo = IIf(maxNumberOfObjects = 0, 0, cnt - maxNumberOfObjects)
            cntTo = IIf(cntTo < 0, 0, cntTo)
            For i = cnt To cntTo Step -1
                DoCmd.DeleteObject objectType, Application.CodeProject.AllForms(i).Name
            Next
        Case acReport
            cnt = Application.CodeProject.AllReports.Count - 1
            cntTo = IIf(maxNumberOfObjects = 0, 0, cnt - maxNumberOfObjects)
            cntTo = IIf(cntTo < 0, 0, cntTo)
            For i = cnt To cntTo Step -1
                DoCmd.DeleteObject objectType, Application.CodeProject.AllReports(i).Name
            Next
        Case acMacro
            cnt = Application.CodeProject.AllMacros.Count - 1
            For i = cnt To 0 Step -1
                DoCmd.DeleteObject objectType, Application.CodeProject.AllMacros(i).Name
            Next
    End Select
    If Not silentMode Then
        MsgBox "Completed", vbOKOnly
    End If
End Sub

Public Sub VBACountModules()
Dim cnt As Long
    cnt = VBE.ActiveVBProject.VBComponents.Count
    MsgBox "Access Project contains #" & cnt & " modules.", vbOKOnly
    ' 20200512 : 1176 modules
End Sub

Public Sub VBACountLines()
Dim cnt As Long
Dim F As VBIDE.VBComponent


    For Each F In VBE.ActiveVBProject.VBComponents
        cnt = cnt + F.CodeModule.CountOfLines
    Next
    MsgBox "Access Project contains #" & cnt & " lines.", vbOKOnly
End Sub
Public Sub VBAFindReplace(findString As String, replaceString As String, Optional silentMode As Boolean = True)
Dim c As VBIDE.VBComponent
Dim i As Long
Dim s As String

    For Each c In VBE.ActiveVBProject.VBComponents
        For i = 1 To c.CodeModule.CountOfLines
            s = c.CodeModule.Lines(i, 1)
            If InStr(1, s, findString) > 0 Then
                If Not silentMode Then
                    Debug.Print "In module: " & c.Name & " -> replaced: " & s
                End If
                c.CodeModule.ReplaceLine i, Replace(s, findString, replaceString)
            End If
        Next i
    Next
End Sub

Public Sub ListOnlyModulesWithoutLines()
Dim F As VBComponent

    For Each F In VBE.ActiveVBProject.VBComponents
        If F.CodeModule.CountOfLines = F.CodeModule.CountOfDeclarationLines Then
            Debug.Print F.Name
        End If
    Next
End Sub

Public Sub RemoveModulesWithoutLines(Optional silentMode As Boolean = True)
Dim F As VBIDE.VBComponent
Dim s As String
Dim i As Long
Dim msg As String
Dim msgBoxResult As VbMsgBoxResult

    For i = VBE.ActiveVBProject.VBComponents.Count To 1 Step -1
        Set F = VBE.ActiveVBProject.VBComponents(i)
        If F.CodeModule.CountOfLines = F.CodeModule.CountOfDeclarationLines And F.CodeModule.CountOfLines < 3 Then
            msg = F.Name
            If F.CodeModule.CountOfLines > 0 Then
                msg = msg & vbCrLf & F.CodeModule.Lines(1, F.CodeModule.CountOfLines)
            End If
            'Debug.Print msg

            If Not silentMode Then
                msgBoxResult = MsgBox(msg & vbCrLf & vbCrLf & "Delete CodeModule for " & F.Name, vbYesNo)
            Else
                msgBoxResult = vbYes
            End If
            If msgBoxResult = vbYes Then
                If F.Type = vbext_ct_ClassModule Or F.Type = vbext_ct_StdModule Then
                    VBE.ActiveVBProject.VBComponents.Remove F
                ElseIf Left(F.Name, Len("Report_")) = "Report_" Then
                    s = Mid(F.Name, Len("Report_") + 1)
                    DoCmd.OpenReport s, acDesign
                    Reports(s).HasModule = False
                    DoCmd.Close acReport, s, acSaveYes
                ElseIf Left(F.Name, Len("Form_")) = "Form_" Then
                    s = Mid(F.Name, Len("Form_") + 1)
                    DoCmd.OpenForm s, acDesign
                    Forms(s).HasModule = False
                    DoCmd.Close acForm, s, acSaveYes
                End If
            End If
        End If
    Next
End Sub
Public Sub DumpAndCommitAndPush(Optional objectType As AcObjectType = AcObjectType.acDefault)
Dim path As String
Dim msg As String

    path = gFso.BuildPath(omFileFunctions.GetUserRootFolder, "Source\Repos\ACenter\")
    omLibraryFunctions.Dump objectType, True, path, False
    omLibraryFunctions.CommitAndPush path

End Sub
Public Sub CommitAndPush(repositoryPath As String)
Dim path As String
Dim fn As String
Dim s As String
Dim commitMessage As String
Dim gitFolder As String

    gitFolder = "C:\Program Files\Git\"
    If gFso.FolderExists(gitFolder) = False Then
        gitFolder = gFso.BuildPath(omFileFunctions.GetUserRootFolder, "AppData\Local\Programs\Git\")
    End If
    commitMessage = InputBox("Geef een omschrijving")
    If Len(commitMessage) = 0 Then
        MsgBox "Not message was given. Execution is aborted!", vbOKOnly
        Exit Sub
    End If
    fn = gFso.BuildPath(repositoryPath, ".git\COMMITMESSAGE")
    omFileFunctions.WriteStringToFile fn, commitMessage
    s = "cd '{{repositoryPath}}'"
    s = s & vbCrLf & "'{{GitFolder}}bin\git.exe' add -A"
    s = s & vbCrLf & "'{{GitFolder}}bin\git.exe' commit -F '{{FilenameCommitMessage}}'"
    s = s & vbCrLf & "'{{GitFolder}}bin\git.exe' push"
    s = Replace(s, "{{repositoryPath}}", repositoryPath)
    s = Replace(s, "{{GitFolder}}", gitFolder)
    s = Replace(s, "{{FilenameCommitMessage}}", fn)
    s = Replace(s, "'", Chr(34))
    fn = gFso.BuildPath(CurrentProject.path, CurrentProject.Name & "_CommitAndPush.bat")
    omFileFunctions.WriteStringToFile fn, s
    Shell fn
End Sub

Public Sub RemoveDeletedObjects()
Dim rs As New ADODB.Recordset
Dim objType As AcObjectType

    rs.Open "SELECT Name,Type FROM MSysobjects WHERE Name LIKE '~%'", CurrentProject.connection, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        objType = -1
        Select Case rs("Type")
            Case 4
                objType = acTable
            Case 5
                objType = acQuery
            Case -32766
                objType = acMacro
            Case Else
                objType = 0
                MsgBox "ObjectType Not Defined " & rs("type")
        End Select
        If objType > -1 Then
            DoCmd.DeleteObject objType, rs("Name")
        End If
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
End Sub

Public Sub ListFormsToTable(Optional tableName As String = "___T_Forms")
Dim rs As New ADODB.Recordset
Dim F As AccessObject

    rs.Open tableName, CurrentProject.connection, adOpenForwardOnly, adLockOptimistic
    For Each F In CodeProject.AllForms
        rs.AddNew
        rs("Name") = F.Name
        rs.Update
    Next
    rs.Close
    Set rs = Nothing
End Sub

Public Sub CreateBinaryColumns()

    Dim td As dao.TableDef
    Dim db As dao.Database
    Dim fd As Field

    Set db = CurrentDb
    Set td = db.TableDefs("Table1")

    Set fd = td.CreateField("BinaryColumn", DataTypeEnum.dbBinary, 100)
    fd.Attributes = fd.Attributes Or dbFixedField
    td.Fields.Append fd

    Set fd = td.CreateField("VarbinaryColumn", DataTypeEnum.dbBinary, 100)
    td.Fields.Append fd

    td.Fields.Refresh
    db.TableDefs.Refresh
End Sub

Public Function GenerateDataMacroAfterInsert(tableName As String) As String
Dim s As String
Dim statementActionTemplate As String
Dim statements As String
Dim F As Field
Dim cdb As dao.Database

    statementActionTemplate = "<Action Name='SetField'>" & vbCrLf
    statementActionTemplate = statementActionTemplate + "<Argument Name='Field'><!FieldName!></Argument>" & vbCrLf
    statementActionTemplate = statementActionTemplate + "<Argument Name='Value'>[<!TableName!>].[<!FieldName!>]</Argument>" & vbCrLf
    statementActionTemplate = statementActionTemplate + "</Action>"
    statementActionTemplate = Replace(statementActionTemplate, "'", Chr(34))

    s = "<?xml version='1.0' encoding='UTF-16' standalone='no'?>" & vbCrLf
    s = s + "<DataMacros xmlns='http://schemas.microsoft.com/office/accessservices/2009/11/application'>" & vbCrLf
    s = s + "<DataMacro Event='AfterInsert'>" & vbCrLf
    s = s + "<Statements>" & vbCrLf
    s = s + "<CreateRecord>" & vbCrLf
    s = s + "<Data>" & vbCrLf
    s = s + "<Reference>Log_<!TableName!></Reference>" & vbCrLf
    s = s + "</Data>" & vbCrLf
    s = s + "<Statements>" & vbCrLf
    s = s + "<!Statements!>"
    s = s + "</Statements>" & vbCrLf
    s = s + "</CreateRecord>" & vbCrLf
    s = s + "</Statements>" & vbCrLf
    s = s + "</DataMacro>" & vbCrLf
    s = s + "</DataMacros>"
    s = Replace(s, "'", Chr(34))
    s = Replace(s, "<!TableName!>", tableName)
    Set cdb = CurrentDb
    For Each F In cdb.TableDefs(tableName).Fields
        statements = statements + Replace(Replace(statementActionTemplate, "<!FieldName!>", F.Name), "<!TableName!>", tableName) & vbCrLf
    Next

    GenerateDataMacroAfterInsert = Replace(s, "<!Statements!>", statements)
End Function
Public Function GenerateDataMacroAfterUpdate(tableName As String) As String
Dim s As String
Dim statementActionTemplate As String
Dim statements As String
Dim F As Field
Dim cdb As dao.Database

    statementActionTemplate = "<Action Name='SetField'>" & vbCrLf
    statementActionTemplate = statementActionTemplate + "<Argument Name='Field'><!FieldName!></Argument>" & vbCrLf
    statementActionTemplate = statementActionTemplate + "<Argument Name='Value'>[<!TableName!>].[<!FieldName!>]</Argument>" & vbCrLf
    statementActionTemplate = statementActionTemplate + "</Action>"
    statementActionTemplate = Replace(statementActionTemplate, "'", Chr(34))

    s = "<?xml version='1.0' encoding='UTF-16' standalone='no'?>" & vbCrLf
    s = s + "<DataMacros xmlns='http://schemas.microsoft.com/office/accessservices/2009/11/application'>" & vbCrLf
    s = s + "<DataMacro Event='AfterUpdate'>" & vbCrLf
    s = s + "<Statements>" & vbCrLf
    s = s + "<CreateRecord>" & vbCrLf
    s = s + "<Data>" & vbCrLf
    s = s + "<Reference>Log_<!TableName!></Reference>" & vbCrLf
    s = s + "</Data>" & vbCrLf
    s = s + "<Statements>" & vbCrLf
    s = s + "<!Statements!>"
    s = s + "</Statements>" & vbCrLf
    s = s + "</CreateRecord>" & vbCrLf
    s = s + "</Statements>" & vbCrLf
    s = s + "</DataMacro>" & vbCrLf
    s = s + "</DataMacros>"
    s = Replace(s, "'", Chr(34))
    s = Replace(s, "<!TableName!>", tableName)
    Set cdb = CurrentDb
    For Each F In cdb.TableDefs(tableName).Fields
        statements = statements + Replace(Replace(statementActionTemplate, "<!FieldName!>", F.Name), "<!TableName!>", tableName) & vbCrLf
    Next

    GenerateDataMacroAfterUpdate = Replace(s, "<!Statements!>", statements)
End Function
