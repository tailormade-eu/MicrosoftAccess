Attribute VB_Name = "omDevelopmentFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=984ff6b4
Public Sub ExportQueries()
Dim rsQueries As New ADODB.Recordset
Dim fno As Long

    fno = FreeFile
    Open "Queries.txt" For Output As fno
    rsQueries.Open "SELECT Name FROM MSysObjects WHERE Type = 5", CurrentProject.connection, adOpenForwardOnly, adLockReadOnly
    While Not rsQueries.EOF
        Print #fno, "Create Procedure " & Chr(34) & rsQueries("Name") & Chr(34) & vbCrLf
        Print #fno, "As" & vbCrLf
        Print #fno, DBEngine(0)(0).QueryDefs(rsQueries("Name")).SQL & vbCrLf
        rsQueries.MoveNext
    Wend
    rsQueries.Close
    Set rsQueries = Nothing
    Close fno
End Sub
Public Sub TemplateCopy(strFind As String, strReplace As String)
    QueryCopy strFind & "_List_Search", strFind, strReplace
    QueryCopy strFind & "_Select", strFind, strReplace
    DBEngine(0)(0).QueryDefs.Refresh
    FormCopy strFind & "_Edit", strFind, strReplace
    FormCopy strFind & "_List", strFind, strReplace
    FormCopy strFind & "_List_Search", strFind, strReplace
End Sub
Public Sub QueryCopy(strQuery As String, strFind As String, strReplace As String)
Dim strNewQuery As String
Dim strSQL As String

    strNewQuery = Replace(strQuery, GetEnglishPlural(strFind), GetEnglishPlural(strReplace))
    strNewQuery = Replace(strNewQuery, strFind, strReplace)
    DoCmd.CopyObject , strNewQuery, acQuery, strQuery
    DBEngine(0)(0).QueryDefs.Refresh
    strSQL = DBEngine(0)(0).QueryDefs(strNewQuery).SQL
    strSQL = Replace(strSQL, GetEnglishPlural(strFind), GetEnglishPlural(strReplace))
    strSQL = Replace(strSQL, strFind, strReplace)
    DBEngine(0)(0).QueryDefs(strNewQuery).SQL = strSQL
    DBEngine(0)(0).QueryDefs.Refresh
End Sub

Public Sub FormCopy(strForm As String, strFind As String, strReplace As String)
Dim strNewForm As String

    strNewForm = Replace(strForm, strFind, strReplace)
    DoCmd.CopyObject , strNewForm, acForm, strForm
    FormControlRename strNewForm, 0, GetEnglishPlural(strFind), GetEnglishPlural(strReplace)
    FormControlRename strNewForm, 0, strFind, strReplace

End Sub
Public Sub FormControlRename(strForm As String, lControlType As Long, strFind As String, strReplace As String, Optional OnlyCaption As Boolean = False, Optional TestRun As Boolean = False)
Dim bChanged As Boolean
Dim frmEdit As Form
Dim lStartLine As Long
Dim lStartCol As Long
Dim lEndLine As Long
Dim lEndCol As Long
Dim strFound As String
Dim lCount As Long
Dim ctl As Control
Dim i As Long

FormControlRename_Restart:
    DoCmd.OpenForm strForm, acDesign, , , , acHidden
    bChanged = False
    For i = 0 To Forms(strForm).Controls.Count - 1
        Set ctl = Forms(strForm).Controls(i)
        With ctl
            If .ControlType = lControlType Or lControlType = 0 Then
                Select Case .ControlType
                    Case acTextBox, acComboBox
                        If (InStr(1, .Name, strFind) > 0 Or InStr(1, .ControlSource, strFind) > 0) And InStr(1, .Name, strReplace) > 0 And InStr(1, .ControlSource, strReplace) > 0 Then
                            If TestRun Then
                                Debug.Print strForm, .Name, .ControlSource
                            Else
                                bChanged = True
                                .Name = Replace(.Name, strFind, strReplace)
                                .ControlSource = Replace(.ControlSource, strFind, strReplace)
                            End If
                        End If
                    Case acLabel, acCommandButton
                        If (InStr(1, .Name, strFind) > 0 Or InStr(1, .Caption, strFind) > 0) And InStr(1, .Caption, strReplace) = 0 And InStr(1, .Name, strReplace) = 0 Then
                            If TestRun Then
                                Debug.Print strForm, .Name, .Caption
                            Else
                                bChanged = True
                                If Not OnlyCaption Then
                                    .Name = Replace(.Name, strFind, strReplace)
                                End If
                                .Caption = Replace(.Caption, strFind, strReplace)
                            End If
                        End If
                    Case acSubform
                        If (InStr(1, .Name, strFind) > 0 Or InStr(1, .SourceObject, strFind) > 0) And InStr(1, .Name, strReplace) = 0 And InStr(1, .SourceObject, strReplace) = 0 Then
                            If TestRun Then
                                Debug.Print strForm, .Name, .SourceObject
                            Else
                                bChanged = True
                                .Name = Replace(.Name, strFind, strReplace)
                                .SourceObject = Replace(.SourceObject, strFind, strReplace)
                            End If
                        End If
                End Select
            End If
        End With
        'If bChanged Then
        '    Exit For
        'End If
    Next
    If bChanged Then
        DoCmd.Close acForm, strForm, acSaveYes
        GoTo FormControlRename_Restart
    End If
    If lControlType = 0 Then
        With Forms(strForm)
            .RecordSource = Replace(.RecordSource, strFind, strReplace)
            .Caption = Replace(.Caption, strFind, strReplace)
        End With
        Set frmEdit = Forms(strForm)
        While frmEdit.Module.Find(strFind, lStartLine, lStartCol, lEndLine, lEndCol)
            strFound = frmEdit.Module.Lines(lStartLine, 1)
            strFound = Replace(strFound, strFind, strReplace)
            frmEdit.Module.ReplaceLine lStartLine, strFound
            lStartLine = lEndLine + 1
            lStartCol = 0
            lEndLine = 0
            lEndCol = 0
        Wend
    End If
    DoCmd.Close acForm, strForm, acSaveYes
End Sub

Public Sub DeleteTempObjects()
Dim rs As New ADODB.Recordset

    rs.Open "SELECT Name, Type FROM MSysObjects WHERE Name Like '~%'", CurrentProject.connection, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        Debug.Print rs(0), rs(1)
        Select Case rs(1)
            Case 5
                DoCmd.DeleteObject acQuery, rs(0)
            Case -32764
                DoCmd.DeleteObject acReport, rs(0)
        End Select
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
End Sub

' Replaces one colour by another in a report or a form (design view, hidden): every control, and every
' section, whose colour property equals oldColor gets newColor.
'   colorProperties : the colour properties to look at, comma separated (a control without that property is skipped)
'   objectType      : acReport (default) or acForm
' Returns the number of properties that were changed. The object is only saved when something changed.
' An object that cannot be opened (wrong name, open in another view) raises its error after being closed.
' Example: ReplaceColorInObject "SInvoice", 1384554, 472939, "BackColor,ForeColor"
Public Function ReplaceColorInObject(objectName As String, oldColor As Long, newColor As Long, _
        Optional colorProperties As String = "BackColor,ForeColor,BorderColor", _
        Optional objectType As AcObjectType = acReport) As Long
Dim obj As Object
Dim props() As String
Dim changed As Long
Dim errNumber As Long
Dim errDescription As String

    props = Split(Replace(colorProperties, " ", ""), ",")
    On Error GoTo ReplaceColorInObject_Error
    If objectType = acForm Then
        DoCmd.OpenForm objectName, acDesign, windowMode:=acHidden
        Set obj = Forms(objectName)
    Else
        DoCmd.OpenReport objectName, acViewDesign, windowMode:=acHidden
        Set obj = Reports(objectName)
    End If
    changed = ReplaceColorInControls(obj.Controls, oldColor, newColor, props)
    changed = changed + ReplaceColorInSections(obj, oldColor, newColor, props)
    DoCmd.Close objectType, objectName, IIf(changed > 0, acSaveYes, acSaveNo)
    ReplaceColorInObject = changed
    Exit Function

ReplaceColorInObject_Error:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    DoCmd.Close objectType, objectName, acSaveNo
    On Error GoTo 0
    Err.Raise errNumber, "ReplaceColorInObject", errDescription
End Function

' The same for all reports (or forms) whose name matches namePattern (Like pattern, "*" = all).
' Returns the total number of properties changed.
' Example: ReplaceColorInAll 1384554, 472939, "BackColor,ForeColor", "S*_Print"
Public Function ReplaceColorInAll(oldColor As Long, newColor As Long, _
        Optional colorProperties As String = "BackColor,ForeColor,BorderColor", _
        Optional namePattern As String = "*", _
        Optional objectType As AcObjectType = acReport) As Long
Dim names As New Collection
Dim obj As AccessObject
Dim n As Variant

    If objectType = acForm Then
        For Each obj In CurrentProject.AllForms
            If obj.Name Like namePattern Then names.Add obj.Name
        Next obj
    Else
        For Each obj In CurrentProject.AllReports
            If obj.Name Like namePattern Then names.Add obj.Name
        Next obj
    End If
    For Each n In names
        ReplaceColorInAll = ReplaceColorInAll + ReplaceColorInObject(CStr(n), oldColor, newColor, colorProperties, objectType)
    Next n
End Function

' Short forms kept for existing callers: only the back colour / only the fore colour of a report.
Public Sub ReportReplaceBackColor(ReportName As String, OldColor As Long, NewColor As Long)
    ReplaceColorInObject ReportName, OldColor, NewColor, "BackColor"
End Sub

Public Sub ReportReplaceForeColor(ReportName As String, OldColor As Long, NewColor As Long)
    ReplaceColorInObject ReportName, OldColor, NewColor, "ForeColor"
End Sub

Private Function ReplaceColorInControls(ctrls As Object, oldColor As Long, newColor As Long, props() As String) As Long
Dim ctrl As Control

    For Each ctrl In ctrls
        ReplaceColorInControls = ReplaceColorInControls + ReplaceColorOfItem(ctrl, oldColor, newColor, props)
    Next ctrl
End Function

Private Function ReplaceColorInSections(obj As Object, oldColor As Long, newColor As Long, props() As String) As Long
Dim i As Long
Dim sec As Object

    ' 0 = detail, 1/2 = header/footer, 3/4 = page header/footer, 5 and up = group levels
    For i = 0 To 30
        Set sec = Nothing
        On Error Resume Next
        Set sec = obj.Section(i)
        On Error GoTo 0
        If Not sec Is Nothing Then
            ReplaceColorInSections = ReplaceColorInSections + ReplaceColorOfItem(sec, oldColor, newColor, props)
        End If
    Next i
End Function

' Sets every listed colour property of a control or section that equals oldColor to newColor.
Private Function ReplaceColorOfItem(item As Object, oldColor As Long, newColor As Long, props() As String) As Long
Dim i As Long
Dim current As Long

    For i = LBound(props) To UBound(props)
        If Len(props(i)) > 0 Then
            Err.Clear
            On Error Resume Next
            current = CallByName(item, props(i), VbGet)
            If Err.Number = 0 Then
                If current = oldColor Then
                    CallByName item, props(i), VbLet, newColor
                    If Err.Number = 0 Then ReplaceColorOfItem = ReplaceColorOfItem + 1
                End If
            End If
            On Error GoTo 0
        End If
    Next i
End Function

Public Sub AllFormsReplace(strFind As String, strReplace As String)
Dim i As Long

    For i = 0 To CurrentProject.AllForms.Count - 1
        'Debug.Print CurrentProject.AllForms(i).Name
        omDevelopmentFunctions.FormControlRename CurrentProject.AllForms(i).Name, acLabel, strFind, strReplace, True, False
        omDevelopmentFunctions.FormControlRename CurrentProject.AllForms(i).Name, acCommandButton, strFind, strReplace, True, False
    Next
End Sub

Public Sub AllFormsDisableDialog()
Dim i As Long

    For i = 0 To CurrentProject.AllForms.Count - 1
        DoCmd.OpenForm CurrentProject.AllForms(i).Name, acDesign, windowMode:=acHidden
        Forms(CurrentProject.AllForms(i).Name).Modal = False
        DoCmd.Close acForm, CurrentProject.AllForms(i).Name, acSaveYes
    Next
End Sub

Public Sub AllFormsDisablePopUp()
Dim i As Long

    For i = 0 To CurrentProject.AllForms.Count - 1
        DoCmd.OpenForm CurrentProject.AllForms(i).Name, acDesign, windowMode:=acHidden
        Forms(CurrentProject.AllForms(i).Name).Popup = False
        DoCmd.Close acForm, CurrentProject.AllForms(i).Name, acSaveYes
    Next
End Sub
