Attribute VB_Name = "omDevelopmentFunctions"
Option Compare Database
Option Explicit
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

Public Sub ReportReplaceBackColor(ReportName As String, OldColor As Long, NewColor As Long)
Dim ctrl As Control
Dim InvalidProperty As Boolean

' ReportReplaceBackColor "SCreditnote",1384554,472939
' ReportReplaceBackColor "SCreditnote_HulchiBelluni2002",1384554,472939
' ReportReplaceBackColor "SCreditnote_ProductCosts",1384554,472939
' ReportReplaceBackColor "SCreditnote_Sub",1384554,472939
' ReportReplaceBackColor "SCreditnote_Sub_HulchiBelluni2002",1384554,472939
' ReportReplaceBackColor "SInvoice",1384554,472939
' ReportReplaceBackColor "SInvoice_HulchiBelluni2002",1384554,472939
' ReportReplaceBackColor "SInvoice_ProductCosts",1384554,472939
' ReportReplaceBackColor "SInvoice_Sub",1384554,472939
' ReportReplaceBackColor "SInvoice_Sub_HulchiBelluni2002",1384554,472939
' ReportReplaceBackColor "SOrder_Outstanding",1384554,472939
' ReportReplaceBackColor "SOrder_Print",1384554,472939
' ReportReplaceBackColor "SOrder_Print_Ordered",1384554,472939
' ReportReplaceBackColor "SOrderDetail_Print",1384554,472939
' ReportReplaceBackColor "SOrderDetail_Print_Ordered",1384554,472939
' ReportReplaceBackColor "SOrderProduct_Print",1384554,472939
' ReportReplaceBackColor "SOrderProductCost_Print",1384554,472939
' ReportReplaceBackColor "SPricelistProduct_Print",1384554,472939
' ReportReplaceBackColor "SPricelistProductFull_Print",1384554,472939
' ReportReplaceBackColor "SPricelistProductFullProductCost_Print",1384554,472939
' ReportReplaceBackColor "SPriceListProductPrice_Print",1384554,472939
' ReportReplaceBackColor "SPriceListProductPrice2_Print",1384554,472939
' ReportReplaceBackColor "SPricelistProductProductCost_Print",1384554,472939

    On Error GoTo ReportReplaceBackColor_Error
    DoCmd.OpenReport ReportName, acViewDesign
    For Each ctrl In Reports(ReportName).Controls
        InvalidProperty = False
        If ctrl.BackColor = OldColor Then
            If Not InvalidProperty Then
                ctrl.BackColor = NewColor
            End If
        End If
    Next
    DoCmd.Close acReport, ReportName, acSaveYes
    Exit Sub
ReportReplaceBackColor_Error:
    InvalidProperty = True
    Resume Next
End Sub

Public Sub ReportReplaceForeColor(ReportName As String, OldColor As Long, NewColor As Long)
Dim ctrl As Control
Dim InvalidProperty As Boolean

' ReportReplaceForeColor "SOrder_Print",1384554,472939

    On Error GoTo ReportReplaceForeColor_Error
    DoCmd.OpenReport ReportName, acViewDesign
    For Each ctrl In Reports(ReportName).Controls
        InvalidProperty = False
        If ctrl.ForeColor = OldColor Then
            If Not InvalidProperty Then
                ctrl.ForeColor = NewColor
            End If
        End If
    Next
    DoCmd.Close acReport, ReportName, acSaveYes
    Exit Sub
ReportReplaceForeColor_Error:
    InvalidProperty = True
    Resume Next
End Sub
Public Sub ReportReplaceColors()

    ReportReplaceBackColor "SCreditnote", 1384554, 472939
    ReportReplaceBackColor "SCreditnote_HulchiBelluni2002", 1384554, 472939
    ReportReplaceBackColor "SCreditnote_ProductCosts", 1384554, 472939
    ReportReplaceBackColor "SCreditnote_Sub", 1384554, 472939
    ReportReplaceBackColor "SCreditnote_Sub_HulchiBelluni2002", 1384554, 472939
    ReportReplaceBackColor "SInvoice", 1384554, 472939
    ReportReplaceBackColor "SInvoice_HulchiBelluni2002", 1384554, 472939
    ReportReplaceBackColor "SInvoice_ProductCosts", 1384554, 472939
    ReportReplaceBackColor "SInvoice_Sub", 1384554, 472939
    ReportReplaceBackColor "SInvoice_Sub_HulchiBelluni2002", 1384554, 472939
    ReportReplaceBackColor "SOrder_Outstanding", 1384554, 472939
    ReportReplaceBackColor "SOrder_Print", 1384554, 472939
    ReportReplaceBackColor "SOrder_Print_Ordered", 1384554, 472939
    ReportReplaceBackColor "SOrderDetail_Print", 1384554, 472939
    ReportReplaceBackColor "SOrderDetail_Print_Ordered", 1384554, 472939
    ReportReplaceBackColor "SOrderProduct_Print", 1384554, 472939
    ReportReplaceBackColor "SOrderProductCost_Print", 1384554, 472939
    ReportReplaceBackColor "SPricelistProduct_Print", 1384554, 472939
    ReportReplaceBackColor "SPricelistProductFull_Print", 1384554, 472939
    ReportReplaceBackColor "SPricelistProductFullProductCost_Print", 1384554, 472939
    ReportReplaceBackColor "SPriceListProductPrice_Print", 1384554, 472939
    ReportReplaceBackColor "SPriceListProductPrice2_Print", 1384554, 472939
    ReportReplaceBackColor "SPricelistProductProductCost_Print", 1384554, 472939

    ReportReplaceForeColor "SCreditnote", 1384554, 472939
    ReportReplaceForeColor "SCreditnote_HulchiBelluni2002", 1384554, 472939
    ReportReplaceForeColor "SCreditnote_ProductCosts", 1384554, 472939
    ReportReplaceForeColor "SCreditnote_Sub", 1384554, 472939
    ReportReplaceForeColor "SCreditnote_Sub_HulchiBelluni2002", 1384554, 472939
    ReportReplaceForeColor "SInvoice", 1384554, 472939
    ReportReplaceForeColor "SInvoice_HulchiBelluni2002", 1384554, 472939
    ReportReplaceForeColor "SInvoice_ProductCosts", 1384554, 472939
    ReportReplaceForeColor "SInvoice_Sub", 1384554, 472939
    ReportReplaceForeColor "SInvoice_Sub_HulchiBelluni2002", 1384554, 472939
    ReportReplaceForeColor "SOrder_Outstanding", 1384554, 472939
    ReportReplaceForeColor "SOrder_Print", 1384554, 472939
    ReportReplaceForeColor "SOrder_Print_Ordered", 1384554, 472939
    ReportReplaceForeColor "SOrderDetail_Print", 1384554, 472939
    ReportReplaceForeColor "SOrderDetail_Print_Ordered", 1384554, 472939
    ReportReplaceForeColor "SOrderProduct_Print", 1384554, 472939
    ReportReplaceForeColor "SOrderProductCost_Print", 1384554, 472939
    ReportReplaceForeColor "SPricelistProduct_Print", 1384554, 472939
    ReportReplaceForeColor "SPricelistProductFull_Print", 1384554, 472939
    ReportReplaceForeColor "SPricelistProductFullProductCost_Print", 1384554, 472939
    ReportReplaceForeColor "SPriceListProductPrice_Print", 1384554, 472939
    ReportReplaceForeColor "SPriceListProductPrice2_Print", 1384554, 472939
    ReportReplaceForeColor "SPricelistProductProductCost_Print", 1384554, 472939

End Sub

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
