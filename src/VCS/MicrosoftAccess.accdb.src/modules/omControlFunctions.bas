Attribute VB_Name = "omControlFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=93b29458

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

#If VBA7 Then
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
#Else
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
#End If

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Public Sub IfNullOrEmpty(ctrl As Control, Optional UseDefaultValue As Boolean = True)
    If omStringFunctions.IsNullOrEmpty(ctrl) Then
        If UseDefaultValue And NotIsNullOrEmpty(ctrl.defaultValue) Then
            SetDefaultValue ctrl
        Else
            ctrl = ctrl.OldValue
        End If
    End If
End Sub

Public Sub SetDefaultValue(ctrl As Control)
Dim vDate As Variant

    If omStringFunctions.NotIsNullOrEmpty(ctrl) Then
        Exit Sub
    End If

    vDate = Replace(Replace(Nz(ctrl.defaultValue, ""), Chr(34), ""), "#", "")
    ctrl = IIf(IsDate(vDate), vDate, ctrl.defaultValue)
End Sub


' Measures the work area of the Access window (where forms open), in twips.
' False = measurement failed; the caller then leaves the windows where Access places them.
Public Function GetWorkArea(ByRef widthTwips As Long, ByRef heightTwips As Long) As Boolean
#If VBA7 Then
    Dim hMdi As LongPtr
    Dim hDC As LongPtr
#Else
    Dim hMdi As Long
    Dim hDC As Long
#End If
    Dim rc As RECT
    Dim dpiX As Long
    Dim dpiY As Long

    hMdi = FindWindowEx(Application.hWndAccessApp, 0, "MDIClient", vbNullString)
    If hMdi = 0 Then Exit Function
    If GetClientRect(hMdi, rc) = 0 Then Exit Function

    hDC = GetDC(0)
    dpiX = GetDeviceCaps(hDC, LOGPIXELSX)
    dpiY = GetDeviceCaps(hDC, LOGPIXELSY)
    ReleaseDC 0, hDC
    If dpiX = 0 Or dpiY = 0 Then Exit Function

    widthTwips = CLng(rc.Right * 1440# / dpiX)
    heightTwips = CLng(rc.Bottom * 1440# / dpiY)
    GetWorkArea = True
End Function

Public Function IsFormOpen(formName As String) As Boolean

On Error GoTo IsFormOpen_Error

    IsFormOpen = (Forms(formName).Name = formName)
    Exit Function
IsFormOpen_Error:
    IsFormOpen = False

End Function

Public Function GetOpenForm(formName As String) As Form

    If IsFormOpen(formName) Then
        Set GetOpenForm = Forms(formName)
    End If
End Function

Public Function NotInList(ctrl As Control, objectName As String, idFieldName As String, valueFieldName As String, NewData As String, Optional OpenEdit As Boolean = False, Optional ParentFieldName As String = "", Optional ParentData As Variant = Null) As Integer
Dim rs As New ADODB.Recordset
Dim Id As Long

    If MsgBox("Nieuw item: " & NewData & " toevoegen?", vbYesNo) = vbYes Then
        rs.Open omStringFunctions.GetEnglishPlural(objectName), CurrentProject.connection, adOpenForwardOnly, adLockOptimistic
        rs.AddNew
        rs(valueFieldName) = NewData
        If NotIsNullOrEmpty(ParentFieldName) Then
            rs(ParentFieldName) = ParentData
        End If
        rs.Update
        Id = rs(idFieldName)
        rs.Close
        Set rs = Nothing
        If OpenEdit Then
            DoCmd.OpenForm objectName & "_Edit", , , "Id=" & Id, acFormEdit, acDialog
        End If
    End If
    If Id = 0 Then
        NotInList = acDataErrContinue
    Else
        ctrl.Value = Id
        NotInList = acDataErrAdded
    End If

End Function

Public Sub Edit(ctrl As Control, objectName As String, Optional idFieldName As String = "Id")
    If NotIsNullOrEmpty(ctrl.Value) Then
        DoCmd.OpenForm objectName & "_Edit", , , idFieldName & "=" & ctrl.Value, acFormEdit, acDialog
    End If
End Sub
