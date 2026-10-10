Attribute VB_Name = "omAPIClipboard"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=e3a76075

' Clipboard text helpers: Unicode text, works in 32-bit and 64-bit Access (VBA7, Access 2010 or later).
'   ClipBoard_SetText "text"   -> True when the text is on the clipboard
'   ClipBoard_GetText()        -> the text on the clipboard ("" when there is none)

Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const OPEN_RETRIES As Long = 10
Private Const OPEN_RETRY_WAIT_MS As Long = 20

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)

' Another program can hold the clipboard for a moment, so try a few times before giving up.
Private Function OpenClipboardWithRetry() As Boolean
Dim i As Long

    For i = 1 To OPEN_RETRIES
        If OpenClipboard(0) <> 0 Then
            OpenClipboardWithRetry = True
            Exit Function
        End If
        omKernalAPIFunctions.Sleep OPEN_RETRY_WAIT_MS
    Next i
End Function

Public Function ClipBoard_SetText(strCopyString As String) As Boolean
Dim hMem As LongPtr
Dim pMem As LongPtr
Dim lngBytes As Long

    lngBytes = Len(strCopyString) * 2
    ' GMEM_ZEROINIT gives the two terminating zero bytes for free
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, lngBytes + 2)
    If hMem = 0 Then
        Exit Function
    End If
    pMem = GlobalLock(hMem)
    If pMem = 0 Then
        GlobalFree hMem
        Exit Function
    End If
    If lngBytes > 0 Then
        CopyMemory pMem, StrPtr(strCopyString), lngBytes
    End If
    GlobalUnlock hMem

    If OpenClipboardWithRetry() Then
        EmptyClipboard
        If SetClipboardData(CF_UNICODETEXT, hMem) <> 0 Then
            ' the clipboard owns the memory now, we must not free it
            ClipBoard_SetText = True
            hMem = 0
        End If
        CloseClipboard
    End If
    If hMem <> 0 Then
        GlobalFree hMem
    End If
End Function

Public Function ClipBoard_GetText() As String
Dim hMem As LongPtr
Dim pMem As LongPtr
Dim lngBytes As Long
Dim strText As String
Dim lngPos As Long

    If IsClipboardFormatAvailable(CF_UNICODETEXT) = 0 Then
        Exit Function
    End If
    If Not OpenClipboardWithRetry() Then
        Exit Function
    End If

    hMem = GetClipboardData(CF_UNICODETEXT)
    If hMem <> 0 Then
        pMem = GlobalLock(hMem)
        If pMem <> 0 Then
            lngBytes = CLng(GlobalSize(hMem))
            strText = Space$(Fix(lngBytes / 2))
            CopyMemory StrPtr(strText), pMem, Len(strText) * 2
            GlobalUnlock hMem
            ' the memory block can be larger than the text: cut at the terminating zero
            lngPos = InStr(1, strText, vbNullChar)
            If lngPos > 0 Then
                strText = Left$(strText, lngPos - 1)
            End If
        End If
    End If
    CloseClipboard
    ClipBoard_GetText = strText
End Function
