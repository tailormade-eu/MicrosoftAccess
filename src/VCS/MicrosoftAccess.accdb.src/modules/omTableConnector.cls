﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omTableConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Code written by Raoul Jacobs
' E. jara@opmaat.be
' Date Craeted : 200703
'
' this class will handle linking tables to the access program
' this class is called only at the startup of the application


Public DataFilename As String
Private m_DB As dao.Database
Private Const prefixLocalTemp = "T_"

Public Enum omTableConnectionType
    DatafileIsSource = 0
End Enum

Public Sub Connect(ConnectionType As omTableConnectionType)
    'If ConnectionType = DatafileIsSource Then
        OpenDB
        Link ConnectionType
        m_DB.Close

    'End If
    SysCmd (acSysCmdClearStatus)
End Sub



'Private clsstrName As String
'Private clsstrTableType As String
'Private clsConnectionType As Integer
'Private clsbReadOnly As Boolean
'
'Public Property Let ReadOnly(bFlag As Boolean)
'
'    clsbReadOnly = bFlag
'
'End Property
'Public Property Get ReadOnly() As Boolean
'
'    ReadOnly = clsbReadOnly
'
'End Property
'
'Public Property Let Name(strName As String)
'
'    clsstrName = strName
'
'End Property
'
'
'Public Property Get Name() As String
'
'    Name = clsstrName
'
'End Property
'Public Property Get Path() As String
'
'    Path = ExtractPath(clsstrName)
'
'End Property
'
'Public Property Let TableType(strTableType As String)
'
'    clsstrTableType = strTableType
'
'End Property
'
'Public Property Get TableType() As String
'
'    TableType = clsstrTableType
'
'End Property
'Public Property Let ConnectionType(dctType As Integer)
'
'    clsConnectionType = dctType
'
'End Property
'Public Property Get ConnectionType() As Integer
'
'    ConnectionType = clsConnectionType
'
'End Property
'
'Public Sub Connect()
'Dim varReturn As Variant
'
'    On Error GoTo Connect_Error
'    ' Open database
'    OpenDB
'    Link
'
'    varReturn = SysCmd(acSysCmdClearStatus)
'    Exit Sub
'
'Connect_Error:
'
'    Select Case Err
'        Case Else
'            ErrorBox.Module = "Class : DataConnection"
'            ErrorBox.Procedure = "Procedure : Connect"
'            ErrorBox.Number = "Error : " & Err
'            ErrorBox.Description = Error
'            ErrorBox.Execute
'            If gbErrorResume = True Then
'                If gbDebugMode Then
'                    Stop
'                    Resume
'                Else
'                    Resume Next
'                End If
'            Else
'                DoCmd.Quit acQuitSaveAll
'            End If
'    End Select
'
'End Sub
'Public Sub Disconnect()
'
'    clsDB.Close
'
'End Sub
'
Private Function OpenDB() As Boolean
Dim lAttribs As Long
Dim lResult As Long

    On Error GoTo OpenDB_Error
    OpenDB = True

    'Select Case clsConnectionType
        'Case 1  ' DataConnectionType = MSAccess
            Set m_DB = DBEngine.OpenDatabase(Me.DataFilename)
        'Case 2  ' DataConnectionType = ODBC
    'End Select

    'lAttribs = GetFileAttributes(clsstrName)
    'If clsbReadOnly = True Then
    '    If Not ((lAttribs And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY) Then
    '        lAttribs = lAttribs Xor FILE_ATTRIBUTE_READONLY
    '    End If
    'Else
    '    If ((lAttribs And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY) Then
    '        lAttribs = lAttribs Xor FILE_ATTRIBUTE_READONLY
    '    End If
    'End If
    'lResult = SetFileAttributes(clsstrName, lAttribs)

    Exit Function

OpenDB_Error:

    Select Case Err
        'Case 3024   ' Database not Found
        '    If CreateDB <> True Then
        '        Exit Function
        '    Else
        '        Resume Next
        '    End If
        Case Else
            OpenDB = False
            MsgBox Error & " (" & Err & ")", vbCritical
            DoCmd.Quit acQuitSaveNone
    End Select
End Function


'Private Function CreateDB() As Boolean
'Dim varReturn As Variant
'
'    On Error GoTo CreateDB_Error
'    CreateDB = True
'    varReturn = SysCmd(acSysCmdSetStatus, "msgCreateDB - " & clsstrName)
'    Set clsDB = DBEngine.CreateDatabase(clsstrName, dbLangGeneral)
'
'    Exit Function
'
'CreateDB_Error:
'
'    Select Case Err
'        Case Else
'            CreateDB = False
'            ErrorBox.Module = "Class : DataConnection"
'            ErrorBox.Procedure = "Procedure : CreateDB"
'            ErrorBox.Number = "Error : " & Err
'            ErrorBox.Description = Error
'            ErrorBox.Execute
'            If gbErrorResume = True Then
'                If gbDebugMode Then
'                    Stop
'                    Resume
'                Else
'                    Resume Next
'                End If
'            Else
'                DoCmd.Quit acQuitSaveAll
'            End If
'    End Select
'
'End Function
'

Private Function Link(ConnectionType As omTableConnectionType, Optional ReconnectAll As Boolean = True) As Boolean
Dim rsToConnect As dao.Recordset
Dim rsConnected As dao.Recordset
Dim rsRemoteTables As dao.Recordset
Dim varReturn As Variant

    On Error GoTo Link_Error
    Link = True
    DoCmd.SetWarnings False

    If ConnectionType = DatafileIsSource Then
        Set rsToConnect = m_DB.OpenRecordset("SELECT Name AS LinkName, Type, Name FROM MSysObjects WHERE Type = 1 AND Flags=0 ORDER BY Name")
    Else
        'Set rsToConnect = CurrentDb.OpenRecordset("SELECT Right([Name],Len([name])-Len('" & clsstrTableType & "')) AS LinkName, Type, Name FROM MSysObjects WHERE (((Type) = 1) And ((Name) Like '" & clsstrTableType & "*')) ORDER BY Right([Name],Len([name])-Len('" & clsstrTableType & "'))")
    End If
    Set rsRemoteTables = m_DB.OpenRecordset("SELECT Name, Type FROM MSysObjects WHERE Type = 1 ORDER BY Name")
    Set rsConnected = CurrentDb.OpenRecordset("SELECT Name,Database FROM MSysObjects WHERE (((Type) = 6) ) ORDER BY Name")
    varReturn = SysCmd(acSysCmdInitMeter, "msgLink", rsToConnect.RecordCount)
    While Not rsToConnect.EOF
        ' Check if Table Exists in Remote Database
        rsRemoteTables.FindFirst "Name = '" & rsToConnect("LinkName") & "'"
        If rsRemoteTables.NoMatch Then
            ' Export Table to Remote Database
            DoCmd.TransferDatabase acExport, "Microsoft Access", m_DB.Name, acTable, rsToConnect("Name"), rsToConnect("LinkName")
        End If
        ' check if Table is Linked
        rsConnected.FindFirst "Name = '" & rsToConnect("LinkName") & "'"
        If rsConnected.NoMatch Then
            ' Link Table
            DoCmd.TransferDatabase acLink, "Microsoft Access", m_DB.Name, acTable, rsToConnect("LinkName"), rsToConnect("LinkName")
        ElseIf (Not rsConnected.NoMatch And rsConnected("Database") <> m_DB.Name) Then
            ' Delete Old & Link New Table
            DoCmd.DeleteObject acTable, rsToConnect("LinkName")
            DoCmd.TransferDatabase acLink, "Microsoft Access", m_DB.Name, acTable, rsToConnect("LinkName"), rsToConnect("LinkName")
        ElseIf ReconnectAll Then
            DoCmd.DeleteObject acTable, rsToConnect("LinkName")
            DoCmd.TransferDatabase acLink, "Microsoft Access", m_DB.Name, acTable, rsToConnect("LinkName"), rsToConnect("LinkName")
        End If

        rsToConnect.MoveNext
        If rsToConnect.AbsolutePosition >= 0 Then
            varReturn = SysCmd(acSysCmdUpdateMeter, rsToConnect.AbsolutePosition + 1)
        End If
        DoEvents
    Wend
    rsToConnect.Close
    rsRemoteTables.Close
    rsConnected.Close
    DoCmd.SetWarnings True
    Exit Function

Link_Error:

    Select Case Err
        Case Else
            Link = False
            MsgBox Error & " (" & Err & ")", vbCritical
            DoCmd.Quit acQuitSaveAll
            Resume
    End Select
End Function

Private Sub Class_Terminate()
    Set m_DB = Nothing
End Sub
