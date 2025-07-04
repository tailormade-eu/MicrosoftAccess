VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omExportManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Const emExcel = 1
Const emWord = 2
Const emPDF = 3
Private bMyDocuments As Boolean
Private strExportPath As String
Private strTemplatePath As String
Private strExportName As String
Private strTemplateFilename As String
Private strTemplate As String
Private bAppendDate As Boolean
Private bAppendName As Boolean
Private iExportType As Integer
Private lPhysicalDocumentTypeId As Long
Private lSupplierId As Long
Private lContractId As Long
Private lContractPeriodId As Long
Private strDBDocument As String
Private fso As New FileSystemObject
Private WShell As New WshShell
Private bAutoOpen As Boolean
Private strFileDestination As String
Public Property Get destinationFilename() As String
    destinationFilename = strFileDestination
End Property
Public Property Get MyDocuments() As Boolean
    MyDocuments = bMyDocuments
End Property
Public Property Let MyDocuments(Value As Boolean)
    bMyDocuments = Value
    strTemplatePath = omDefaultFunctions.GetDefault("Templates")
    If bMyDocuments Then
        strExportPath = fso.BuildPath(WShell.SpecialFolders("MyDocuments"), "DBExports")
        If Not fso.FolderExists(strExportPath) Then
            fso.createFolder strExportPath
        End If
    Else
        strExportPath = omDefaultFunctions.GetDefault("ExportPath")
    End If
End Property
Public Property Get AppendDate() As Boolean
    AppendDate = bAppendDate
End Property
Public Property Let AppendDate(Value As Boolean)
    bAppendDate = Value
End Property
Public Property Get AppendName() As Boolean
    AppendName = bAppendName
End Property
Public Property Let AppendName(Value As Boolean)
    bAppendName = Value
End Property
Public Property Get ExportName() As String
    ExportName = strExportName
End Property
Public Property Let ExportName(Value As String)
    strExportName = Value
End Property
Public Property Get AutoOpen() As Boolean
    AutoOpen = bAutoOpen
End Property
Public Property Let AutoOpen(Value As Boolean)
    bAutoOpen = Value
End Property

Public Property Get Template() As String
    Template = strTemplate
End Property
Public Property Let Template(Value As String)
    strTemplate = Value
End Property
Public Property Get TemplatePath() As String
    TemplatePath = strTemplatePath
End Property
Public Property Let TemplatePath(Value As String)
    strTemplatePath = Value
End Property

Public Property Get exportPath() As String
    exportPath = strExportPath
End Property
Public Property Let exportPath(Value As String)
    strExportPath = Value
End Property
Public Property Get ExportType() As Integer
    ExportType = iExportType
End Property
Public Property Let ExportType(Value As Integer)
    iExportType = Value
End Property
Public Property Get PhysicalDocumentTypeId() As Long
    PhysicalDocumentTypeId = lPhysicalDocumentTypeId
End Property
Public Property Let PhysicalDocumentTypeId(Value As Long)
    lPhysicalDocumentTypeId = Value
End Property
Public Property Get SupplierId() As Long
    SupplierId = lSupplierId
End Property
Public Property Let SupplierId(Value As Long)
    lSupplierId = Value
End Property
Public Property Get ContractId() As Long
    ContractId = lContractId
End Property
Public Property Let ContractId(Value As Long)
    lContractId = Value
End Property
Public Property Get ContractPeriodId() As Long
    ContractPeriodId = lContractPeriodId
End Property
Public Property Let ContractPeriodId(Value As Long)
    lContractPeriodId = Value
End Property

Public Property Get dbDocument() As String
    dbDocument = strDBDocument
End Property
Public Property Let dbDocument(Value As String)
    strDBDocument = Value
End Property

' Example implementation
'Public Sub Export()
'Dim strFileSource As String
'Dim strSupplierName As String
'Dim rs As New ADODB.Recordset
'
'    If bAppendName Then
'        rs.Open "SELECT DISTINCT Supplier_LastName, Supplier_FirstName FROM (Suppliers LEFT JOIN Contracts ON Suppliers.Supplier_ID = Contracts.Contract_Supplier_ID) LEFT JOIN ContractPeriods ON Contracts.Contract_ID = ContractPeriods.ContractPeriod_Contract_ID WHERE Supplier_ID=" & lSupplierId & " OR Contract_ID=" & lContractId & " OR ContractPeriod_ID=" & lContractPeriodId, CurrentProject.connection, adOpenForwardOnly, adLockReadOnly
'        If Not rs.EOF Then
'            strSupplierName = MakeFilenameSafe(rs("Supplier_LastName") & "_" & rs("Supplier_FirstName"))
'            strExportName = strExportName & "_" & strSupplierName
'        End If
'        rs.Close
'        Set rs = Nothing
'    End If
'    If bAppendDate Then
'        strExportName = strExportName & format(Now, "_YYYYMMDD_hhmm")
'    End If
'    Select Case iExportType
'        Case emExcel
'            If Not gFso.FolderExists("c:\DBExports") Then
'                gFso.createFolder "c:\dbexports"
'            End If
'            strFileSource = fso.BuildPath("c:\dbexports", strExportName & ".xls")
'            If lPhysicalDocumentTypeId <> 0 And lSupplierId <> 0 Then
'                strExportPath = gSupplierPhysicalDocument.GetPhysicalDocumentFolderByType(lSupplierId, lPhysicalDocumentTypeId)
'            End If
'            strFileDestination = fso.BuildPath(strExportPath, strExportName & ".xls")
'            If fso.FileExists(strFileSource) Then
'                fso.DeleteFile strFileSource
'            End If
'            If fso.FileExists(strFileDestination) Then
'                fso.DeleteFile strFileDestination
'            End If
'            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, strDBDocument, strFileSource
'        Case emWord
'            If Len(Template) > 0 Then
'                If lPhysicalDocumentTypeId <> 0 And lSupplierId <> 0 Then
'                    strExportPath = gSupplierPhysicalDocument.GetPhysicalDocumentFolderByType(lSupplierId, lPhysicalDocumentTypeId)
'                End If
'                strFileDestination = fso.BuildPath(strExportPath, strExportName & ".doc")
'                If gFso.FileExists(gFso.BuildPath(strTemplatePath, strTemplate)) Then
'                    gWordEngine.Translate gFso.BuildPath(strTemplatePath, strTemplate), strFileDestination, strDBDocument
'                End If
'            Else
'                strFileSource = fso.BuildPath("C:\", strExportName & ".doc")
'                strFileDestination = fso.BuildPath(strExportPath, strExportName & ".doc")
'                If fso.FileExists(strFileSource) Then
'                    fso.DeleteFile strFileSource
'                End If
'                If fso.FileExists(strFileDestination) Then
'                    fso.DeleteFile strFileDestination
'                End If
'                If ReportExists(strDBDocument) Then
'                    DoCmd.OutputTo acOutputReport, strDBDocument, acFormatRTF, strFileSource
'                ElseIf QueryExists(strDBDocument) Then
'                    DoCmd.OutputTo acOutputQuery, strDBDocument, acFormatRTF, strFileSource
'                End If
'            End If
'        Case emPDF
'            If Len(Template) > 0 Then
'                If lPhysicalDocumentTypeId <> 0 And lSupplierId <> 0 Then
'                    strExportPath = gSupplierPhysicalDocument.GetPhysicalDocumentFolderByType(lSupplierId, lPhysicalDocumentTypeId)
'                End If
'                strFileDestination = fso.BuildPath(strExportPath, strExportName & ".pdf")
'                strTemplateFilename = gFso.BuildPath(strTemplatePath, strTemplate)
'                If gFso.FileExists(strTemplateFilename) Then
'                    PDFForm.PDFFormFilling strTemplateFilename, strFileDestination, strDBDocument, "", "", "", "", "_", "."
'                End If
'            End If
'
'    End Select
'    If fso.FileExists(strFileSource) Then
'        fso.moveFile strFileSource, strFileDestination
'    End If
'    If bAutoOpen Then
'        WShell.Run Chr(34) & strFileDestination & Chr(34)
'    End If
'End Sub
