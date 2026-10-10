Attribute VB_Name = "omPrinterFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=00b32d3e

' The printers of this pc in a table, and a few lookups on it.
' The table omPrinters belongs to this module (hence the name): EnsurePrinterTable creates it when it is missing and
' adds the field IsDefault when an older table does not have it, LoadPrinters (re)fills it:
'   Id        = position in Application.Printers (0, 1, 2, ...)
'   Name      = the device name of the printer
'   IsDefault = the Windows default printer
' Extra yes/no fields of your own in this table (for example IsDymo) are kept; LoadPrinters can fill one of them.
Public Const PRINTER_TABLE As String = "omPrinters"

' Creates the table omPrinters when it does not exist; adds IsDefault to an existing table that has no such field.
Public Sub EnsurePrinterTable()
Dim db As DAO.Database
Dim td As DAO.TableDef
Dim fld As DAO.Field
Dim hasDefault As Boolean

    Set db = CurrentDb
    On Error Resume Next
    Set td = db.TableDefs(PRINTER_TABLE)
    On Error GoTo 0
    If td Is Nothing Then
        db.Execute "CREATE TABLE [" & PRINTER_TABLE & "] (Id LONG CONSTRAINT PK_" & PRINTER_TABLE & " PRIMARY KEY, Name TEXT(255), IsDefault YESNO)"
        Exit Sub
    End If
    For Each fld In td.Fields
        If StrComp(fld.Name, "IsDefault", vbTextCompare) = 0 Then hasDefault = True
    Next fld
    If Not hasDefault Then
        db.Execute "ALTER TABLE [" & PRINTER_TABLE & "] ADD COLUMN IsDefault YESNO"
    End If
End Sub

' Empties the table and fills it with the printers that Access knows (creates the table first when needed).
' flagField / flagText (optional): when the table has a yes/no field with the name flagField, it is set to True for every
' printer whose name contains flagText (not case sensitive). Example for a label printer: LoadPrinters "IsDymo", "Dymo"
Public Sub LoadPrinters(Optional flagField As String = "", Optional flagText As String = "")
Dim rs As New ADODB.Recordset
Dim prn As Printer
Dim i As Long
Dim defaultName As String
Dim hasFlag As Boolean
Dim fld As ADODB.Field

    EnsurePrinterTable
    CurrentProject.connection.Execute "DELETE FROM [" & PRINTER_TABLE & "]"
    defaultName = GetDefaultPrinterName
    rs.Open PRINTER_TABLE, CurrentProject.connection, adOpenForwardOnly, adLockOptimistic
    If Len(flagField) > 0 Then
        For Each fld In rs.Fields
            If StrComp(fld.Name, flagField, vbTextCompare) = 0 Then hasFlag = True
        Next fld
    End If
    For Each prn In Application.Printers
        rs.AddNew
        rs("Id") = i
        rs("Name") = prn.DeviceName
        rs("IsDefault") = (Len(defaultName) > 0 And StrComp(prn.DeviceName, defaultName, vbTextCompare) = 0)
        If hasFlag Then rs(flagField) = (InStr(1, prn.DeviceName, flagText, vbTextCompare) > 0)
        rs.Update
        i = i + 1
    Next prn
    rs.Close
    Set rs = Nothing
End Sub

' The Id (see above) of the first printer in the table whose name contains namePart (not case sensitive); Null when there is
' none. preferredId (optional): when that printer is one of the matches, it is returned instead of the first.
' The table is read as LoadPrinters left it, so call LoadPrinters first.
Public Function GetPrinterId(namePart As String, Optional preferredId As Variant = Null) As Variant
Dim rs As New ADODB.Recordset

    GetPrinterId = Null
    EnsurePrinterTable
    rs.Open "SELECT Id, Name FROM [" & PRINTER_TABLE & "] ORDER BY Id", CurrentProject.connection, adOpenStatic, adLockReadOnly
    Do While Not rs.EOF
        If InStr(1, rs("Name"), namePart, vbTextCompare) > 0 Then
            If IsNull(GetPrinterId) Then GetPrinterId = rs("Id")
            If IsNull(preferredId) Then Exit Do
            If rs("Id") = preferredId Then
                GetPrinterId = rs("Id")
                Exit Do
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Function

' The device name of the Windows default printer, or "" when there is none.
Public Function GetDefaultPrinterName() As String
    On Error Resume Next
    GetDefaultPrinterName = Application.Printer.DeviceName
End Function
