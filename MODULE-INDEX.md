# Module Index

Index van alle modules in de library — automatisch gegenereerd op basis van Public Sub/Function/Property-signaturen. Bij twijfel: lees het bronbestand zelf.

## omADODBFunctions.bas
- `ADODBCreateCopyRecord(rsSrc As ADODB.Recordset, rsDst As ADODB.Recordset, Optional PrimaryKeyName As String, Optional LinkName As String, Optional LinkId As Long, Optional ignoreFields As String) As Long` — Create copy of record from source to destination recordset
- `ADODBCopyRecord(rsSrc As ADODB.Recordset, rsDst As ADODB.Recordset, Optional PrimaryKeyName As String, Optional LinkName As String, Optional LinkId As Long, Optional ignoreFields As String)` — Copy field values between recordsets, excluding specified fields

## omAPI32FileFunctions.bas
- `GetShortFileName(sFileName As String) As String` — Get 8.3 DOS-compatible filename
- `GetShortFolderName(sFolderName As String) As String` — Get 8.3 DOS-compatible folder path
- `ShellWait(Pathname As String, Optional WindowStyle As Long)` — Execute process and wait for completion
- `GetUserName() As Variant` — Retrieve logged-in Windows username

## omAPIClipboard.bas
- `ClipBoard_SetText(strCopyString As String) As Boolean` — Copy text string to system clipboard
- `ClipBoard_GetText() As String` — Retrieve text from system clipboard
- `CopyOlePiccy(Piccy As Object)` — Copy OLE object to clipboard

## omAPISpecialFolders.bas
- `SpecialFolder(ByVal SFConst As SpecialFolderConst) As String` — Get path to Windows special folder by constant

## omAppSettings.cls
- `Load(SettingsName As String) As Variant` — Load setting value from appSettings table
- `Save(SettingsName As String, NewValue As String)` — Save or update setting in appSettings table

## omArrayFunctions.bas
- `ObjectArrayAdd(objectArray() As Object, Value As Object, Optional Clear As Boolean)` — Add object to dynamic array, optionally clearing first
- `StringArrayAdd(stringArray() As String, Value As String, Optional Clear As Boolean)` — Add string to dynamic array, optionally clearing first
- `StringArrayClear(stringArray() As String)` — Clear all elements from string array
- `TestAddString()` — Test routine for array operations
- `StringArrayAddArray(stringArray() As String, valueArray() As String)` — Append array contents to another array
- `StringArrayCount(stringArray() As String) As Long` — Get count of elements in string array
- `StringArrayFind(stringArray() As String, Value As String, includeContains As Boolean) As Long` — Find element index by exact match or contains
- `StringArrayContains(stringArray() As String, Value As String) As Boolean` — Check if array contains value
- `StringArrayDoesNotContain(stringArray() As String, Value As String) As Boolean` — Check if array does not contain value
- `StringArrayLength(stringArray() As String) As Long` — Get total length of string array
- `StringArrayToString(stringArray() As String, Optional delimiter As String, Optional skipNullBlankSpaces As Boolean) As String` — Join array elements with delimiter
- `StringArrayRemoveEmpty(stringArray() As String) As String()` — Return new array with empty elements filtered out
- `IsArrayInitialized(arr As Variant) As Boolean` — True when the array has been dimensioned (also with 0 elements)

## omBankAccountFunctions.bas
- `ValidateIbanAndFormatBankAccount(ByVal BankNr As Variant) As String` — Validate IBAN checksum and format with spaces
- `FormatIbanWithSpaces(IBAN As String) As String` — Format IBAN with 4-character spacing
- `ValidateIban(ByVal psIban As String) As Boolean` — Validate IBAN modulo-97 checksum
- `IsOnlyNumbers(ByVal psText As String) As Boolean` — Check if string contains only digits
- `IsAlphaNum(ByVal psText As String) As Boolean` — Check if string contains only alphanumeric characters
- `ValidateBelgianAccountNumber(IBAN As String) As Boolean` — Validate Belgian IBAN account number checksum

## omBaseClass.cls
- `Load(CountryId As Long)` — Load country record by ID
- `Clear()` — Clear all properties to null/empty

## omByteArrayFunctions.bas
- `ByteArrayToHexString(ba() As Byte) As String` — Convert byte array to hexadecimal string with 0x prefix
- `ByteArrayToString(ba() As Byte) As LongLong` — Convert byte array to numeric value
- `ByteArrayToHexString_Test()` — Test routine for hex conversion
- `ByteArrayToString_Test()` — Test routine for numeric conversion

## omCollectionFunctions.bas
- `SortCollectionTest()` — Test routine for collection sorting
- `SortCollection(data As Collection, Optional asNumeric As Boolean) As Boolean` — Sort collection of key-value objects
- `CompareKeys(x As Variant, y As Variant, Optional asNumeric As Boolean) As Boolean` — Compare two values as numeric or string

## omCommandBarFunctions.bas
- `RegisterPrintContextMenu()` — Create and register custom print context menu
- `ListCommandBarObjects(Optional contains As String)` — List all CommandBar objects to Immediate window, optionally filtered
- `fnPrint()` — Execute print command

## omConnectionString.cls
- `ParseByTableName(Name As String)` — Parse connection string from linked table definition
- `ParseByQueryName(Name As String)` — Parse connection string from query definition
- `GetValue(Name As String) As Variant` — Get connection property value by name
- `ToString() As String` — Convert connection properties to formatted string
- `Driver` (get/let) — database driver name with version detection
- `GetConnectionString(databaseName As String, serverName As String, ODBCConnection As Boolean, Optional ConnectionType, Optional DSN, Optional Login, Optional Password, Optional GroupName, Optional portNr, Optional encryptType) As String` — Build connection string for database with specified parameters
- `GetCurrentConnectionStringForType(ODBCConnection As Boolean, Optional ConnectionType, Optional encryptType) As String` — Build connection string for current connection type
- `GetEncryptString(encrypType As EncryptTypes) As String` — Get encryption clause for connection string

## omControl.cls
- `Load(ctrl As Object, Optional autoRename As Boolean)` — Load control metadata and optionally rename control using conventions

## omControlFunctions.bas
- `IfNullOrEmpty(ctrl As Control, Optional UseDefaultValue As Boolean)` — Restore null/empty control to default or previous value
- `SetDefaultValue(ctrl As Control)` — Set control value from defaultValue property
- `IsFormOpen(formName As String) As Boolean` — Check if form is currently open
- `GetOpenForm(formName As String) As Form` — Get reference to open form, returns Nothing if not open
- `NotInList(ctrl As Control, objectName As String, idFieldName As String, valueFieldName As String, NewData As String, Optional OpenEdit As Boolean, Optional ParentFieldName As String, Optional ParentData) As Integer` — Handle NotInList event to add new record
- `Edit(ctrl As Control, objectName As String, Optional idFieldName As String)` — Open edit form for control's linked record

## omDAOFunctions.bas
- `SetQueryDefProperty(queryName As String, propertyName As String, Value As String, Optional propertyType As dao.DataTypeEnum)` — Create or update custom property on QueryDef

## omDatabaseFunctions.bas
- `LinkTable(databaseFilename As String, sourceName As String, destinationName As String)` — Link table from another Access database
- `UnlinkTable(Name As String)` — Remove linked table
- `CreateDataBase(filename As String) As Boolean` — Create new Access database file
- `ExportTable(filename As String, sourceName As String, destinationName As String) As Boolean` — Export table to another Access database
- `CompactDataBase(filename As String)` — Compact and repair database
- `CopyTable(databaseFilename As String, TableSource As String, TableDestination As String) As Variant` — Copy table between Access databases

## omDateFunctions.bas
- `IsDateInPast(dt As Variant) As Boolean` — Empty/Null = False; otherwise Date > dt (day-granularity, no time-of-day)
- `GetMonthYearEndDate(ByVal Month As Long, ByVal year As Long) As Date` — Get last day of month/year
- `GetPeriodEndDate(ByVal Period As Long) As Date` — Get end date from YYYYMM period
- `GetMonthYearStartDate(ByVal Month As Long, ByVal year As Long) As Date` — Get first day of month/year
- `GetPeriodStartDate(ByVal Period As Long) As Date` — Get start date from YYYYMM period
- `GetVarPeriodStartDate(ByVal Period As Variant) As Date` — Get start date from variant period
- `GetVarPeriodEndDate(ByVal Period As Variant) As Date` — Get end date from variant period
- `ConvertDateToYYYYMMDD(ByVal dt As Date) As Long` — Convert date to YYYYMMDD format
- `ConvertVarDateToYYYYMMDD(ByVal dt As Variant) As Long` — Convert variant date to YYYYMMDD
- `ConvertDateToPeriod(ByVal dt As Date) As Long` — Convert date to YYYYMM period
- `ConvertVarDateToPeriod(ByVal dt As Variant, Optional NullValue As Long = 0) As Long` — Convert variant date to period
- `GetDateMonthStartDate(ByVal dt As Date)` — Get first day of month for given date
- `GetDateMonthEndDate(ByVal dt As Date)` — Get last day of month for given date
- `ConvertVarDateToDate(ByVal dt As Variant) As Date` — Convert variant to date
- `NextBirthDate(ByVal dt As Variant) As Variant` — Calculate next occurring birthdate
- `GetTimeStamp(Optional ts As Date = 0) As String` — Get formatted timestamp
- `YYYYMM_Offset(lMonthTemp As Long, lMonthOffset As Long) As Long` — Add/subtract months from YYYYMM
- `YYYYWW(ByVal dateTemp As Variant) As Long` — Convert date to YYYYWW (year-week)
- `YYYYWW_Date(ByVal lYYYYWW As Long) As Date` — Convert YYYYWW to date
- `WorkingDays(ByVal lYearMonth As Long) As Long` — Count working days in month
- `WorkingDaysFrom(ByVal lYearMonth As Long, ByVal dateFrom As Variant) As Long` — Count working days from date onwards
- `WorkingDaysTill(ByVal lYearMonth As Long, ByVal dateFrom As Variant) As Long` — Count working days until date
- `WorkingDaysBetween(ByVal dateStart As Variant, ByVal dateEnd As Variant) As Long` — Count working days between dates
- `MonthsBetween(ByVal dateStart As Variant, ByVal dateEnd As Variant) As Double` — Calculate months between dates
- `Period(ByVal dt As Variant) As Long` — Convert variant date to period (wrapper)
- `PeriodAddMonths(ByVal dt As Variant, months As Double) As Long` — Add months to date, return period
- `MonthAddMonths(ByVal dt As Variant, months As Double) As Long` — Get month after adding months
- `YearAddMonths(ByVal dt As Variant, months As Double) As Long` — Get year after adding months
- `DateYYYYMMDD(ByVal dt As Variant) As Long` — Convert variant to YYYYMMDD
- `StartDate(ByVal Period As Long) As Date` — Get start date of period
- `UniDate(dt As Date) As Long` — Convert date to YYYYMMDD (wrapper)
- `DateOnly(dt As Variant) As Date` — Extract date without time
- `YYYYMM(dt As Variant) As Long` — Convert date to YYYYMM (wrapper)
- `IsPeriod(Period As Variant, Optional yearDelta = 100) As Boolean` — Validate period is within year range
- `AddToDate(dt As Date, Optional addYear As Long = 0, Optional addMonth As Long = 0, Optional addDay As Long = 0, Optional EndOfMonth As Boolean = False) As Date` — Add years/months/days to date
- `IsTimeOnly(dt As Variant) As Boolean` — Check if variant contains time without date component
- `ConcatenateCurrentDateIfTimeOnly(dt As Variant, Optional defaultDate As Date = 0) As Variant` — Combine current date with time-only value
- `NumberToDHMS(t As Double) As String` — Format number as days/hours/minutes/seconds
- `GetTimeBetween(StartDate As Date, EndDate As Date, excludeWeekends As Boolean)` — Calculate time span
- `GetTimeBetweenTimeRangesFlat(StartDate As Date, EndDate As Date, excludeWeekends As Boolean, range1StartDate As Date, range1EndDate As Date, Optional range2..4)` — Calculate time in multiple fixed ranges
- `AddTimeRangeToArray(arr() As Object, StartDate As Date, EndDate As Date, Optional AllowZeros As Boolean = False, Optional Clear As Boolean = False)` — Add time range to array
- `AddTimeRangeToArray_Test()` — Test time range array addition
- `GetTimeBetweenTimeRanges(StartDate As Date, EndDate As Date, excludeWeekends As Boolean, timeRanges() As Object)` — Calculate time in variable ranges
- `GetTimeBetweenTimeRanges_Test()` — Test time range calculation
- `GetTimeBetweenTimeRangesFlat_Test()` — Test flat time range calculation
- `GetTimeInSeconds(t As Date, Optional ZeroIsDay As Boolean = False) As Long` — Convert time to seconds
- `IsWeekDay(dt As Date, Optional FirstDayOfWeek As VbDayOfWeek = vbMonday) As Boolean` — Check if date is weekday
- `GetQuarter(dt As Date, Optional FirstDayOfWeek, Optional FirstWeekOfYear) As Long` — Get quarter number
- `GetWeekNumber(dt As Date, Optional FirstDayOfWeek, Optional FirstWeekOfYear) As Long` — Get week number
- `GetWeekDay(dt As Date, Optional FirstDayOfWeek, Optional FirstWeekOfYear) As Long` — Get day of week
- `PopulateCalendar()` — Populate calendar table
- `DaysBetween(StartDate As Date, EndDate As Date) As Long` — Count days between dates
- `EOMonth(dt As Date) As Date` — Get last day of month
- `CalculatePayDate(dt As Date, Days As Integer, EndOfMonth As Boolean, ExtraDays As Integer) As Date` — Calculate payment date with offset
- `YearsBetween(StartDate As Date, EndDate As Date) As Long` — Number of whole years between two dates

## omDefaultFunctions.bas
- `GetDefault(Name As String) As Variant` — Retrieve default value by name
- `SaveDefault(Name As String, Value As Variant)` — Save default value
- `SaveSystemDefault(Name As String, Value As Variant)` — Save system-wide default
- `GetSystemDefault(Name As String) As Variant` — Retrieve system default by name
- `Initialize(Optional Reset As Boolean = False)` — Initialize default storage system
- `CreateDefaultTables()` — Create default storage tables

## omDefaults.cls
- `Load(defaultName As String) As Variant` — Load default value from storage
- `Save(defaultName As String, Value As Variant)` — Save default value to storage

## omDevelopmentFunctions.bas
- `ExportQueries()` — Export all queries to file
- `TemplateCopy(strFind As String, strReplace As String)` — Copy template with find/replace
- `QueryCopy(strQuery As String, strFind As String, strReplace As String)` — Copy query with text replacement
- `FormCopy(strForm As String, strFind As String, strReplace As String)` — Copy form with text replacement
- `FormControlRename(strForm As String, lControlType As Long, strFind As String, strReplace As String, Optional OnlyCaption As Boolean = False, Optional TestRun As Boolean = False)` — Rename form controls by type
- `DeleteTempObjects()` — Delete temporary database objects
- `ReportReplaceBackColor(ReportName As String, OldColor As Long, NewColor As Long)` — Replace report background colors
- `ReportReplaceForeColor(ReportName As String, OldColor As Long, NewColor As Long)` — Replace report text colors
- `ReportReplaceColors()` — Replace all report colors
- `AllFormsReplace(strFind As String, strReplace As String)` — Replace text in all forms
- `AllFormsDisableDialog()` — Disable dialog mode on all forms
- `AllFormsDisablePopUp()` — Disable popup on all forms

## omDynamicRecordState.cls
- `SetupFieldsFromRecordset(rsSrc As ADODB.Recordset, fieldNames As String)` — Initialize tracking from recordset fields
- `SetupFields()` — Initialize field tracking
- `SetOriginalRecord(rsSrc As ADODB.Recordset)` — Store original record state
- `SetDatabaseRecord(rsSrc As ADODB.Recordset)` — Store database record state
- `Compare()` — Compare record states

## omDynamicRecordStateTest.bas
- `Test()` — Test dynamic record state tracking
- `TestFromRecordset()` — Test setup from recordset
- `TestAutoClass()` — Test automatic class setup

## omExcel.cls
- `LoadWB(filename As String)` — Open Excel workbook
- `CloseWB(Optional saveChanges As Boolean = True, Optional filename As String = "")` — Close workbook with save option
- `SetVisible(Optional visible As Boolean = True)` — Show/hide Excel application
- `SelectColumn(columnIndex As Long)` — Select column by index
- `SelectRow(rowIndex As Long)` — Select row by index
- `SetSelectionBold(Optional boldFlag As Boolean = True)` — Apply bold to selection
- `FormatColumn(columnIndex As Long, numberFormat As String)` — Apply number format to column
- `FormatCells(...)` — Format cell range (row/column start-end, number format, alignment, merge, wrap)
- `WriteCell(rowIndex As Long, columnIndex As Long, Value As Variant, Optional numberFormat, Optional horizontalAlignment, Optional verticalAlignment)` — Write value to cell
- `InsertRows(rowIndex As Long, Optional numberOfRows As Long = 1)` — Insert rows in sheet
- `SetHeaderBold()` — Apply bold to header row
- `AutoFit()` — Auto-fit columns and rows
- `SelectionUpdateBold(bold As Boolean)` — Update bold on selected cells

## omExcelEngine.cls
- `Merge(Template As String, SaveAs As String, rs As ADODB.Recordset, Optional show As Boolean = False)` — Merge recordset into Excel template
- `MergeFromCurrentRecord(xlWs As Excel.Worksheet, rs As ADODB.Recordset, Optional delimiters As String = "<<,>>")` — Merge current recordset row into worksheet

## omExcelFunctions.bas
- `ReplaceFirst(sh As Excel.Worksheet, findValue As String, replaceValue As String)` — Replace first occurrence in sheet
- `FindFirst(sh As Excel.Worksheet, findValue As String) As Excel.range` — Find first matching cell
- `DeleteRow(sh As Excel.Worksheet, rowIndex As Long)` — Delete row
- `DeleteSheet(sh As Excel.Worksheet)` — Delete worksheet
- `HideRow(sh As Excel.Worksheet, rowIndex As Long)` — Hide row
- `ShowRow(sh As Excel.Worksheet, rowIndex As Long)` — Show row
- `SelectRow(sh As Excel.Worksheet, rowIndex As Long)` — Select row
- `FormatRow(sh As Excel.Worksheet, rowIndex As Long, numberFormat As String)` — Apply number format to row
- `HideColumn(sh As Excel.Worksheet, columnIndex As Long)` — Hide column
- `ShowColumn(sh As Excel.Worksheet, columnIndex As Long)` — Show column
- `SelectColumn(sh As Excel.Worksheet, columnIndex As Long)` — Select column
- `SetSelectionBold(sh As Excel.Worksheet, Optional boldFlag As Boolean = True)` — Apply bold to selection
- `FormatColumn(sh As Excel.Worksheet, columnIndex As Long, numberFormat As String)` — Apply number format to column
- `FormatCells(sh As Excel.Worksheet, ...)` — Format cell range
- `WriteCell(sh As Excel.Worksheet, rowIndex As Long, columnIndex As Long, Value As Variant)` — Write value to cell
- `SelectCell(sh As Excel.Worksheet, rowIndex As Long, columnIndex As Long)` — Select cell
- `SelectRange(sh As Excel.Worksheet, rowIndexStart As Long, columnIndexStart As Long, rowIndexEnd As Long, columnIndexEnd As Long)` — Select range
- `GoToNextFullCell(sh As Excel.Worksheet, direction As XlDirection)` — Navigate to next filled cell
- `GetCurrentRow(sh As Excel.Worksheet) As Long` — Get active cell row
- `GetCurrentColumn(sh As Excel.Worksheet) As Long` — Get active cell column
- `NextRowHasValue(sh As Excel.Worksheet) As Boolean` — Check if next row has value
- `NextColumnHasValue(sh As Excel.Worksheet) As Boolean` — Check if next column has value
- `InsertColumn(sh As Excel.Worksheet, columnIndex As Long, Optional direction As XlDirection = xlToRight)` — Insert column
- `InsertRow(sh As Excel.Worksheet, rowIndex As Long, Optional copyRowIndex, Optional moveRowIndex, Optional direction, Optional activate)` — Insert row
- `FindSheetByName(wb As Excel.Workbook, sheetName As String) As Excel.Worksheet` — Find sheet by name
- `FindSheetByCell(wb As Excel.Workbook, Value As String, Optional rowIndex As Long = 1, Optional columnIndex As Long = 1) As Excel.Worksheet` — Find sheet containing value
- `FindRowByColumn(xs As Excel.Worksheet, columnIndex As Long, Value As String, Optional startRowIndex, Optional endRowIndex) As Long` — Find row with value in column
- `FindRow(xs As Excel.Worksheet, Value As String) As Long` — Find row containing value
- `FindColumn(xs As Excel.Worksheet, Value As String) As Long` — Find column containing value
- `RenameSheet(sh As Excel.Worksheet, newName As String) As String` — Rename worksheet
- `SetVisible(xa As Excel.Application, visible As Boolean)` — Show/hide Excel application
- `AutoFitColumns(sh As Excel.Worksheet)` — Auto-fit column widths
- `AutoFitRows(sh As Excel.Worksheet)` — Auto-fit row heights
- `GetMaxRow(sh As Excel.Worksheet)` — Get last row with data
- `GetMaxColumn(sh As Excel.Worksheet)` — Get last column with data
- `ConvertToImage(sh As Excel.Worksheet, columnIndex As Long, findString As String, replaceString As String)` — Convert cell range to image
- `SetRowHeight(sh As Excel.Worksheet, startRow As Long, endRow As Long, height As Double)` — Set row height
- `SetRowWidth(sh As Excel.Worksheet, startColumn As Long, endColumn As Long, width As Double)` — Set column width
- `UpdateFontOnSelection(sh As Excel.Worksheet, bold As Boolean)` — Update font bold on selection

## omExportFunctions.bas
- `ExportQuery(queryName As String, Optional exportPath As String = "", Optional filename As String, Optional withTimestamp As Boolean = False) As String` — Export query results to Excel

## omExportManager.cls
- `destinationFilename` (get/let) — destination file path
- `MyDocuments` (get/let) — use My Documents folder
- `AppendDate` (get/let) — append date to filename
- `AppendName` (get/let) — append name to filename
- `ExportName` (get/let) — export name
- `AutoOpen` (get/let) — auto-open after export
- `Template` (get/let) — template file
- `TemplatePath` (get/let) — template folder
- `exportPath` (get/let) — export folder
- `ExportType` (get/let) — export type ID
- `PhysicalDocumentTypeId` (get/let) — document type ID
- `SupplierId` (get/let) — supplier ID
- `ContractId` (get/let) — contract ID
- `ContractPeriodId` (get/let) — contract period ID
- `dbDocument` (get/let) — database document reference

## omFileFunctions.bas
- `BuildPathFileExists(path As String, filename As String) As String` — Build and verify file path exists
- `FolderExists(path As String) As Boolean` — Check if folder exists
- `FileExists(path As String, filename As String) As String` — Check if file exists
- `OpenPathAndFile(AppPath As String, filename As String) As String` — Open file if found
- `OpenFile(filename As String)` — Open file in Explorer
- `OpenUrl(url As String)` — Open URL in browser
- `OpenFolder(path As String, Optional createPath As Boolean = False)` — Open folder in Explorer
- `CreateDatebase(filename As String)` — Create new database file
- `DeleteFile(filename As String, Optional force As Boolean = False)` — Delete file
- `DeleteFolder(folderName As String, Optional force As Boolean = False)` — Delete folder
- `RenameFile(strOldName As String, strNewName As String, Optional strNewExtension As String = "")` — Rename file with optional extension change
- `ExtractPath(strFileName As String) As String` — Extract folder path from full filename
- `RemoveExtension(Name As String) As String` — Remove extension from filename
- `GetExtension(filename As String) As String` — Get file extension
- `MakeFilenameSafe(filename As String) As String` — Remove invalid characters from filename
- `CreateFolderPath(strPath As String, Optional OpenFolderInExplorer As Boolean = False)` — Create folder hierarchy
- `ReadFilesIntoTable(folderPath As String, tableName As String, Optional pattern As String = "")` — Import file list into table
- `FileLocked(strFileName As String, Optional displayError As Boolean = False) As Boolean` — Check if file is open elsewhere
- `GetDesktopFolder() As String` — Get Desktop folder path
- `GetUserRootFolder() As String` — Get user home folder path
- `ReadFileToString(path As String) As String` — Read entire file into string
- `PrintFile(filename As String)` — Print file using default application
- `WriteStringToFile(filename As String, s As String, Optional unicode As Boolean = False)` — Write string to file
- `CopyToLocalTemp(filename As String) As String` — Copy a file to the temp folder and return the path of the copy ("" when a locked copy is in the way)

## omFormFunctions.bas
- `UpdateEnableByTag(Form As Form, tag As String, state As Boolean, Optional delimiter = ",")` — Enable/disable controls by tag
- `UpdateVisibleByTag(Form As Form, tag As String, state As Boolean, Optional delimiter = ",")` — Show/hide controls by tag
- `OpenForm(Name As String, Optional view As AcFormView = acFormDS) As Variant` — Open form by name
- `CloseForms(Optional keepOpen As String = "flow")` — Close all forms except named
- `OpenEditScreen(parentForm As Form, formName As String, Optional keyName As String = "Id", Optional windowMode As AcWindowMode = acDialog, Optional Requery As Boolean = True)` — Open edit dialog from parent
- `UpdateModifyTracking(frm As Form)` — Update modification timestamp
- `UpdateCreateTracking(frm As Form, Optional UpdateModify = True)` — Update creation timestamp
- `ClearMultiValueField(frm As Form, fieldName As String, Optional whereClause As String = "", Optional Requery As Boolean = True)` — Clear multi-value field contents (DAO ComplexType/Recordset2, i.p.v. UPDATE/.Value=Null die error 3032 geeft)
- `ListFormFields(formName As String)` — Debug: list all form fields
- `UpdateModifyTrackingDao(rs As DAO.Recordset, Optional isDelete As Boolean)` — Fill create/modify (or delete) date and user fields of a recordset that is being edited
- `UpdateModifyTrackingADodb(rs As ADODB.Recordset, Optional isDelete As Boolean)` — Same for an ADODB recordset
- `UpdateLockedWithoutTag(Form As Form, state As Boolean)` — Lock or unlock every control that has no tag

## omFTP.bas
- `UploadFile(ftpAddress As String, Login As String, Password As String, filename As String, Optional ftpPath As String = "/") As String` — Upload file to FTP server

## omGlobalDeclarations.bas
*(geen Public Sub/Function/Property — bevat enkel globale declaraties/constanten)*

## omGuidFunctions.bas
- `CreateGuid() As String` — Generate GUID
- `CreateRandomEAN13() As String` — Generate random EAN13 barcode

## omHashAlgorithms.cls
- `MD5(sIn As String, Optional bB64 As Boolean) As String` — MD5 hash of text (hex or base64), also SHA1, SHA256, SHA384, SHA512 with the same signature
- `ConvToBase64String(vIn As Variant) As Variant` — Bytes to base64 string
- `ConvToHexString(vIn As Variant) As Variant` — Bytes to hex string

## omHashAlgorithmsFunctions.bas
- `MD5(sIn As String, Optional bB64 As Boolean) As String` — Module-level wrapper around omHashAlgorithms, also SHA1, SHA256, SHA384, SHA512
- `ConvToBase64String(vIn As Variant) As Variant` — Bytes to base64 string

## omHTML.bas
- `Run()` — Main entry point for HTML generation
- `GenerateDomain(DomainName As String, DomainFolder As String, DefaultCount As Long)` — Generate HTML domain structure
- `GenerateTemplate()` — Generate HTML template from database
- `GeneratePage(filename As String, Title As String, source As String, Keywords As String, Menu As String, UrlGroup As String)` — Generate HTML page file
- `GenerateMenu(DomainName As String, DomainFolder As String, UrlGroupId As Long) As String` — Generate HTML menu markup
- `GenerateMenuBlock(DomainName As String, DomainFolder As String, UrlGroupId As Long, UrlGroup) As String` — Generate HTML menu block for URL group
- `GenerateFooter() As String` — Generate HTML footer markup
- `SaveFile(filename As String, text As String)` — Save text to file
- `ReadFile(filename As String) As String` — Read text from file
- `GetAnchor(ByVal HRef As String, ByVal text As String, Optional CheckHttp As Boolean = False, Optional TargetBlank As Boolean = True) As String` — Generate HTML anchor tag

## omHTMLExport.bas
- `OutputHTMLRaw(sql As String, title As String, filename As String)` — Write the result of a SELECT as an HTML table to a UTF-8 file (values are HTML-encoded)

## omInvoiceFunctions.bas
- `GetStructuredPaymentReference(number As Long, Optional alignLeft As Boolean = True) As String` — Generate structured payment reference number

## omJSON.cls
- `request(url As String, Optional postParameters As String = "", Optional format As ResponseFormat = json) As String` — Send HTTP request and parse JSON response
- `setJsonRoot(rootPath As String)` — Set root path for JSON object navigation
- `getJsonObjectCount()` — Get count of objects in JSON
- `getJsonObjectValue(path As String)` — Get value from JSON object by path
- `getJsonArrayValue(index, key As String)` — Get value from JSON array by index and key
- `ResponseText` (get) — raw response text from HTTP request
- `ResponseJson` (get) — parsed JSON response object
- `ScriptControl` (get) — script control object for JSON processing

## omJsonConverter.bas
- `ParseJson(ByVal JsonString As String) As Object` — Parse JSON string to object
- `ConvertToJson(ByVal JsonValue As Variant, Optional Whitespace As Variant, Optional json_CurrentIndentation As Long = 0) As String` — Convert value to JSON string
- `ParseUtc(utc_UtcDate As Date) As Date` — Parse UTC date value
- `ConvertToUtc(utc_LocalDate As Date) As Date` — Convert local date to UTC
- `ParseIso(utc_IsoString As String) As Date` — Parse ISO 8601 date string
- `ConvertToIso(utc_LocalDate As Date) As String` — Convert date to ISO 8601 format

## omKernalAPIFunctions.bas
*(geen Public Sub/Function/Property gevonden)*

## omKeyObject.cls
*(geen Public Sub/Function/Property gevonden)*

## omKeyValue.cls
*(geen Public Sub/Function/Property gevonden)*

## omLibraryFunctions.bas
- `ExportLibrary(Optional createFolder As Boolean = False, Optional startsWith As String = "om", Optional filenameWithTimestamp As Boolean = False) As String` — Export VBA modules to text files
- `UpdateLibrary(Optional backupComponents As Boolean = True, Optional useUpdatesFolder As Boolean = True, Optional deleteAfterUpdate As Boolean = True)` — Update library components from external files
- `RemoveEqualFiles(sourcePath As String, destinationPath As String, Optional moveFile As Boolean = False)` — Remove or move duplicate files between folders
- `LoadFromText(objectType As AcObjectType, objectName As String, filename As String)` — Load Access object from text file
- `LoadAsAXL(objectType As AcObjectType, objectName As String, filename As String)` — Load Access object from AXL file
- `SaveAsText(objectType As AcObjectType, objectName As String, filename As String)` — Save Access object to text file
- `SaveAsAXL(objectType As AcObjectType, objectName As String, filename As String)` — Save Access object to AXL file
- `ExportFormControlProperties(formName As String, Optional WriteToFile As Boolean = False, Optional controlEscaped As Boolean = False)` — Export form control properties
- `Dump(Optional objectType As AcObjectType = acDefault, Optional silentMode As Boolean = True, Optional destinationPath As String, Optional addTimeStamp As Boolean = True)` — Dump Access objects to files
- `DumpImport(objectType As AcObjectType, Optional silentMode As Boolean = True, Optional sourcePath As String)` — Import dumped objects from files
- `DeleteByObjectType(objectType As AcObjectType, Optional silentMode As Boolean = True, Optional maxNumberOfObjects = 0)` — Delete Access objects by type
- `VBACountModules()` — Count VBA modules in project
- `VBACountLines()` — Count lines of VBA code
- `VBAFindReplace(findString As String, replaceString As String, Optional silentMode As Boolean = True)` — Find and replace in VBA code
- `ListOnlyModulesWithoutLines()` — List empty VBA modules
- `RemoveModulesWithoutLines(Optional silentMode As Boolean = True)` — Remove empty VBA modules
- `DumpAndCommitAndPush(Optional objectType As AcObjectType = acDefault)` — Dump, commit and push to repository
- `CommitAndPush(repositoryPath As String)` — Commit changes and push to repository
- `RemoveDeletedObjects()` — Remove deleted object references
- `ListFormsToTable(Optional tableName As String = "___T_Forms")` — Export form list to table
- `CreateBinaryColumns()` — Create binary data columns in tables
- `GenerateDataMacroAfterInsert(tableName As String) As String` — Generate after-insert data macro
- `GenerateDataMacroAfterUpdate(tableName As String) As String` — Generate after-update data macro

## omLibraryStamp.bas
- `LibraryCheck(Optional manifestFile As String, Optional prefix As String) As String` — One line per om* module that needs attention: NO STAMP, CHANGED LOCALLY, LIBRARY NEWER, LOCAL ONLY; empty = all in order
- `GetModuleHash(moduleName As String) As String` — Calculated hash (8 hex) of a module of the current database
- `GetModuleStamp(moduleName As String) As String` — The stored stamp line (date and sha) of a module
- `NormalizeModuleText(text As String) As String` — The text that is hashed (public for tests)

## omLogging.cls
- `currentdate` (get/let) — current logging date
- `SetCurrentDate(Optional dt As Date = 0, Optional Description As String = "")` — Set current date and description
- `WriteToFile(Optional diffOnly As String = "y", Optional updateCurrentDate As Boolean = True, Optional Description As String = "") As String` — Write log to file
- `ToString(Optional diffOnly As String = "y", Optional updateCurrentDate As Boolean = True, Optional Description As String = "") As String` — Format log as string
- `WriteToStream(Optional diffOnly As String = "y", Optional updateCurrentDate As Boolean = True, Optional Description As String = "") As String` — Write log to stream/string
- `Reset()` — Reset logging state

## omLoggingFunctions.bas
- `StartLogging(Optional enable As Boolean = False)` — Initialize logging system
- `SetLoggingEnable(enable As Boolean)` — Enable or disable logging
- `WriteToLog(Optional diffOnly As String = "y", Optional updateCurrentDate As Boolean = True, Optional Description As String = "") As String` — Write entry to log
- `StopLogging()` — Terminate logging system

## omMSAccessFunctions.bas
- `FormExists(formName As String) As Boolean` — Check if form exists
- `ReportExists(ReportName As String) As Boolean` — Check if report exists
- `TableExists(tableName As String) As Boolean` — Check if table exists
- `QueryExists(queryName As String) As Boolean` — Check if query exists
- `DeleteQuery(queryName As String)` — Delete query by name
- `CreatePassthroughQuery(queryName As String, SQL As String, connection As String)` — Create SQL pass-through query
- `GetQuerySQL(queryName As String) As String` — Retrieve query SQL definition
- `HideNavigationPane()` — Hide Access navigation pane
- `MinimizeNavigationPane()` — Minimize Access navigation pane
- `UnhideNavigationPane()` — Show Access navigation pane
- `ExportAllTables()` — Export all tables to external format
- `ConvertToLocalTables(Optional databaseType As String = ".accdb")` — Convert linked tables to local
- `MakeLinkedTablesLocal()` — Convert all linked tables to local
- `MakeLinkedTableLocal(tableName As String, Optional structureOnly As Boolean = False)` — Convert single linked table to local
- `DropIndexUsingDAO(tableName As String, indexName As String)` — Drop index via DAO
- `CreateIndexUsingDAO(tableName As String, indexName As String, fieldNames As String, Optional setPrimary, Optional setUnique, Optional setDisallowNull)` — Create index via DAO
- `DropIndexUsingSQL(tableName As String, indexName As String)` — Drop index via SQL
- `CreateIndexUsingSQL(tableName As String, indexName As String, fieldNames As String, Optional setPrimary, Optional setUnique, Optional setDisallowNull)` — Create index via SQL
- `DeleteTables(tablePrefix As String)` — Delete tables by name prefix
- `SetAccessProperty(propertyName As String, Value As Variant, Optional propertyType As dao.DataTypeEnum = dbText)` — Set Access database property
- `DeleteTable(tableName As String)` — Delete table by name
- `IsTableLocal(tableName As String) As Boolean` — Check if table is local (not linked)
- `FormFields_Extract()` — Extract form field metadata
- `cmdBarFields_Extract(strcmdBar As String, ctrl As Object)` — Extract command bar field properties
- `newField_ID() As Long` — Generate new field ID
- `RegisterCurrentLocation()` — Register current database location
- `IsRuntimeMode() As Boolean` — Check if running in Access runtime
- `TruncateMSAccessTable(tableName As String)` — Clear all data from table

## omObjectFunctions.bas
- `IsNothing(obj As Object) As Boolean` — Check if object is Nothing
- `NotIsNothing(obj As Object) As Boolean` — Check if object is not Nothing
- `SetAsNothing(obj As Object)` — Set object to Nothing

## omOutlook.cls
- `Mail(EmailAddresses As String, subject As String, body As String, replyAddress As String, Importance As OlImportance, ReadReceipt As Boolean, attachments() As String, Popup As Boolean, Optional bccAddresses, Optional ccAddresses, Optional font, Optional fontSize, Optional keepSignature, Optional useReplyAddressAsOnBehalf)` — Send email via Outlook
- `DoesMailboxExist(Email As String) As Boolean` — Check if mailbox exists in Outlook
- `GetMailbox(mailaddress As String) As Object` — Get mailbox object for email address

## omPDFFunctions.bas
- `SaveReportAsPDF(ReportName As String, filename As String, Optional whereCondition As String = "")` — Export Access report to PDF file

## omPictureFunctions.bas
- `savePict(pImage As Access.Image)` — Save image control to JPG file
- `savePictEMF(pImage As Access.Image)` — Save image control to EMF file

## omPrinterSettingsManager.cls
- `GetSettings()` — Retrieve current printer settings

## omPrinterSettingsManagerFunctions.bas
- `GetPrinterSettings()` — Retrieve and display printer settings information

## omRecordsetFunctions.bas
- `CopyRecord(rsSrc As ADODB.Recordset, rsDst As ADODB.Recordset, Optional PrimaryKeyName As String = "", Optional LinkName As String = "", Optional LinkId As Long = 0) As Long` — Copy field values from source to destination recordset
- `GetList(SQL As String, Optional ColumnDelimiter As String = ", ", Optional RowDelimiter As String = vbCrLf) As String` — Return combined string from query with configurable delimiters

## omReferenceFunctions.bas
- `ListReferences()` — List all application references in the Immediate window
- `AddReference(filename As String)` — Add a reference from file path
- `FindReference(Name As String) As Reference` — Search for and return a reference by name or path
- `RemoveReference(filename As String)` — Remove a reference by name or path
- `AddVBIDEReference()` — Add VB IDE reference from system path
- `AddAdoDbReference()` — Add ADODB reference from system path
- `AddScriptingReference()` — Add Scripting reference from system path

## omReportFunctions.bas
- `ExportToPDF(ReportName As String, filename As String, Optional whereCondition As String, Optional windowMode As AcWindowMode = acWindowNormal, Optional outputQuality As AcExportQuality = acExportQualityPrint)` — Export report to PDF file with optional filtering and quality settings
- `UpdateVisibleByTag(Report As Report, tag As String, state As Boolean)` — Set control visibility based on tag value

## omResult.cls
- `Ok(Optional reason As String)` — Mark result as successful
- `Failed(Optional reason As String)` — Mark result as failed (also the state of a new object)
- `IsSuccess() As Boolean` — True after Ok
- `IsFailed() As Boolean` — True until Ok is called

## omSeqFunctions.bas
- `omSeqItemLevelToString(Value As omSeqItemLevel) As String` — Convert omSeqItemLevel enum to string representation
- `SendSeqItem(Optional seqItem As omSeqItem)` — Send structured logging item to Seq server
- `OpenSeq(Optional filter As String = "")` — Open Seq logging interface with optional filter
- `GetUtcDateTime()` — Get current UTC date/time from system
- `DtIsoToAccess(myisodate As String) As Date` — Convert ISO format date string to Access date
- `DtAccessToIso(myaccdate As Date, Optional myUtcDate As Date = 0) As String` — Convert Access date to ISO format string with timezone

## omSeqItem.cls
- `ToJson(Optional formatted As Boolean = False) As String` — Serialize logging item to JSON format

## omSFTP.bas
- `SftpPut(host As String, user As String, localFile As String, remotePath As String, Optional password As String, Optional privateKeyFile As String, Optional hostKey As String, Optional port As Long, Optional winscpPath As String) As String` — Upload a file over SFTP with WinSCP (password or key) or OpenSSH sftp.exe (key); returns "" when it worked, otherwise the reason
- `SftpFindWinScp(Optional winscpPath As String) As String` — Full path of WinSCP.com, or "" when not found
- `SftpBuildWinScpScript(host As String, user As String, localFile As String, remotePath As String, password As String, privateKeyFile As String, hostKey As String, port As Long) As String` — The WinSCP script text (public for tests; contains the password)

## omSourceObject.cls
- `Load(obj As Object, Optional objType As AcObjectType = 0)` — Load and track form or report object metadata

## omSourceObjectControl.cls
- `Load(soId As Long, c As omControl)` — Load and track control metadata for a source object

## omSQLConnector.cls
- `ExecuteCommandTextNonQuery(commandText As String, Optional Parameter1..N As Variant = Null)` — Execute parameterized SQL command without returning results
- `ExecuteStoredProcedureNonQuery(storedProcedureName As String, Optional Parameter1..N As Variant = Null)` — Execute stored procedure without returning results
- `ExecuteSQLPassThrough(commandText As String, queryName As String)` — Set and execute pass-through query definition
- `Execute(commandText As String)` — Execute SQL command on active connection
- `ActiveConnection` (get) — active ADODB connection object
- `ConnectionString` (get/let) — current connection string; setting opens a new connection

## omSQLConnectorFunctions.bas
- `ConnectorTest()` — Test SQL connector with parameterized update commands

## omSQLDataAdapter.cls
- `spExecuteNonQuery(storedProcedureName As String, Optional Parameter1..N As Variant = Null)` — Execute stored procedure without returning results
- `spExecuteScalar(storedProcedureName As String, Optional Parameter1..N As Variant = Null) As Variant` — Execute stored procedure and return first scalar value
- `sqlOpenRecordset(SQLStatement As String) As ADODB.Recordset` — Execute SQL query and return recordset
- `sqlExecute(SQLStatement As String)` — Execute SQL command without returning results
- `CloseConnection()` — Close database connection

## omSQLFunctions.bas
- `WhereAnd(whereClause1 As String, whereClause2 As String, Optional embraceWhereClause1 = False, Optional embraceWhereClause2 = False)` — Combine two WHERE clauses with AND
- `WhereOr(whereClause1 As String, whereClause2 As String, Optional embraceWhereClause1 = False, Optional embraceWhereClause2 = False)` — Combine two WHERE clauses with OR
- `InnerJoin(joinTable As String, Table1 As String, Field1 As String, Table2 As String, Field2 As String)` — Build INNER JOIN clause
- `LeftJoin(joinTable As String, Table1 As String, Field1 As String, Table2 As String, Field2 As String)` — Build LEFT JOIN clause
- `BuildSQL(ByVal selectClause As String, Optional fromClause, Optional whereClause, Optional orderClause, Optional insertClause, Optional deleteClause) As String` — Construct complete SQL statement from clause components
- `GetSelect(SQL As String) As String` — Extract SELECT clause from SQL statement
- `GetFrom(SQL As String) As String` — Extract FROM clause from SQL statement
- `GetWhere(SQL As String) As String` — Extract WHERE clause from SQL statement
- `GetOrderBy(ByVal SQL As String) As String` — Extract ORDER BY clause from SQL statement
- `GetAccessDate(dt As Date)` — Format date for Access SQL syntax
- `GetSQLServerDate(dt As Date)` — Format date for SQL Server syntax
- `GetAccessString(str As String, Optional prefixWildcard = True, Optional suffixWildcard = True)` — Format string for Access SQL LIKE operator
- `GetSQLServerString(str As String, Optional prefixWildcard = True, Optional suffixWildcard = True)` — Format string for SQL Server LIKE operator
- `GetAccessWeekFunction()` — Get Access SQL function for week extraction
- `GetSQLServerWeekFunction()` — Get SQL Server function for week extraction
- `GetSQLClausePos(ByVal SQL As String, clause As String, Optional startPos As Long = 1) As Long` — Find position of SQL clause in statement
- `ReplaceSQLClause(ByVal SQL As String, clause As String, Optional replaceString As String = "", Optional startPos As Long = 0) As String` — Replace SQL clause with new text
- `NormalizeSQL(ByVal SQL As String) As String` — Clean and normalize SQL statement whitespace
- `WhereOperatorForString(str As String) As String` — Determine = or LIKE operator based on wildcard characters
- `ReplaceDatabaseInUseClause(ByVal SQL As String, sourceDatabase As String, destinationDatabase As String) As String` — Replace database name in USE clause

## omSQLObject.cls
- `SplitSQL(SQL As String)` — Parse SQL statement into SELECT, FROM, WHERE, ORDER BY components
- `BuildSQL(Optional selectClause, Optional fromClause, Optional whereClause, Optional orderByClause) As String` — Construct SQL from stored components

## omSQLServerFunctions.bas
- `GetPrimaryKeyScript(tableName As String) As String` — Generate SQL Server primary key constraint DDL from Access table definition
- `TableCreateDDL(TableDef As dao.TableDef) As String` — Generate SQL Server CREATE TABLE DDL from Access table definition
- `ExecuteScript()` — Execute SQL Server DDL alteration script

## omSSMAAConnector.bas
- `LocalCon() As ADODB.connection` — Return current Access database connection
- `ServerCon(Optional ConnectionType = SQLNCLI, Optional ODBCConnection As Boolean = False, Optional encryptType = EncryptOptional) As String` — Get or create SQL Server connection string
- `LinkUsingSSMA(Optional Group As String = "", Optional ConnectionType = SQLNCLI, Optional SavePassword As Boolean = False, Optional alwaysUpdate As Boolean = False, Optional encryptType = EncryptOptional) As Boolean` — Link SQL Server tables using SSMA configuration
- `LinkTableViewUsingSSMA(existingTableName As String, toLinkName As String, Optional ConnectionType, Optional encryptType)` — Link specific SQL Server table or view
- `CreateTable()` — Create SSMAA_ODBC_Tables configuration table
- `PopulateTable()` — Populate SSMAA configuration table from existing SQL Server tables
- `GetSQLNCLIVersion() As String` — Detect installed SQL Server Native Client version
- `GetSQLODBCVersion() As String` — Detect installed SQL Server ODBC driver version
- `GetConnectionStringByProperty(Optional tableName, Optional databaseName, Optional serverName, Optional dsnName, Optional GroupName, Optional ConnectionType, Optional ODBCConnection, Optional encryptType) As String` — Build connection string from SSMAA_ODBC_Tables configuration
- `ImportTables(Optional storageTable As String = "SSMAA_ODBC_Tables", Optional overwrite As Boolean = False)` — Import SQL Server table list into configuration
- `LinkPassthroughQueries(Optional ConnectionType, Optional encryptType)` — Link SQL Server pass-through queries
- `PassthroughQueriesReplaceDatabaseName(sourceDatabase As String, destinationDatabase As String, Optional schema As String = "dbo")` — Update pass-through queries with new database name
- `PassthroughQueriesUpdateDescriptionWithGroup()` — Update pass-through query descriptions with group information
- `LinkDeleteTables(Optional Group As String = "", Optional AttachOnly As Boolean = False)` — Delete linked tables for a group
- `LinkDeletePTs()` — Delete all pass-through queries
- `SetSQLConnectorConnectionString(tableName As String, Optional encryptType)` — Set global SQL connector connection string from table configuration
- `DeleteSSMAABackupTables()` — Delete SSMA backup table copies
- `UpdateSSMAAGroups()` — Update SSMA group configuration
- `IsConnectingPossible(tableName As String, Optional connectionTimout As Long = 5, Optional encryptType) As Boolean` — Test connection to SQL Server without raising error
- `DeleteSSMAATables(Optional Group As String = "")` — Delete SSMA configuration records for a group
- `LinkMSAccess(Optional linkLocal As Boolean = False)` — Link local Access tables or database

## omSSMAConnector_Test.bas
- `ChangeDatabase(GroupName As String, sourceDatabase As String, destinationDatabase As String)` — Swap source/destination database in SSMA configuration and relink tables
- `LinkMySQL()` — Link clientOnlineNW group with password saving enabled
- `GetCurrentConnectedDatabase(Optional tableName As String = "T_Medewerkers") As String` — Extract connected database name from linked table properties
- `GetCurrentConnectedServer(Optional tableName As String = "T_Medewerkers") As String` — Extract connected server name from linked table properties
- `HerlinkACenter()` — Relink ACenter_client group using SQLODBC connection type
- `ChangeServer(sourceServer As String, destinationServer As String)` — Swap source/destination server in SSMAA configuration
- `TestAdodbRecordsetConnectionstringEncrypt()` — Test ADODB recordset with encrypted connection string

## omStartup.bas
- `Main(Optional linkLocal As Boolean = False) As Integer` — Opens application startup form based on user role
- `LinkAccessDatabase() As Variant` — Links to local MS Access database
- `LinkSSMADatabase() As Variant` — Links to SSMA database with SQL authentication
- `LinkSSMADatabaseSQLODBC() As Variant` — Links to SSMA database via SQL ODBC connection

## omStringFunctions.bas
- `CleanString(Value As Variant) As String` — Trims null/empty values to empty string
- `IsNullOrEmpty(Value As Variant) As Boolean` — Tests if value is null or empty
- `IsNullOrEmptyOrZero(Value As Variant) As Boolean` — Tests if value is null, empty, or zero
- `NotIsNullOrEmptyOrZero(Value As Variant) As Boolean` — Tests if value has non-zero content
- `NotIsNullOrEmpty(Value As Variant) As Boolean` — Tests if value is not null/empty
- `ParseValue(arguments As Variant, valueName As String, Optional assignChar, Optional splitchar, Optional notFoundValue As Variant) As Variant` — Extracts key=value from delimited string
- `StringSplit(Value As String, splitchar As String, Optional removeEmptyValues As Boolean) As String()` — Splits string into array
- `StringSplitGetByIndex(Value As String, splitchar As String, index As Long) As Variant` — Returns split element by index
- `IsIdNullOrZero(Value As Variant) As Boolean` — Tests if ID is null or zero
- `NotIsIdNullOrZero(Value As Variant) As Boolean` — Tests if ID is not null/zero
- `IsIdNullOrEmptyOrZero(Value As Variant) As Boolean` — Tests if ID is null, empty, or zero
- `NotIsIdNullOrEmptyOrZero(Value As Variant) As Boolean` — Tests if ID has non-zero value
- `AreIdsEqual(id1 As Variant, id2 As Variant) As Boolean` — Compares two ID values
- `AreStringsEqual(string1 As Variant, string2 As Variant) As Boolean` — Compares two strings
- `StringFormat(source As String, replace0 As String, Optional replace1-12 As Variant) As String` — Replaces placeholders {0}-{12} in string
- `RemoveChars(source As Variant, findPattern As String, replaceString As String)` — Removes characters from string
- `KeepChars(source As Variant, keepPattern As String, replaceString As String)` — Keeps only specified characters in string
- `ContainsString(source As Variant, containString As String, Optional useDelimiter As String) As Long` — Finds substring position with optional delimiter
- `GetEnglishPlural(Value As String) As String` — Returns English plural form
- `ReplaceCharPattern(source As Variant, findPattern As String, Replacement As String) As Variant` — Replaces characters matching pattern
- `CleanCommunication(Value As Variant) As Variant` — Removes formatting characters from communication string
- `ReverseCommunication(Value As Variant) As Variant` — Reverses cleaned communication string
- `replaceString(ByVal Value As String, oldString As String, newString As String) As String` — Replaces all occurrences of substring
- `GetvbCrLf() As String` — Returns CRLF constant
- `GetDelimitedValue(Value As String, Optional Position As Long, Optional delimiter As String, Optional EmbraceChar As String)` — Extracts delimited value by position
- `GetScrambleNumbers(Length As Long) As String` — Generates random digit string
- `IsStringInPattern(Value As String, pattern As String) As Boolean` — Tests if string is in pattern
- `CleanStringUsingPattern(ByVal Value As Variant, Optional findPattern, Optional replaceString, Optional replaceDoubleBlanks As Boolean, Optional trimResult As Boolean) As String` — Cleans string using character pattern replacement
- `StringPadLeft(s As String, totalLength As Long, paddingChar As String) As String` — Pads string on left to specified length
- `StringPadRight(s As String, totalLength As Long, paddingChar As String) As String` — Pads string on right to specified length
- `StringPad(s As String, totalLength As Long, paddingChar As String, paddingMode As StringPaddingMode) As String` — Pads string to specified length with mode
- `ContainsCharFromPattern(s As String, pattern As String) As Boolean` — Tests if string contains any character from pattern
- `ExtractLong(text As Variant, findtext As String) As Long` — Extracts long integer from text
- `GetSubstringAfterChar(s As String, Optional c As String, Optional autoTrim As Boolean) As String` — Returns substring after last occurrence of character
- `ByteArrayToHexString(ba() As Byte) As String` — Converts byte array to hex string
- `ByteArrayToString(ba() As Byte) As String` — Converts byte array to numeric string
- `CheckIllegalCharInString(strCheck As String) As String` — Checks for illegal characters in string

## omStringTagFunctions.bas
- `GetOpenTagPosition(ByVal startPos As Long, ByVal text As String) As Long` — Finds position of next opening tag
- `GetTag(ByVal startPos As Long, ByVal text As String) As tag` — Extracts tag structure from text at position
- `GetNextTag(ByVal startPos, ByVal text As String) As tag` — Gets next tag structure from starting position
- `GetTags(ByVal text As String) As tag()` — Extracts all tags from text into array

## omSystemProcessFunctions.bas
*(geen Public Sub/Function/Property gevonden)*

## omTableConnector.cls
- `Connect(ConnectionType As omTableConnectionType)` — Establishes table connections based on connection type

## omTableFunctions.bas
- `TableAddId(tableName As String)` — Adds auto-increment ID field to table
- `AddField(tableName As String, fieldName As String, FieldType As dao.DataTypeEnum, Optional AutoIncrement As Boolean)` — Adds field to table
- `GenerateSelect(tableName As String, Optional includeTableName As Boolean, Optional excludeFields As String) As String` — Generates SELECT statement for table
- `GenerateInsert(tableName As String, Optional idFieldName As String, Optional excludeFields As String, Optional excludeValuesPart As Boolean) As String` — Generates INSERT statement for table
- `GenerateUpdate(tableName As String, Optional idFieldName As String, Optional excludeFields As String) As String` — Generates UPDATE statement for table

## omTesseractOCR.bas
- `GetTxtFromImage(filename As String) As String` — Extracts text from image using Tesseract OCR
- `GetTxtFromImageTest()` — Tests OCR extraction (debug function)

## omTimeRange.cls
- `StartTime` (get/let) As Date — start time
- `EndTime` (get/let) As Date — end time
- `GetStartTimeInSeconds() As Long` — Returns start time converted to seconds
- `GetEndTimeInSeconds() As Long` — Returns end time converted to seconds

## omTranslationEngine.cls
- `Translate(obj As Object, LanguageId As Long)` — Translates form/report captions and controls
- `ClearAll()` — Truncates all translation-related tables
- `IndexAll()` — Indexes all forms and reports for translation
- `IndexForm(formName As String)` — Indexes single form for translation
- `IndexReport(ReportName As String)` — Indexes single report for translation
- `IndexByObject(obj As Object, objType As AcObjectType)` — Indexes form/report object for translation
- `InsertTranslateCode(Optional TranslationClassName As String)` — Injects translation code into form/report Open events

## omTranslationEngineFunctions.bas
- `GetCurrentLanguage() As Long` — Returns current selected language ID
- `UpdateTranslations()` — Updates all translations via IndexAll
- `ClearTranslations()` — Clears all translations via ClearAll
- `InjectCodeInFormsAndReports()` — Injects translation code into all forms and reports

## omUser.cls
- `LoginByNameOnly(LoginName, Optional windowsLogon As Boolean, Optional addWindowsLogon As Boolean)` — Logs in by username only
- `LoginById(UserId As Long, Password As String) As Boolean` — Logs in by ID and password
- `LoginByName(LoginName As String, Password As String) As Boolean` — Logs in by name and password
- `UpdatePassword(Password As String)` — Updates password for current user
- `LoggedIn() As Boolean` — Tests if user is currently logged in

## omUserFunctions.bas
- `GetWindowsUserName() As String` — Returns Windows login username
- `GetWindowsComputerName() As String` — Returns Windows computer name
- `LoginWithWindowsUser(Optional addWindowsLogon As Boolean)` — Logs in with Windows username
- `UseWindowsUser()` — Sets current user to Windows username
- `UserLock()` — Opens UserLock form if exists
- `AuthenticateUser()` — Authenticates user via configured method (Windows or form)
- `RecordsetUpdateCreateTracking(rs As ADODB.Recordset, Optional UpdateModify)` — Updates recordset with creation audit tracking
- `RecordsetUpdateModifyTracking(rs As ADODB.Recordset)` — Updates recordset with modification audit tracking
- `GetUserId() As Long` — Returns current user ID
- `HasRole(roleCode As String) As Boolean` — Tests if user has specified role code

## omUserRole.cls
- `Name` (get/let) — role name
- `Code` (get/let) — role code
- `LoadById(Id As Long)` — Loads role data by ID

## omUser_Edan.cls
- `LoginById(UserId As Long, Password As String)` — Logs in by user ID and password
- `LoginByName(LoginName As String, Password As String)` — Logs in by name and password

## omVATFunctions.bas
- `OpenBEVATNumberLookupWebsite(VATNumber As Variant)` — Open the KBO public search for a Belgian VAT number
- `CleanBEVatNumber(VATNumber As Variant) As String` — Digits only (removes spaces, dots, commas, BE)
- `IsValidBEVatNumber(VATNumber As Variant) As Boolean` — BE prefix, 10 digits and Mod 97 check
- `FormatBEVatNumber(VATNumber As Variant, Optional formatString As String) As String` — Format a valid number, e.g. BE 0123.456.789

## omWinSpoolFunctions.bas
- `GetPaperListText() As String` — Paper number and name of every paper of the default printer, one per line
- `GetPaperList()` — Show that list in a message box
- `GetPaperSize(paperName As String) As Long` — Paper number of a paper name on the default printer (-1 = not found)

## omZipFunctions.bas
- `ZipFolder(ByVal folderName As String, Optional filenameZip As String) As String` — Creates ZIP file from folder
- `UnzipToFolder(ByVal zipFileName As String, Optional fileNameFolder As String) As String` — Extracts ZIP file to folder

---

Coverage: 79/79 bestanden verwerkt (51 .bas + 28 .cls). Vier bestanden bevatten geen Public members (`omGlobalDeclarations.bas`, `omKernalAPIFunctions.bas`, `omKeyObject.cls`, `omKeyValue.cls`, `omSystemProcessFunctions.bas`) — expliciet zo gemarkeerd i.p.v. overgeslagen.
</content>
