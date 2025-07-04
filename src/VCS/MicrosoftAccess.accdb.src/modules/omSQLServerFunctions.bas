Attribute VB_Name = "omSQLServerFunctions"
Option Compare Database
Option Explicit

Public Function GetPrimaryKeyScript(tableName As String) As String
Dim db As dao.Database
Dim tblDef As TableDef
Dim i As Long
Dim idx As index
Dim fld As Field
Dim str As String

    Set db = CurrentDb
    Set tblDef = db.TableDefs(tableName)
    For Each idx In tblDef.Indexes
        If idx.Primary Then
            str = "ALTER TABLE " & tableName & " ADD CONSTRAINT "
            str = str & idx.Name & " PRIMARY KEY CLUSTERED (" 'NONCLUSTERED
            For Each fld In idx.Fields
                str = str & fld.Name & IIf((fld.Attributes And dbDescending) = dbDescending, " DESC", "") & ","
            Next
            str = Left(str, Len(str) - 1)
            str = str & ") WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
            GetPrimaryKeyScript = str
            Exit Function
        End If
    Next
End Function


' https://msdn.microsoft.com/en-us/library/ms714540(v=vs.85).aspx
' http://blogs.msdn.com/b/ssma/archive/2011/03/06/access-to-sql-server-migration-understanding-data-type-conversions.aspx
' https://technet.microsoft.com/en-us/library/cc917602.aspx?f=255&MSPPError=-2147217396
' http://stackoverflow.com/questions/8454393/comparing-data-type-of-sql-server-and-ms-access-in-c-sharp
' http://www.codeproject.com/Questions/305072/how-to-generate-sql-table-scripts-through-query-in
' http://www.everythingaccess.com/tutorials.asp?ID=Field-type-names-%28JET%2C-DDL%2C-DAO-and-ADOX%29
' http://www.w3schools.com/asp/ado_datatypes.asp

Public Function TableCreateDDL(TableDef As dao.TableDef) As String

         Dim fldDef As dao.Field
         Dim FieldIndex As Integer
         Dim fldName As String, fldDataInfo As String
         Dim DDL As String
         Dim tableName As String
         Dim pkScript As String


         tableName = TableDef.Name
         tableName = Replace(tableName, " ", "_")
         pkScript = GetPrimaryKeyScript(tableName)
         DDL = "create table " & tableName & "(" & vbCrLf
         With TableDef
            For FieldIndex = 0 To .Fields.Count - 1
                Set fldDef = .Fields(FieldIndex)
                With fldDef
                    fldName = .Name
                    fldName = Replace(fldName, " ", "_")
                    Select Case .Type
                        'Case DAO.DataTypeEnum.dbAttachment
                        Case dao.DataTypeEnum.dbBigInt
                           fldDataInfo = "BIGINT"
                        Case dao.DataTypeEnum.dbBinary
                           fldDataInfo = "BINARY"
                        Case dao.DataTypeEnum.dbBoolean
                           fldDataInfo = "BIT"
                        Case dao.DataTypeEnum.dbByte
                           fldDataInfo = "TINYINT"
                        Case dao.DataTypeEnum.dbChar
                           fldDataInfo = "CHAR"
                        Case dao.DataTypeEnum.dbCurrency
                           fldDataInfo = "MONEY"
                        Case dao.DataTypeEnum.dbDate
                           fldDataInfo = "DATETIME"
                        Case dao.DataTypeEnum.dbDecimal
                           fldDataInfo = "FLOAT"
                        Case dao.DataTypeEnum.dbDouble
                           fldDataInfo = "FLOAT"
                        Case dao.DataTypeEnum.dbFloat
                           fldDataInfo = "FLOAT"
                        Case dao.DataTypeEnum.dbGUID
                           fldDataInfo = "uniqueidentifier"
                        Case dao.DataTypeEnum.dbInteger
                           fldDataInfo = "smallint"
                        Case dao.DataTypeEnum.dbLong
                           fldDataInfo = "int"
                        'Case DAO.DataTypeEnum.dbLongBinary
                        Case dao.DataTypeEnum.dbMemo
                           fldDataInfo = "VARCHAR(MAX)"
                        'Case DAO.DataTypeEnum.dbNumeric
                        Case dao.DataTypeEnum.dbSingle
                           fldDataInfo = "REAL"
                        Case dao.DataTypeEnum.dbText
                           fldDataInfo = "VARCHAR(" & .Size & ")"
                        'Case DAO.DataTypeEnum.dbTime
                        'Case DAO.DataTypeEnum.dbTimeStamp
                        Case dao.DataTypeEnum.dbVarBinary
                           fldDataInfo = "VARBINARY(MAX)"
                    End Select
                    If .Required Or InStr(1, pkScript, .Name) > 1 Then
                        fldDataInfo = fldDataInfo & " NOT NULL"
                    End If
                    ' AllowZerolength => constraint
                    ' DefaultValue => Constraint
                End With
                If FieldIndex > 0 Then
                    DDL = DDL & ", " & vbCrLf
                End If
                DDL = DDL & "  " & fldName & " " & fldDataInfo
            Next FieldIndex
         End With
         DDL = DDL & ")"
         TableCreateDDL = DDL
End Function


Sub ExportAllTableCreateDDL()

    Dim lTbl As Long
    Dim dBase As dao.Database
    Dim Handle As Integer

    Set dBase = CurrentDb

    Handle = FreeFile

    Open GetDesktopFolder & "\TableCreateDDL.txt" For Output Access Write As #Handle

    For lTbl = 0 To dBase.TableDefs.Count - 1
         'If the table name is a temporary or system table then ignore it
        If Left(dBase.TableDefs(lTbl).Name, 1) = "~" Or _
        Left(dBase.TableDefs(lTbl).Name, 4) = "MSYS" Then
             '~ indicates a temporary table
             'MSYS indicates a system level table
        Else
            'If InStr(1, dBase.TableDefs(lTbl).Name, "PUB_") > 0 Then
            If dBase.TableDefs(lTbl).Connect = "" Then
                Print #Handle, TableCreateDDL(dBase.TableDefs(lTbl))
                Print #Handle, GetPrimaryKeyScript(dBase.TableDefs(lTbl).Name)
            End If
        End If
    Next lTbl
    Close Handle
    Set dBase = Nothing
End Sub



Public Sub ExecuteScript()
Dim cmd As New ADODB.Command

    cmd.commandText = "ALTER TABLE dbo.PriceLists ADD IsVisibleInSO bit NOT NULL CONSTRAINT DF_PriceLists_Test DEFAULT 0"
    cmd.ActiveConnection = GetConnectionStringByProperty(tableName:="PriceLists", ODBCConnection:=False)
    cmd.Execute

End Sub
