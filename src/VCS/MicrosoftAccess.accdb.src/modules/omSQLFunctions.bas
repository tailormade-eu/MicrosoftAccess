Attribute VB_Name = "omSQLFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=85c05af6

' Limits of GetSelect/GetFrom/GetWhere/GetOrderBy/GetSQLClausePos/ReplaceSQLClause:
' they cut the SQL text by searching for the keywords FROM, WHERE, GROUP BY, HAVING, ORDER BY and UNION.
' Reliable for flat queries only. They do NOT understand:
' - subqueries: a FROM/WHERE/ORDER BY inside brackets is taken for the outer one
' - text values: 'GROUP BY' inside a string literal ends the clause too early
' NormalizeSQL replaces every semicolon and double space, also inside string literals.
' ReplaceSQLClause only returns the text after the keyword; replaceString is not used yet.

Public Function WhereAnd(whereClause1 As String, whereClause2 As String, Optional embraceWhereClause1 = False, Optional embraceWhereClause2 = False)
    If Len(whereClause1) > 0 And embraceWhereClause1 Then
        whereClause1 = "(" & whereClause1 & ")"
    End If
    If Len(whereClause2) > 0 And embraceWhereClause2 Then
        whereClause2 = "(" & whereClause2 & ")"
    End If
    If Len(whereClause1) = 0 Then
        WhereAnd = whereClause2
    ElseIf Len(whereClause2) = 0 Then
        WhereAnd = whereClause1
    ElseIf Len(whereClause1) <> 0 And Len(whereClause2) <> 0 Then
        WhereAnd = whereClause1 & " AND " & whereClause2
    End If
End Function

Public Function WhereOr(whereClause1 As String, whereClause2 As String, Optional embraceWhereClause1 = False, Optional embraceWhereClause2 = False)
    If Len(whereClause1) > 0 And embraceWhereClause1 Then
        whereClause1 = "(" & whereClause1 & ")"
    End If
    If Len(whereClause2) > 0 And embraceWhereClause2 Then
        whereClause2 = "(" & whereClause2 & ")"
    End If
    If Len(whereClause1) = 0 Then
        WhereOr = whereClause2
    ElseIf Len(whereClause2) = 0 Then
        WhereOr = whereClause1
    ElseIf Len(whereClause1) <> 0 And Len(whereClause2) <> 0 Then
        WhereOr = whereClause1 & " OR " & whereClause2
    End If
End Function

Public Function InnerJoin(joinTable As String, Table1 As String, Field1 As String, Table2 As String, Field2 As String)
Dim strJoin As String
    strJoin = "(!JoinTable! INNER JOIN [!Table2!] ON [!Table1!].[!Field1!]=[!Table2!].[!Field2!])"
    strJoin = Replace(strJoin, "!JoinTable!", IIf(joinTable = "", "[" & Table1 & "]", joinTable))
    strJoin = Replace(strJoin, "!Table1!", Table1)
    strJoin = Replace(strJoin, "!Field1!", Field1)
    strJoin = Replace(strJoin, "!Table2!", Table2)
    strJoin = Replace(strJoin, "!Field2!", Field2)
    InnerJoin = strJoin
End Function

Public Function LeftJoin(joinTable As String, Table1 As String, Field1 As String, Table2 As String, Field2 As String)
Dim strJoin As String
    strJoin = "(!JoinTable! LEFT JOIN [!Table2!] ON [!Table1!].[!Field1!]=[!Table2!].[!Field2!])"
    strJoin = Replace(strJoin, "!JoinTable!", IIf(joinTable = "", "[" & Table1 & "]", joinTable))
    strJoin = Replace(strJoin, "!Table1!", Table1)
    strJoin = Replace(strJoin, "!Field1!", Field1)
    strJoin = Replace(strJoin, "!Table2!", Table2)
    strJoin = Replace(strJoin, "!Field2!", Field2)
    LeftJoin = strJoin
End Function

Public Function BuildSQL(ByVal selectClause As String, Optional ByVal fromClause As String = "", Optional ByVal whereClause As String = "", Optional ByVal orderClause As String = "", Optional ByVal insertClause As String = "", Optional ByVal deleteClause As String = "") As String

    selectClause = NormalizeSQL(selectClause)
    fromClause = NormalizeSQL(fromClause)
    whereClause = NormalizeSQL(whereClause)
    insertClause = NormalizeSQL(insertClause)
    deleteClause = NormalizeSQL(deleteClause)

    If Len(selectClause) > 0 Then
        BuildSQL = IIf(Left(selectClause, 6) <> "select", "SELECT ", " ") & selectClause
    End If
    If Len(deleteClause) > 0 Then
        BuildSQL = IIf(Left(deleteClause, 6) <> "delete", "DELETE ", " ") & deleteClause
    End If
    If Len(insertClause) > 0 Then
        BuildSQL = "INSERT INTO " & insertClause & " " & BuildSQL
    End If
    If Len(fromClause) > 0 Then
        BuildSQL = BuildSQL & IIf(Left(fromClause, 4) <> "from", " FROM ", " ") & fromClause
    End If
    If Len(whereClause) > 0 Then
        BuildSQL = BuildSQL & IIf(Left(whereClause, 5) <> "where", " WHERE ", " ") & whereClause
    End If
    If Len(orderClause) > 0 Then
        BuildSQL = BuildSQL & IIf(Left(orderClause, 8) <> "order by", " ORDER BY ", " ") & orderClause
    End If
    BuildSQL = Replace(BuildSQL, "  ", " ")
End Function


Public Function GetSelect(ByVal sql As String, Optional skipnormalization As Boolean = False) As String
Dim posFrom As Long

    ' skipnormalization = True: the caller promises the SQL is already normalized (NormalizeSQL).
    ' We normalize once here and use skipnormalization:=True for every inner call, so all
    ' positions refer to the same string.
    If skipnormalization = False Then
        sql = NormalizeSQL(sql)
    End If
    posFrom = GetSQLClausePos(sql, "FROM", skipnormalization:=True)
    If posFrom = 0 Then
        posFrom = Len(sql) + 1
    End If
    GetSelect = ReplaceSQLClause(Mid(sql, 1, posFrom - 1), "SELECT", "", skipnormalization:=True)

End Function
Public Function GetFrom(ByVal sql As String, Optional skipnormalization As Boolean = False) As String
Dim posFrom As Long
Dim posEnd As Long

    If skipnormalization = False Then
        sql = NormalizeSQL(sql)
    End If
    posFrom = GetSQLClausePos(sql, "FROM", skipnormalization:=True)
    If posFrom = 0 Then
        Exit Function
    End If
    ' the FROM clause ends at the FIRST of the clauses that can follow it
    posEnd = FirstClausePos(sql, posFrom + 1, "WHERE", "GROUP BY", "HAVING", "ORDER BY", "UNION")
    If posEnd = 0 Then
        posEnd = Len(sql) + 1
    End If
    GetFrom = ReplaceSQLClause(Mid(sql, posFrom, posEnd - posFrom), "FROM", "", skipnormalization:=True)

End Function

Public Function GetWhere(ByVal sql As String, Optional skipnormalization As Boolean = False) As String
Dim posWhere As Long
Dim posEnd As Long

    If skipnormalization = False Then
        sql = NormalizeSQL(sql)
    End If
    posWhere = GetSQLClausePos(sql, "WHERE", skipnormalization:=True)
    If posWhere = 0 Then
        Exit Function
    End If
    posEnd = FirstClausePos(sql, posWhere + 1, "GROUP BY", "HAVING", "ORDER BY", "UNION")
    If posEnd = 0 Then
        posEnd = Len(sql) + 1
    End If
    GetWhere = ReplaceSQLClause(Mid(sql, posWhere, posEnd - posWhere), "WHERE", "", skipnormalization:=True)

End Function

Public Function GetOrderBy(ByVal sql As String, Optional skipnormalization As Boolean = False) As String
Dim posOrderBy As Long


'SELECT Course.Course_ID, Course.Course_Year
'FROM Course
'where
'GROUP BY Course.Course_ID, Course.Course_Year
'HAVING (((Course.Course_Year) = [1]))
'ORDER BY Course.Course_Year;
    If skipnormalization = False Then
        sql = NormalizeSQL(sql)
    End If
    posOrderBy = GetSQLClausePos(sql, "ORDER BY", skipnormalization:=True)
    If posOrderBy <> 0 Then
        GetOrderBy = ReplaceSQLClause(Mid(sql, posOrderBy), "ORDER BY", "", skipnormalization:=True)
    End If


End Function

Public Function GetAccessDate(dt As Date)
    GetAccessDate = "#" & Date & "#"
End Function
Public Function GetSQLServerDate(dt As Date)
    GetSQLServerDate = "'" & format(dt, "yyyy-mm-dd hh:mm:ss") & "'"
End Function
Public Function GetAccessString(str As String, Optional prefixWildcard As Boolean = True, Optional suffixWildcard As Boolean = True)
    GetAccessString = Chr(34) & IIf(prefixWildcard, "*", "") & str & IIf(suffixWildcard, "*", "") & Chr(34)
End Function
Public Function GetSQLServerString(str As String, Optional prefixWildcard As Boolean = True, Optional suffixWildcard As Boolean = True)
    GetSQLServerString = "'" & IIf(prefixWildcard, "%", "") & str & IIf(suffixWildcard, "%", "") & "'"
End Function

Public Function GetAccessWeekFunction()
    GetAccessWeekFunction = "datepart(" & Chr(34) & "ww" & Chr(34) & ",<<Value>>)"
End Function

Public Function GetSQLServerWeekFunction()
    GetSQLServerWeekFunction = "datepart(wk,<<Value>>)"
End Function

Public Function GetSQLClausePos(ByVal sql As String, clause As String, Optional startPos As Long = 1, Optional skipnormalization As Boolean = False) As Long
Dim pos As Long

    If skipnormalization = False Then
        sql = NormalizeSQL(sql)
    End If
    pos = InStr(startPos, sql, " " & clause & " ")
    If pos = 0 Then
        pos = InStr(startPos, sql, clause & " ")
    End If

    GetSQLClausePos = pos
End Function

' Position of the first of the given clauses at or after startPos (0 = none found).
' Expects normalized SQL.
Private Function FirstClausePos(sql As String, startPos As Long, ParamArray clauses() As Variant) As Long
Dim i As Long
Dim pos As Long

    For i = LBound(clauses) To UBound(clauses)
        pos = GetSQLClausePos(sql, CStr(clauses(i)), startPos, True)
        If pos <> 0 Then
            If FirstClausePos = 0 Or pos < FirstClausePos Then
                FirstClausePos = pos
            End If
        End If
    Next i
End Function

Public Function ReplaceSQLClause(ByVal sql As String, clause As String, Optional replaceString As String = "", Optional startPos As Long = 0, Optional skipnormalization As Boolean = False) As String
Dim pos As Long
Dim posEnd As Long

    ' Returns the text after the clause keyword (replaceString is not used yet).
    If skipnormalization = False Then
        sql = NormalizeSQL(sql)
    End If
    If startPos = 0 Then
        pos = GetSQLClausePos(sql, clause, skipnormalization:=True)
    Else
        pos = startPos
    End If
    If pos = 0 Then
        Exit Function
    End If
    posEnd = InStr(pos + Len(clause) - 1, sql, " ")
    If posEnd = 0 Then
        Exit Function
    End If
    ReplaceSQLClause = Trim(Mid(sql, posEnd))
End Function
Public Function NormalizeSQL(ByVal SQL As String) As String

    SQL = Replace(SQL, vbCrLf, " ")
    SQL = Replace(SQL, vbCr, " ")
    SQL = Replace(SQL, vbLf, " ")
    SQL = Replace(SQL, vbTab, " ")
    SQL = Replace(SQL, ";", " ")
    ' repeat until stable: one pass turns three spaces into two, so the result must not depend
    ' on how many times this function is applied
    Do While InStr(SQL, "  ") > 0
        SQL = Replace(SQL, "  ", " ")
    Loop
    NormalizeSQL = Trim(SQL)
End Function

Public Function WhereOperatorForString(str As String) As String

    WhereOperatorForString = IIf(InStr(1, str, "*") = 0 And InStr(1, str, "?") = 0, " = ", " LIKE ")
End Function

Public Function ReplaceDatabaseInUseClause(ByVal SQL As String, sourceDatabase As String, destinationDatabase As String) As String
Dim startPos As Long
Dim endPos As Long
Dim checkPattern As String
Dim i As Long

    checkPattern = " " & vbCrLf & "/;"
    startPos = InStr(startPos + 1, SQL, "USE ")
    endPos = startPos
    While startPos > 0
        ' check character before USE is blank, vbcr,vbcrlf,vblf, /
        If startPos > 1 Then
            If Not omStringFunctions.IsStringInPattern(Mid(SQL, startPos - 1, 1), checkPattern) Then
                startPos = 0
            End If
        End If
        If startPos <> 0 Then
            For i = startPos + 4 To Len(SQL)
                endPos = i
                If omStringFunctions.IsStringInPattern(Mid(SQL, i, 1), checkPattern) Then
                    endPos = endPos - 1
                    Exit For
                End If
            Next i
            If Mid(SQL, startPos + 4, endPos + 1 - startPos - 4) = sourceDatabase Then
                SQL = Left(SQL, startPos - 1) & "USE " & destinationDatabase & Mid(SQL, endPos + 1)
            End If
        End If
        startPos = endPos
        startPos = InStr(startPos + 1, SQL, "USE ")
        endPos = startPos
    Wend
    ReplaceDatabaseInUseClause = SQL
End Function
