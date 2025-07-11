﻿Attribute VB_Name = "omSQLFunctions"
Option Compare Database
Option Explicit

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


Public Function GetSelect(SQL As String) As String
Dim posFrom As Long

    posFrom = GetSQLClausePos(SQL, "FROM")
    GetSelect = ReplaceSQLClause(Mid(SQL, 1, posFrom - 1), "SELECT", "")

End Function
Public Function GetFrom(SQL As String) As String
Dim posFrom As Long
Dim poswhere As Long
Dim posOrderBy As Long
Dim posGroupBy As Long
Dim posUse As Long

    posFrom = GetSQLClausePos(SQL, "FROM")
    poswhere = GetSQLClausePos(SQL, "WHERE")
    posOrderBy = GetSQLClausePos(SQL, "ORDER BY", poswhere + 1)
    posGroupBy = GetSQLClausePos(SQL, "GROUP BY", poswhere + 1)
    If poswhere <> 0 Or posOrderBy <> 0 Or posGroupBy <> 0 Then
        If poswhere <> 0 Then
            posUse = poswhere
        End If
        If posOrderBy <> 0 Then
            posUse = posOrderBy
        End If
        If posGroupBy <> 0 Then
            posUse = posGroupBy
        End If
    End If
    If posUse = 0 Then
        posUse = Len(SQL) + 1
    End If
    GetFrom = ReplaceSQLClause(Mid(SQL, posFrom, posUse - (posFrom - 1)), "FROM", "")


End Function

Public Function GetWhere(SQL As String) As String
Dim poswhere As Long
Dim posOrderBy As Long
Dim posGroupBy As Long
Dim posUse As Long

    poswhere = GetSQLClausePos(SQL, "WHERE")
    If poswhere = 0 Then
        Exit Function
    End If
    posOrderBy = GetSQLClausePos(SQL, "ORDER BY", poswhere)
    posGroupBy = GetSQLClausePos(SQL, "GROUP BY", poswhere)
    If posOrderBy <> 0 Or posGroupBy <> 0 Then
        If posOrderBy <> 0 Then
            posUse = posOrderBy
        End If
        If posUse <> 0 And posGroupBy <> 0 Then
            posUse = posGroupBy
        End If
    End If
    If posUse = 0 Then
        posUse = Len(SQL) + 1
    End If
    GetWhere = ReplaceSQLClause(Mid(SQL, poswhere, posUse - (poswhere - 1)), "WHERE", "")


End Function

Public Function GetOrderBy(ByVal SQL As String) As String
Dim posOrderBy As Long


'SELECT Course.Course_ID, Course.Course_Year
'FROM Course
'where
'GROUP BY Course.Course_ID, Course.Course_Year
'HAVING (((Course.Course_Year) = [1]))
'ORDER BY Course.Course_Year;
    SQL = NormalizeSQL(SQL)
    posOrderBy = GetSQLClausePos(SQL, "ORDER BY")
    If posOrderBy <> 0 Then
        GetOrderBy = ReplaceSQLClause(Mid(SQL, posOrderBy + 1, Len(SQL) - (posOrderBy - 1)), "ORDER BY", "")
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

Public Function GetSQLClausePos(ByVal SQL As String, clause As String, Optional startPos As Long = 1) As Long
Dim pos As Long

    SQL = NormalizeSQL(SQL)
    pos = InStr(startPos, SQL, " " & clause & " ")
    If pos = 0 Then
        pos = InStr(startPos, SQL, clause & " ")
    End If

    GetSQLClausePos = pos
End Function

Public Function ReplaceSQLClause(ByVal SQL As String, clause As String, Optional replaceString As String = "", Optional startPos As Long = 0) As String
Dim pos As Long
Dim posEnd As Long
Dim result As String

    SQL = NormalizeSQL(SQL)
    If startPos = 0 Then
        pos = GetSQLClausePos(SQL, clause)
    Else
        pos = startPos
    End If
    If pos <> 0 Then
        posEnd = InStr(pos + Len(clause) - 1, SQL, " ")
        If posEnd = 0 Then
            posEnd = InStr(pos, SQL, vbCrLf)
        End If
        If posEnd = 0 Then
            posEnd = InStr(pos, SQL, vbCr)
        End If
        result = Mid(SQL, posEnd)
    End If
    ReplaceSQLClause = Trim(result)
End Function
Public Function NormalizeSQL(ByVal SQL As String) As String

    SQL = Replace(SQL, vbCrLf, " ")
    SQL = Replace(SQL, vbCr, " ")
    SQL = Replace(SQL, ";", " ")
    SQL = Replace(SQL, "  ", " ")
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
