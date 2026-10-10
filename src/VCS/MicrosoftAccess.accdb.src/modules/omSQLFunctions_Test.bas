Attribute VB_Name = "omSQLFunctions_Test"
Option Compare Database
Option Explicit

' Tests for the clause helpers in omSQLFunctions.
' Run from the Immediate window:  ?omSQLFunctions_Test.RunAll
' Returns "OK (n tests)" or one line per failing case.

Private mFailures As String
Private mTests As Long

Private Sub Check(testName As String, actual As String, expected As String)
    mTests = mTests + 1
    If StrComp(actual, expected, vbBinaryCompare) <> 0 Then
        mFailures = mFailures & testName & ": got [" & actual & "] expected [" & expected & "]" & vbCrLf
    End If
End Sub

Public Function RunAll() As String
Dim q1 As String
Dim q2 As String
Dim q3 As String
Dim copy As String

    mFailures = ""
    mTests = 0

    ' multi-line SQL with joins and a trailing semicolon
    q1 = "SELECT a.Id, b.Name" & vbCrLf & "FROM Customers AS a" & vbCrLf & "INNER JOIN Orders AS b" & vbCrLf & _
         "  ON a.Id = b.CustomerId" & vbCrLf & "LEFT JOIN Items AS c ON c.OrderId = b.Id" & vbCrLf & "WHERE a.Active = 1;"
    Check "multiline GetSelect", omSQLFunctions.GetSelect(q1), "a.Id, b.Name"
    Check "multiline GetFrom", omSQLFunctions.GetFrom(q1), "Customers AS a INNER JOIN Orders AS b ON a.Id = b.CustomerId LEFT JOIN Items AS c ON c.OrderId = b.Id"
    Check "multiline GetWhere", omSQLFunctions.GetWhere(q1), "a.Active = 1"

    ' FROM must stop at the first following clause
    Check "GetFrom WHERE+ORDER BY", omSQLFunctions.GetFrom("SELECT a FROM t WHERE x=1 ORDER BY a"), "t"
    Check "GetFrom WHERE+GROUP BY", omSQLFunctions.GetFrom("SELECT a FROM t WHERE x=1 GROUP BY a"), "t"
    Check "GetFrom plain", omSQLFunctions.GetFrom("SELECT a FROM t"), "t"
    Check "GetFrom no FROM", omSQLFunctions.GetFrom("SELECT 1"), ""

    ' WHERE must stop before GROUP BY / ORDER BY
    Check "GetWhere GROUP BY only", omSQLFunctions.GetWhere("SELECT a FROM t WHERE x = 1 GROUP BY a"), "x = 1"
    Check "GetWhere ORDER BY only", omSQLFunctions.GetWhere("SELECT a FROM t WHERE x = 1 ORDER BY a"), "x = 1"
    Check "GetWhere none", omSQLFunctions.GetWhere("SELECT a FROM t"), ""

    ' SELECT without FROM must not raise an error
    Check "GetSelect no FROM", omSQLFunctions.GetSelect("SELECT 1"), "1"

    ' ORDER BY / HAVING
    Check "GetOrderBy DESC", omSQLFunctions.GetOrderBy("SELECT a FROM t ORDER BY a DESC, b"), "a DESC, b"
    Check "GetOrderBy none", omSQLFunctions.GetOrderBy("SELECT a FROM t"), ""
    q2 = "SELECT a, COUNT(*) FROM t GROUP BY a HAVING COUNT(*) > 1 ORDER BY a"
    Check "HAVING GetFrom", omSQLFunctions.GetFrom(q2), "t"
    Check "HAVING GetWhere", omSQLFunctions.GetWhere(q2), ""
    Check "HAVING GetOrderBy", omSQLFunctions.GetOrderBy(q2), "a"

    ' NormalizeSQL must give the same result when applied twice
    q3 = "A   B    C"
    Check "Normalize 3 spaces", omSQLFunctions.NormalizeSQL(q3), "A B C"
    Check "Normalize idempotent", omSQLFunctions.NormalizeSQL(omSQLFunctions.NormalizeSQL(q3)), omSQLFunctions.NormalizeSQL(q3)

    ' skipnormalization = True on already normalized SQL gives the same result
    Check "skip = True", omSQLFunctions.GetFrom(omSQLFunctions.NormalizeSQL(q1), True), omSQLFunctions.GetFrom(q1)

    ' the caller's variable must not be changed
    copy = q1
    omSQLFunctions.GetSelect copy
    Check "ByVal", copy, q1

    If Len(mFailures) = 0 Then
        RunAll = "OK (" & mTests & " tests)"
    Else
        RunAll = mFailures
    End If
End Function
