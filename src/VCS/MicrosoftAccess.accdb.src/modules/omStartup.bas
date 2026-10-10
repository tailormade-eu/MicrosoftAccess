Attribute VB_Name = "omStartup"
Option Compare Database
Option Explicit
'@lib 2026-10-10 sha=b79bf44a

' Helpers to link the back-end tables. The macro omLinkDatabase runs LinkSSMADatabase.
' A project specific Main (which form to open at startup) does not belong in the library.

Public Function LinkAccessDatabase() As Variant
    omSSMAAConnector.LinkMSAccess False
End Function

Public Function LinkSSMADatabase() As Variant
    omSSMAAConnector.LinkUsingSSMA SavePassword:=True, alwaysUpdate:=True
End Function

Public Function LinkSSMADatabaseSQLODBC() As Variant
    omSSMAAConnector.LinkUsingSSMA ConnectionType:=SQLODBC_Deprecated, SavePassword:=True, alwaysUpdate:=True
End Function
