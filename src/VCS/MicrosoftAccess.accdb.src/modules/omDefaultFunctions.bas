Attribute VB_Name = "omDefaultFunctions"
Option Compare Database
Option Explicit

Public gSystemDefaults As New omDefaults
Public gDefaults As New omDefaults

Public Function GetDefault(Name As String) As Variant
    omDefaultFunctions.Initialize
    GetDefault = gDefaults.Load(Name)
End Function
Public Function SaveDefault(Name As String, Value As Variant)
    omDefaultFunctions.Initialize
    gDefaults.Save Name, Value
End Function
Public Function SaveSystemDefault(Name As String, Value As Variant)
    omDefaultFunctions.Initialize
    gSystemDefaults.Save Name, Value
End Function

Public Function GetSystemDefault(Name As String) As Variant
    omDefaultFunctions.Initialize
    GetSystemDefault = gSystemDefaults.Load(Name)
End Function

Public Sub Initialize(Optional Reset As Boolean = False)
    If omObjectFunctions.IsNothing(gSystemDefaults) Or Reset Or Not gSystemDefaults.Initialized Then
        gSystemDefaults.Mode = LocalMode
    End If
    If omObjectFunctions.IsNothing(gDefaults) Or Reset Or Not gDefaults.Initialized Then
        gDefaults.Mode = serverMode
    End If
    gSystemDefaults.Development = gDevelopmentMode
    gDefaults.Development = gDevelopmentMode
End Sub

Public Sub CreateDefaultTables()
    'omTableFunctions.CreateTable "omSysDefaults", "Id", dbLong, True
    'omTableFunctions.AddField "omSysDefaults", "Name", dbText
    'omTableFunctions.AddField "omSysDefaults", "Value", dbMemo
    'omTableFunctions.AddField "omSysDefaults", "ModifyDate", dbDate

    'omTableFunctions.CreateTable "omDefaults", "Id", dbLong, True
    'omTableFunctions.AddField "omDefaults", "Name", dbText
    'omTableFunctions.AddField "omDefaults", "Value", dbMemo
    'omTableFunctions.AddField "omDefaults", "ModifyDate", dbDate
End Sub
