Operation =1
Option =0
Where ="(((UserRoles.Code) Like \"*\" & Forms!UserRole_List!txtSearch & \"*\")) Or (((Us"
    "erRoles.Name) Like \"*\" & Forms!UserRole_List!txtSearch & \"*\")) Or (((UserRol"
    "es.Description) Like \"*\" & Forms!UserRole_List!txtSearch & \"*\")) Or (((UserR"
    "oles.Memo) Like \"*\" & Forms!UserRole_List!txtSearch & \"*\"))"
Begin InputTables
    Name ="UserRoles"
End
Begin OutputColumns
    Expression ="UserRoles.Id"
    Expression ="UserRoles.Code"
    Expression ="UserRoles.Name"
    Expression ="UserRoles.Description"
    Expression ="UserRoles.Sort"
    Expression ="UserRoles.Active"
    Expression ="UserRoles.Memo"
End
Begin OrderBy
    Expression ="UserRoles.Name"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Users.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Users.Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Users.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Users.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Users.Sort"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Users.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Users.Memo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UserRoles.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UserRoleName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UserRoles.Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UserRoles.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UserRoles.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UserRoles.Sort"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UserRoles.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UserRoles.Memo"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =3403
    Bottom =1624
    Left =-1
    Top =-1
    Right =3370
    Bottom =249
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =480
        Top =24
        Right =768
        Bottom =312
        Top =0
        Name ="UserRoles"
        Name =""
    End
End
