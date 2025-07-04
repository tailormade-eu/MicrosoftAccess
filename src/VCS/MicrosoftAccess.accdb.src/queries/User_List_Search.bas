Operation =1
Option =0
Where ="(((Users.Code) Like \"*\" & Forms!User_List!txtSearch & \"*\")) Or (((Users.Name"
    ") Like \"*\" & Forms!User_List!txtSearch & \"*\")) Or (((Users.Description) Like"
    " \"*\" & Forms!User_List!txtSearch & \"*\")) Or (((Users.Memo) Like \"*\" & Form"
    "s!User_List!txtSearch & \"*\"))"
Begin InputTables
    Name ="Users"
    Name ="UserRoles"
End
Begin OutputColumns
    Expression ="Users.Id"
    Alias ="UserRoleName"
    Expression ="UserRoles.Name"
    Expression ="Users.Code"
    Expression ="Users.Name"
    Expression ="Users.Description"
    Expression ="Users.Sort"
    Expression ="Users.Active"
    Expression ="Users.Memo"
End
Begin Joins
    LeftTable ="Users"
    RightTable ="UserRoles"
    Expression ="Users.UserRoleId = UserRoles.Id"
    Flag =2
End
Begin OrderBy
    Expression ="Users.Name"
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
End
Begin
    State =0
    Left =0
    Top =0
    Right =2442
    Bottom =1624
    Left =-1
    Top =-1
    Right =2409
    Bottom =317
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =96
        Top =24
        Right =384
        Bottom =312
        Top =0
        Name ="Users"
        Name =""
    End
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
