Operation =1
Option =0
Begin InputTables
    Name ="UserRoles"
End
Begin OutputColumns
    Expression ="UserRoles.Id"
    Expression ="UserRoles.Name"
End
Begin OrderBy
    Expression ="UserRoles.Sort"
    Flag =0
    Expression ="UserRoles.Name"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UserRoles.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UserRoles.Sort"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UserRoles.Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Users.Name"
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
    Bottom =402
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
