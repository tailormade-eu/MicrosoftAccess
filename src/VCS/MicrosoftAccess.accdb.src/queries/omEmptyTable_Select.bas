Operation =1
Option =0
Begin InputTables
    Name ="omEmptyTables"
End
Begin OutputColumns
    Expression ="omEmptyTables.Id"
    Expression ="omEmptyTables.Name"
End
Begin OrderBy
    Expression ="omEmptyTables.Sort"
    Flag =0
    Expression ="omEmptyTables.Name"
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
        dbText "Name" ="IdentityGroups.Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IdentityGroups.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omEmptyTables.Id"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =2120
    Bottom =1624
    Left =-1
    Top =-1
    Right =2087
    Bottom =504
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="omEmptyTables"
        Name =""
    End
End
