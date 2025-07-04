Operation =1
Option =0
Begin InputTables
    Name ="Users"
End
Begin OutputColumns
    Expression ="Users.Id"
    Expression ="Users.Name"
End
Begin OrderBy
    Expression ="Users.Sort"
    Flag =0
    Expression ="Users.Name"
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
        dbText "Name" ="Users.Id"
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
    Right =2442
    Bottom =1624
    Left =-1
    Top =-1
    Right =2409
    Bottom =368
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
End
