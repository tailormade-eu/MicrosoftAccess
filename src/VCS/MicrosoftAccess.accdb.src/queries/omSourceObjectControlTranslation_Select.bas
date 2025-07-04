Operation =1
Option =0
Begin InputTables
    Name ="omSourceObjectControlTranslations"
End
Begin OutputColumns
    Expression ="omSourceObjectControlTranslations.Id"
    Expression ="omSourceObjectControlTranslations.Name"
End
Begin OrderBy
    Expression ="omSourceObjectControlTranslations.Sort"
    Flag =0
    Expression ="omSourceObjectControlTranslations.Name"
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
    Right =3081
    Bottom =1624
    Left =-1
    Top =-1
    Right =3048
    Bottom =470
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
