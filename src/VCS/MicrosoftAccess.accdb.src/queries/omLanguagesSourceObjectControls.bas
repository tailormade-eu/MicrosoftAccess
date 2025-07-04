Operation =1
Option =0
Begin InputTables
    Name ="omLanguages"
    Name ="omSourceObjectControls"
End
Begin OutputColumns
    Alias ="LanguageId"
    Expression ="omLanguages.Id"
    Alias ="SourceObjectControlId"
    Expression ="omSourceObjectControls.Id"
    Expression ="omSourceObjectControls.ControlDefault"
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
        dbText "Name" ="omLanguages.Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LanguageId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omSourceObjectControls.Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SourceObjectControlId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omSourceObjectControls.ControlDefault"
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
    Bottom =1115
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =295
        Top =117
        Right =583
        Bottom =405
        Top =0
        Name ="omLanguages"
        Name =""
    End
    Begin
        Left =839
        Top =108
        Right =1127
        Bottom =627
        Top =0
        Name ="omSourceObjectControls"
        Name =""
    End
End
