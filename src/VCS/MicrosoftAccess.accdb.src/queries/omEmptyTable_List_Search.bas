Operation =1
Option =0
Where ="(((omEmptyTables.Code) Like \"*\" & Forms!omEmptyTable_List!txtSearch & \"*\")) "
    "Or (((omEmptyTables.Name) Like \"*\" & Forms!omEmptyTable_List!txtSearch & \"*\""
    ")) Or (((omEmptyTables.Description) Like \"*\" & Forms!omEmptyTable_List!txtSear"
    "ch & \"*\")) Or (((omEmptyTables.Memo) Like \"*\" & Forms!omEmptyTable_List!txtS"
    "earch & \"*\"))"
Begin InputTables
    Name ="omEmptyTables"
End
Begin OutputColumns
    Expression ="omEmptyTables.Id"
    Expression ="omEmptyTables.Code"
    Expression ="omEmptyTables.Name"
    Expression ="omEmptyTables.Description"
    Expression ="omEmptyTables.Sort"
    Expression ="omEmptyTables.Active"
    Expression ="omEmptyTables.Memo"
End
Begin OrderBy
    Expression ="omEmptyTables.Name"
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
        dbText "Name" ="IdentityGroups.Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IdentityGroups.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IdentityGroups.Sort"
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
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omEmptyTables.Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omEmptyTables.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omEmptyTables.Memo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omEmptyTables.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omEmptyTables.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omEmptyTables.Sort"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omEmptyTables.Description"
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
    Bottom =385
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =386
        Top =0
        Name ="omEmptyTables"
        Name =""
    End
End
