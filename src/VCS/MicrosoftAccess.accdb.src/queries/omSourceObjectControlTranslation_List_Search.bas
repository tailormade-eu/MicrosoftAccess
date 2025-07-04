Operation =1
Option =0
Where ="(((omLanguages.Name) Like \"*\" & Forms!omSourceObjectControlTranslation_List!tx"
    "tSearch & \"*\")) Or (((omSourceObjects.Name) Like \"*\" & Forms!omSourceObjectC"
    "ontrolTranslation_List!txtSearch & \"*\")) Or (((omSourceObjectControls.ControlN"
    "ame) Like \"*\" & Forms!omSourceObjectControlTranslation_List!txtSearch & \"*\")"
    ") Or (((omSourceObjectControls.ControlDefault) Like \"*\" & Forms!omSourceObject"
    "ControlTranslation_List!txtSearch & \"*\")) Or (((omSourceObjectControlTranslati"
    "ons.Default) Like \"*\" & Forms!omSourceObjectControlTranslation_List!txtSearch "
    "& \"*\")) Or (((omSourceObjectControlTranslations.Short) Like \"*\" & Forms!omSo"
    "urceObjectControlTranslation_List!txtSearch & \"*\")) Or (((omSourceObjectContro"
    "lTranslations.Long) Like \"*\" & Forms!omSourceObjectControlTranslation_List!txt"
    "Search & \"*\"))"
Begin InputTables
    Name ="omSourceObjectControlTranslations"
    Name ="omSourceObjectControls"
    Name ="omSourceObjects"
    Name ="omLanguages"
End
Begin OutputColumns
    Expression ="omSourceObjectControlTranslations.Id"
    Expression ="omSourceObjectControls.SourceObjectId"
    Alias ="SourceObjectName"
    Expression ="omSourceObjects.Name"
    Expression ="omSourceObjectControls.ControlName"
    Expression ="omSourceObjectControlTranslations.LanguageId"
    Alias ="LanguageName"
    Expression ="omLanguages.Name"
    Expression ="omSourceObjectControls.ControlDefault"
    Expression ="omSourceObjectControlTranslations.Default"
    Expression ="omSourceObjectControlTranslations.Short"
    Expression ="omSourceObjectControlTranslations.Long"
End
Begin Joins
    LeftTable ="omSourceObjectControlTranslations"
    RightTable ="omSourceObjectControls"
    Expression ="omSourceObjectControlTranslations.SourceObjectControlId = omSourceObjectControls"
        ".Id"
    Flag =1
    LeftTable ="omSourceObjectControls"
    RightTable ="omSourceObjects"
    Expression ="omSourceObjectControls.SourceObjectId = omSourceObjects.Id"
    Flag =1
    LeftTable ="omSourceObjectControlTranslations"
    RightTable ="omLanguages"
    Expression ="omSourceObjectControlTranslations.LanguageId = omLanguages.Id"
    Flag =1
End
Begin OrderBy
    Expression ="omSourceObjects.Name"
    Flag =0
    Expression ="omSourceObjectControls.ControlName"
    Flag =0
    Expression ="omLanguages.Name"
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
        dbText "Name" ="omSourceObjectControlTranslations.Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LanguageName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SourceObjectName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omSourceObjectControls.ControlName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omSourceObjectControls.ControlDefault"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omSourceObjectControlTranslations.LanguageId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omSourceObjectControls.SourceObjectId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omSourceObjectControlTranslations.Default"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omSourceObjectControlTranslations.Short"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omSourceObjectControlTranslations.Long"
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
    Bottom =705
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =96
        Top =24
        Right =795
        Bottom =535
        Top =0
        Name ="omSourceObjectControlTranslations"
        Name =""
    End
    Begin
        Left =930
        Top =13
        Right =1218
        Bottom =301
        Top =0
        Name ="omSourceObjectControls"
        Name =""
    End
    Begin
        Left =1515
        Top =32
        Right =1803
        Bottom =320
        Top =0
        Name ="omSourceObjects"
        Name =""
    End
    Begin
        Left =1162
        Top =318
        Right =1450
        Bottom =606
        Top =0
        Name ="omLanguages"
        Name =""
    End
End
