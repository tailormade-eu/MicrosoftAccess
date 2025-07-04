Operation =3
Name ="omSourceObjectControlTranslations"
Option =0
Where ="(((omSourceObjectControlTranslations.Id) Is Null))"
Begin InputTables
    Name ="omSourceObjectControlTranslations"
    Name ="omLanguagesSourceObjectControls"
End
Begin OutputColumns
    Name ="LanguageId"
    Expression ="omLanguagesSourceObjectControls.LanguageId"
    Name ="SourceObjectControlId"
    Expression ="omLanguagesSourceObjectControls.SourceObjectControlId"
    Name ="Default"
    Expression ="omLanguagesSourceObjectControls.ControlDefault"
    Name ="Short"
    Expression ="omLanguagesSourceObjectControls.ControlDefault"
    Name ="Long"
    Expression ="omLanguagesSourceObjectControls.ControlDefault"
    Alias ="CreateDate"
    Name ="CreateDate"
    Expression ="Now()"
    Alias ="LastUsedDate"
    Name ="LastUsedDate"
    Expression ="Now()"
End
Begin Joins
    LeftTable ="omLanguagesSourceObjectControls"
    RightTable ="omSourceObjectControlTranslations"
    Expression ="omLanguagesSourceObjectControls.LanguageId = omSourceObjectControlTranslations.L"
        "anguageId"
    Flag =2
    LeftTable ="omLanguagesSourceObjectControls"
    RightTable ="omSourceObjectControlTranslations"
    Expression ="omLanguagesSourceObjectControls.SourceObjectControlId = omSourceObjectControlTra"
        "nslations.SourceObjectControlId"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="omSourceObjectControlTranslations.Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omLanguagesSourceObjectControls.ControlDefault"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omLanguagesSourceObjectControls.LanguageId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="omLanguagesSourceObjectControls.SourceObjectControlId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CreateDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastUsedDate"
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
    Bottom =1081
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =1127
        Top =147
        Right =1415
        Bottom =662
        Top =0
        Name ="omSourceObjectControlTranslations"
        Name =""
    End
    Begin
        Left =376
        Top =136
        Right =664
        Bottom =712
        Top =0
        Name ="omLanguagesSourceObjectControls"
        Name =""
    End
End
