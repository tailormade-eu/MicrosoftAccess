Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6994
    DatasheetFontHeight =11
    ItemSuffix =308
    Right =18060
    Bottom =11925
    RecSrcDt = Begin
        0x34d5ceb26d32e540
    End
    RecordSource ="User_List_Search"
    DatasheetFontName ="Calibri"
    OnDblClick ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =5955
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1530
                    Top =1815
                    Width =1695
                    Height =315
                    TabIndex =2
                    Name ="Name"
                    ControlSource ="Name"
                    StatusBarText ="Gemeente"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1530
                    LayoutCachedTop =1815
                    LayoutCachedWidth =3225
                    LayoutCachedHeight =2130
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =1815
                            Width =1440
                            Height =315
                            Name ="Label252"
                            Caption ="Name"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =30
                            LayoutCachedTop =1815
                            LayoutCachedWidth =1470
                            LayoutCachedHeight =2130
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1530
                    Top =2805
                    Width =1695
                    Height =315
                    TabIndex =4
                    Name ="Sort"
                    ControlSource ="Sort"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1530
                    LayoutCachedTop =2805
                    LayoutCachedWidth =3225
                    LayoutCachedHeight =3120
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =2805
                            Width =1440
                            Height =315
                            Name ="Label266"
                            Caption ="Sort"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =30
                            LayoutCachedTop =2805
                            LayoutCachedWidth =1470
                            LayoutCachedHeight =3120
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1530
                    Top =1320
                    Width =1695
                    Height =315
                    TabIndex =1
                    Name ="Code"
                    ControlSource ="Code"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1530
                    LayoutCachedTop =1320
                    LayoutCachedWidth =3225
                    LayoutCachedHeight =1635
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =1320
                            Width =1440
                            Height =315
                            Name ="Label273"
                            Caption ="Code"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =30
                            LayoutCachedTop =1320
                            LayoutCachedWidth =1470
                            LayoutCachedHeight =1635
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1530
                    Top =2310
                    Width =1695
                    Height =315
                    TabIndex =3
                    Name ="Description"
                    ControlSource ="Description"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1530
                    LayoutCachedTop =2310
                    LayoutCachedWidth =3225
                    LayoutCachedHeight =2625
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =2310
                            Width =1440
                            Height =315
                            Name ="Label280"
                            Caption ="Description"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =30
                            LayoutCachedTop =2310
                            LayoutCachedWidth =1470
                            LayoutCachedHeight =2625
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1530
                    Top =3300
                    Width =1695
                    Height =315
                    TabIndex =5
                    Name ="Active"
                    ControlSource ="Active"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1530
                    LayoutCachedTop =3300
                    LayoutCachedWidth =3225
                    LayoutCachedHeight =3615
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =3300
                            Width =1440
                            Height =315
                            Name ="Label287"
                            Caption ="Active"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =30
                            LayoutCachedTop =3300
                            LayoutCachedWidth =1470
                            LayoutCachedHeight =3615
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1530
                    Top =3795
                    Width =1695
                    Height =315
                    TabIndex =6
                    Name ="Memo"
                    ControlSource ="Memo"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1530
                    LayoutCachedTop =3795
                    LayoutCachedWidth =3225
                    LayoutCachedHeight =4110
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =3795
                            Width =1440
                            Height =315
                            Name ="Label294"
                            Caption ="Memo"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =30
                            LayoutCachedTop =3795
                            LayoutCachedWidth =1470
                            LayoutCachedHeight =4110
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1530
                    Top =840
                    Width =1695
                    Height =293
                    Name ="UserRoleName"
                    ControlSource ="UserRoleName"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1530
                    LayoutCachedTop =840
                    LayoutCachedWidth =3225
                    LayoutCachedHeight =1133
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =840
                            Width =1440
                            Height =293
                            Name ="Label301"
                            Caption ="Role"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =30
                            LayoutCachedTop =840
                            LayoutCachedWidth =1470
                            LayoutCachedHeight =1133
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "User_List_Search.cls"
