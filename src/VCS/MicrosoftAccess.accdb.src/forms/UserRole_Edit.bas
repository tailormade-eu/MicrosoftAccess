Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =7276
    DatasheetFontHeight =11
    ItemSuffix =101
    Right =18060
    Bottom =11925
    RecSrcDt = Begin
        0xc88f95d61d36e540
    End
    RecordSource ="UserRoles"
    BeforeInsert ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
            Height =3075
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1620
                    Top =600
                    Width =5610
                    Height =315
                    TabIndex =1
                    Name ="Name"
                    ControlSource ="Name"
                    StatusBarText ="Gemeente"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1620
                    LayoutCachedTop =600
                    LayoutCachedWidth =7230
                    LayoutCachedHeight =915
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
                            Left =120
                            Top =600
                            Width =1440
                            Height =315
                            Name ="Label51"
                            Caption ="Name"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =600
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =915
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
                    Left =1620
                    Top =1590
                    Width =5610
                    Height =315
                    TabIndex =3
                    Name ="Sort"
                    ControlSource ="Sort"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1620
                    LayoutCachedTop =1590
                    LayoutCachedWidth =7230
                    LayoutCachedHeight =1905
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
                            Left =120
                            Top =1590
                            Width =1440
                            Height =315
                            Name ="Label65"
                            Caption ="Sort"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =1590
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =1905
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1620
                    Top =105
                    Width =5610
                    Height =315
                    Name ="Code"
                    ControlSource ="Code"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1620
                    LayoutCachedTop =105
                    LayoutCachedWidth =7230
                    LayoutCachedHeight =420
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =105
                            Width =1440
                            Height =315
                            Name ="Label72"
                            Caption ="Code"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =105
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =420
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1620
                    Top =1095
                    Width =5610
                    Height =315
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Description"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1620
                    LayoutCachedTop =1095
                    LayoutCachedWidth =7230
                    LayoutCachedHeight =1410
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
                            Left =120
                            Top =1095
                            Width =1440
                            Height =315
                            Name ="Label79"
                            Caption ="Description"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =1095
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =1410
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1620
                    Top =2085
                    Width =5610
                    Height =315
                    TabIndex =4
                    Name ="Active"
                    ControlSource ="Active"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1620
                    LayoutCachedTop =2085
                    LayoutCachedWidth =7230
                    LayoutCachedHeight =2400
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
                            Left =120
                            Top =2085
                            Width =1440
                            Height =315
                            Name ="Label86"
                            Caption ="Active"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =2085
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =2400
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
                    Left =1620
                    Top =2580
                    Width =5610
                    Height =315
                    TabIndex =5
                    Name ="Memo"
                    ControlSource ="Memo"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1620
                    LayoutCachedTop =2580
                    LayoutCachedWidth =7230
                    LayoutCachedHeight =2895
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
                            Left =120
                            Top =2580
                            Width =1440
                            Height =315
                            Name ="Label94"
                            Caption ="Memo"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =2580
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =2895
                            RowStart =5
                            RowEnd =5
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
' See "UserRole_Edit.cls"
