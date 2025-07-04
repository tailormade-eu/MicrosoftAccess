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
    Width =7370
    DatasheetFontHeight =11
    ItemSuffix =148
    Left =10613
    Top =4493
    Right =17985
    Bottom =5970
    RecSrcDt = Begin
        0xc37be6ad6d32e540
    End
    RecordSource ="Users"
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
            Height =1500
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1635
                    Top =105
                    Width =5610
                    Height =293
                    Name ="txtPassword"
                    InputMask ="Password"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1635
                    LayoutCachedTop =105
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =398
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
                            Width =1448
                            Height =293
                            Name ="Label108"
                            Caption ="Password"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =105
                            LayoutCachedWidth =1568
                            LayoutCachedHeight =398
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1635
                    Top =1065
                    Width =5610
                    Height =285
                    ForeColor =255
                    Name ="lblStatus"
                    Caption =" "
                    GroupTable =1
                    BottomPadding =150
                    LayoutCachedLeft =1635
                    LayoutCachedTop =1065
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =1350
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =120
                    Top =1065
                    Width =1448
                    Height =285
                    Name ="EmptyCell138"
                    GroupTable =1
                    BottomPadding =150
                    LayoutCachedLeft =120
                    LayoutCachedTop =1065
                    LayoutCachedWidth =1568
                    LayoutCachedHeight =1350
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1635
                    Top =585
                    Width =5610
                    Height =293
                    TabIndex =1
                    Name ="txtPasswordVerify"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="Password"
                    GroupTable =1
                    BottomPadding =150

                    LayoutCachedLeft =1635
                    LayoutCachedTop =585
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =878
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
                            Top =585
                            Width =1448
                            Height =293
                            Name ="Label141"
                            Caption ="Verify Password"
                            GroupTable =1
                            BottomPadding =150
                            LayoutCachedLeft =120
                            LayoutCachedTop =585
                            LayoutCachedWidth =1568
                            LayoutCachedHeight =878
                            RowStart =1
                            RowEnd =1
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
' See "User_EditPassword.cls"
