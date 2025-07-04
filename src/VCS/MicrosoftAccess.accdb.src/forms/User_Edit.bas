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
    Width =7483
    DatasheetFontHeight =11
    ItemSuffix =141
    Right =15023
    Bottom =11865
    Filter ="Id=5"
    RecSrcDt = Begin
        0xc37be6ad6d32e540
    End
    RecordSource ="Users"
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
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Tab
            Width =5103
            Height =3402
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
            Width =1701
            Height =1701
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =5669
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1628
                    Top =1088
                    Width =5610
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Name"
                    ControlSource ="Name"
                    StatusBarText ="Gemeente"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =1628
                    LayoutCachedTop =1088
                    LayoutCachedWidth =7238
                    LayoutCachedHeight =1403
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
                            Top =1088
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label51"
                            Caption ="Name"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =1088
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =1403
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
                    Left =1628
                    Top =4043
                    Width =5610
                    Height =315
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Sort"
                    ControlSource ="Sort"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =1628
                    LayoutCachedTop =4043
                    LayoutCachedWidth =7238
                    LayoutCachedHeight =4358
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =4043
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label65"
                            Caption ="Sort"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =4043
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =4358
                            RowStart =8
                            RowEnd =8
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1628
                    Top =585
                    Width =5610
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Code"
                    ControlSource ="Code"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =1628
                    LayoutCachedTop =585
                    LayoutCachedWidth =7238
                    LayoutCachedHeight =900
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
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label72"
                            Caption ="Code"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =585
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =900
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
                    Left =1628
                    Top =3540
                    Width =5610
                    Height =315
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Description"
                    ControlSource ="Description"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =1628
                    LayoutCachedTop =3540
                    LayoutCachedWidth =7238
                    LayoutCachedHeight =3855
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =3540
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label79"
                            Caption ="Description"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =3540
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =3855
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1628
                    Top =4538
                    Width =5610
                    Height =315
                    TabIndex =9
                    BorderColor =10921638
                    Name ="Active"
                    ControlSource ="Active"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =1628
                    LayoutCachedTop =4538
                    LayoutCachedWidth =7238
                    LayoutCachedHeight =4853
                    RowStart =9
                    RowEnd =9
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =4538
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label86"
                            Caption ="Active"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =4538
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =4853
                            RowStart =9
                            RowEnd =9
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1628
                    Top =5040
                    Width =5610
                    Height =315
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Memo"
                    ControlSource ="Memo"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =1628
                    LayoutCachedTop =5040
                    LayoutCachedWidth =7238
                    LayoutCachedHeight =5355
                    RowStart =10
                    RowEnd =10
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =5040
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label94"
                            Caption ="Memo"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =5040
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =5355
                            RowStart =10
                            RowEnd =10
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1628
                    Top =105
                    Width =5610
                    Height =293
                    BorderColor =10921638
                    Name ="UserRoleId"
                    ControlSource ="UserRoleId"
                    RowSourceType ="Table/Query"
                    RowSource ="UserRole_Select"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="3"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =1628
                    LayoutCachedTop =105
                    LayoutCachedWidth =7238
                    LayoutCachedHeight =398
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =120
                            Top =105
                            Width =1440
                            Height =293
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label101"
                            Caption ="Role"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =105
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =398
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1628
                    Top =2093
                    Width =5610
                    Height =300
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Password"
                    ControlSource ="Password"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =1628
                    LayoutCachedTop =2093
                    LayoutCachedWidth =7238
                    LayoutCachedHeight =2393
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
                            Top =2093
                            Width =1440
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label108"
                            Caption ="Password"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =2093
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =2393
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
                    Left =1628
                    Top =1590
                    Width =5610
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Login"
                    ControlSource ="Login"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =1628
                    LayoutCachedTop =1590
                    LayoutCachedWidth =7238
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
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift115"
                            Caption ="Login"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
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
                    Left =1628
                    Top =2580
                    Width =5610
                    Height =293
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Email"
                    ControlSource ="Email"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =1628
                    LayoutCachedTop =2580
                    LayoutCachedWidth =7238
                    LayoutCachedHeight =2873
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
                            Height =293
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label122"
                            Caption ="Email"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =2580
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =2873
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1628
                    Top =3060
                    Width =5610
                    Height =293
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="WorkHoursPerDay"
                    ControlSource ="WorkHoursPerDay"
                    Format ="Fixed"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="8"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =1628
                    LayoutCachedTop =3060
                    LayoutCachedWidth =7238
                    LayoutCachedHeight =3353
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
                            Left =120
                            Top =3060
                            Width =1440
                            Height =293
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label129"
                            Caption ="HoursPerDay"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =120
                            LayoutCachedTop =3060
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =3353
                            RowStart =6
                            RowEnd =6
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
' See "User_Edit.cls"
