Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =16416
    DatasheetFontHeight =11
    ItemSuffix =895
    Left =945
    Top =1260
    Right =17610
    Bottom =7305
    RecSrcDt = Begin
        0xc153b49aed8be440
    End
    RecordSource ="SELECT BillingRulesClient.* FROM BillingRulesClient;"
    Caption ="GroupBillingRules_subform"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
        Begin Rectangle
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
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
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
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
            LabelX =-1800
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
            LabelX =-1800
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
        Begin ToggleButton
            TextFontCharSet =177
            TextFontFamily =0
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
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
            PressedForeThemeColorIndex =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =645
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =465
                    Width =16416
                    Height =180
                    Name ="Image142"
                    Picture ="3_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =465
                    LayoutCachedWidth =16416
                    LayoutCachedHeight =645
                End
                Begin Label
                    OverlapFlags =247
                    Left =120
                    Width =7860
                    Height =540
                    FontSize =20
                    Name ="Label126"
                    Caption ="Billing and Fee Rules"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =540
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =75.0
                End
            End
        End
        Begin Section
            Height =5190
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =93
                    Left =7740
                    Top =240
                    Width =8205
                    Height =2736
                    Name ="SplitBox"
                    LayoutCachedLeft =7740
                    LayoutCachedTop =240
                    LayoutCachedWidth =15945
                    LayoutCachedHeight =2976
                    BorderThemeColorIndex =2
                End
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =93
                    Left =255
                    Top =240
                    Width =285
                    Height =2736
                    Name ="Box866"
                    LayoutCachedLeft =255
                    LayoutCachedTop =240
                    LayoutCachedWidth =540
                    LayoutCachedHeight =2976
                    BorderThemeColorIndex =2
                End
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =93
                    Left =6045
                    Top =240
                    Width =285
                    Height =2736
                    Name ="Box865"
                    LayoutCachedLeft =6045
                    LayoutCachedTop =240
                    LayoutCachedWidth =6330
                    LayoutCachedHeight =2976
                    BorderThemeColorIndex =2
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1224
                    Top =3192
                    Width =1620
                    Height =276
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =3
                    Name ="PerDiemType"
                    ControlSource ="PerDiemType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"PerDiem\") AND ((BillingRules.Deleted)=False));"
                    ColumnWidths ="0;1440;0"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set Per diem - Can differ from RENT settings"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =1224
                    LayoutCachedTop =3192
                    LayoutCachedWidth =2844
                    LayoutCachedHeight =3468
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =420
                            Top =3192
                            Width =795
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label147"
                            Caption ="Bill Type"
                            GroupTable =4
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =420
                            LayoutCachedTop =3192
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =3468
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =4
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1224
                    Top =3540
                    Width =1620
                    Height =276
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =4
                    Name ="PerDiemStartOn"
                    ControlSource ="PerDiemStartOn"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"StartOn\") AND ((BillingRules.Deleted)=False));"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    Tag ="RENT"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =1224
                    LayoutCachedTop =3540
                    LayoutCachedWidth =2844
                    LayoutCachedHeight =3816
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =420
                            Top =3540
                            Width =795
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="PerDiemStart_Label"
                            Caption ="Starts On"
                            GroupTable =4
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =420
                            LayoutCachedTop =3540
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =3816
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =4
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =420
                    Top =2832
                    Width =2424
                    Height =288
                    FontWeight =700
                    Name ="PerDiem_Label"
                    Caption ="Per Diem Rules"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =420
                    LayoutCachedTop =2832
                    LayoutCachedWidth =2844
                    LayoutCachedHeight =3120
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1224
                    Top =3888
                    Width =1620
                    Height =276
                    FontSize =10
                    TabIndex =5
                    Name ="PerDiemFeeToUse"
                    ControlSource ="PerDiemFeeToUse"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"Fee\") AND ((BillingRules.Deleted)=False));"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    Tag ="RENT"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =1224
                    LayoutCachedTop =3888
                    LayoutCachedWidth =2844
                    LayoutCachedHeight =4164
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =420
                            Top =3888
                            Width =795
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label161"
                            Caption ="Fee"
                            GroupTable =4
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =420
                            LayoutCachedTop =3888
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =4164
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =4
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3708
                    Top =3192
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =7
                    Name ="PerDiemFeeLabel"
                    ControlSource ="PerDiemFeeLabel"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"Feelabel\") AND ((BillingRules.Deleted)=False)) ORDER B"
                        "Y BillingRules.RuleValue;"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    Tag ="PDfee"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =3708
                    LayoutCachedTop =3192
                    LayoutCachedWidth =5028
                    LayoutCachedHeight =3468
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =2844
                            Top =3192
                            Width =852
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label173"
                            Caption ="Fee Label"
                            GroupTable =4
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2844
                            LayoutCachedTop =3192
                            LayoutCachedWidth =3696
                            LayoutCachedHeight =3468
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =4
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =7068
                    Top =3156
                    Width =1428
                    Height =288
                    FontSize =10
                    TabIndex =11
                    BackColor =15921906
                    ForeColor =8210719
                    Name ="BillType"
                    ControlSource ="BillType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"Bill\") AND ((BillingRules.Deleted)=False)) ORDER BY Bi"
                        "llingRules.RuleValue;"
                    ColumnWidths ="0;2880;0"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set Bill Type- Can differ from Per Diem settings"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =7068
                    LayoutCachedTop =3156
                    LayoutCachedWidth =8496
                    LayoutCachedHeight =3444
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =3
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeShade =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6216
                            Top =3156
                            Width =816
                            Height =288
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label185"
                            Caption ="Bill Type"
                            GroupTable =6
                            LayoutCachedLeft =6216
                            LayoutCachedTop =3156
                            LayoutCachedWidth =7032
                            LayoutCachedHeight =3444
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7068
                    Top =3516
                    Width =1428
                    Height =276
                    FontSize =10
                    TabIndex =12
                    BackColor =15921906
                    ForeColor =8210719
                    Name ="StartRentOn"
                    ControlSource ="StartRentOn"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"StartOn\") AND ((BillingRules.Deleted)=False));"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    Tag ="RENT"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =7068
                    LayoutCachedTop =3516
                    LayoutCachedWidth =8496
                    LayoutCachedHeight =3792
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =3
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeShade =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6216
                            Top =3516
                            Width =816
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label187"
                            Caption ="Rent Start"
                            GroupTable =6
                            LayoutCachedLeft =6216
                            LayoutCachedTop =3516
                            LayoutCachedWidth =7032
                            LayoutCachedHeight =3792
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =6216
                    Top =2832
                    Width =1680
                    Height =288
                    FontWeight =700
                    Name ="Label188"
                    Caption ="Rent Billing Rules"
                    GroupTable =6
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6216
                    LayoutCachedTop =2832
                    LayoutCachedWidth =7896
                    LayoutCachedHeight =3120
                    ColumnEnd =1
                    LayoutGroup =3
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    BorderShade =50.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =6
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11592
                    Top =3156
                    Width =1092
                    Height =288
                    FontSize =10
                    TabIndex =16
                    BackColor =15921906
                    ForeColor =8210719
                    Name ="FeeToUse"
                    ControlSource ="FeeToUse"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"Fee\") AND ((BillingRules.Deleted)=False));"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    Tag ="RENT"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =11592
                    LayoutCachedTop =3156
                    LayoutCachedWidth =12684
                    LayoutCachedHeight =3444
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeShade =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =10824
                            Top =3156
                            Width =756
                            Height =288
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label190"
                            Caption ="Fee"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =10824
                            LayoutCachedTop =3156
                            LayoutCachedWidth =11580
                            LayoutCachedHeight =3444
                            RowStart =1
                            RowEnd =1
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =13536
                    Top =3156
                    Width =1320
                    Height =288
                    FontSize =10
                    TabIndex =20
                    Name ="FeeLabel"
                    ControlSource ="FeeLabel"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"Feelabel\") AND ((BillingRules.Deleted)=False)) ORDER B"
                        "Y BillingRules.RuleValue;"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    Tag ="RentFee"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =13536
                    LayoutCachedTop =3156
                    LayoutCachedWidth =14856
                    LayoutCachedHeight =3444
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =12684
                            Top =3156
                            Width =840
                            Height =288
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label194"
                            Caption ="Fee Label"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =12684
                            LayoutCachedTop =3156
                            LayoutCachedWidth =13524
                            LayoutCachedHeight =3444
                            RowStart =1
                            RowEnd =1
                            ColumnStart =8
                            ColumnEnd =8
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9324
                    Top =3156
                    Width =1500
                    Height =288
                    FontSize =10
                    TabIndex =14
                    BackColor =15921906
                    ForeColor =8210719
                    Name ="RemitTo"
                    ControlSource ="RemitTo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"Remit\") AND ((BillingRules.Deleted)=False));"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =9324
                    LayoutCachedTop =3156
                    LayoutCachedWidth =10824
                    LayoutCachedHeight =3444
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =3
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeShade =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =8496
                            Top =3156
                            Width =816
                            Height =288
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label200"
                            Caption ="Remit To"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =8496
                            LayoutCachedTop =3156
                            LayoutCachedWidth =9312
                            LayoutCachedHeight =3444
                            RowStart =1
                            RowEnd =1
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9324
                    Top =3516
                    Width =1500
                    Height =276
                    FontSize =10
                    TabIndex =24
                    BackColor =15921906
                    Name ="Remittee"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =9324
                    LayoutCachedTop =3516
                    LayoutCachedWidth =10824
                    LayoutCachedHeight =3792
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =3
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =87
                            TextAlign =3
                            Left =8496
                            Top =3516
                            Width =816
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label485"
                            Caption ="Assignee"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =8496
                            LayoutCachedTop =3516
                            LayoutCachedWidth =9312
                            LayoutCachedHeight =3792
                            RowStart =2
                            RowEnd =2
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin EmptyCell
                    Left =11592
                    Top =2832
                    Width =1092
                    Height =288
                    Name ="EmptyCell220"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11592
                    LayoutCachedTop =2832
                    LayoutCachedWidth =12684
                    LayoutCachedHeight =3120
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =10824
                    Top =2832
                    Width =756
                    Height =288
                    Name ="EmptyCell230"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10824
                    LayoutCachedTop =2832
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =3120
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =7896
                    Top =3864
                    Width =600
                    Height =276
                    FontSize =9
                    FontWeight =700
                    Name ="DayofMthLabel"
                    Caption ="th  Day"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =7896
                    LayoutCachedTop =3864
                    LayoutCachedWidth =8496
                    LayoutCachedHeight =4140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =6
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11592
                    Top =3864
                    Width =1092
                    Height =276
                    FontSize =10
                    TabIndex =18
                    BackColor =15921906
                    Name ="FeeTerm"
                    ControlSource ="FeeTerm"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="25"
                    Tag ="RentFee"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =11592
                    LayoutCachedTop =3864
                    LayoutCachedWidth =12684
                    LayoutCachedHeight =4140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =10824
                            Top =3864
                            Width =756
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label274"
                            Caption ="Term"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =10824
                            LayoutCachedTop =3864
                            LayoutCachedWidth =11580
                            LayoutCachedHeight =4140
                            RowStart =3
                            RowEnd =3
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7068
                    Top =3864
                    Width =828
                    Height =276
                    FontSize =10
                    TabIndex =13
                    BackColor =15921906
                    Name ="DayOfMonth"
                    ControlSource ="DayOfMonth"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="25"
                    Tag ="RENT"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =7068
                    LayoutCachedTop =3864
                    LayoutCachedWidth =7896
                    LayoutCachedHeight =4140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6216
                            Top =3864
                            Width =816
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label255"
                            Caption ="Bill On"
                            GroupTable =6
                            LayoutCachedLeft =6216
                            LayoutCachedTop =3864
                            LayoutCachedWidth =7032
                            LayoutCachedHeight =4140
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =6210
                    Top =30
                    Width =5085
                    Height =2850
                    FontSize =10
                    TabIndex =1
                    ForeColor =10040879
                    Name ="BillDescRule"
                    TextFormat =1

                    LayoutCachedLeft =6210
                    LayoutCachedTop =30
                    LayoutCachedWidth =11295
                    LayoutCachedHeight =2880
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9324
                    Top =3864
                    Width =1500
                    Height =276
                    FontSize =10
                    TabIndex =15
                    BackColor =15921906
                    ForeColor =8210719
                    Name ="ShowContract"
                    ControlSource ="ShowContract"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"ShowHide\") AND ((BillingRules.Deleted)=False));"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =9324
                    LayoutCachedTop =3864
                    LayoutCachedWidth =10824
                    LayoutCachedHeight =4140
                    DatasheetCaption ="`"
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =3
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeShade =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8496
                            Top =3864
                            Width =816
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label531"
                            Caption ="Contract"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =8496
                            LayoutCachedTop =3864
                            LayoutCachedWidth =9312
                            LayoutCachedHeight =4140
                            RowStart =3
                            RowEnd =3
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =13536
                    Top =3516
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =21
                    Name ="ShowFeeColumn"
                    ControlSource ="ShowContract"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"ShowHide\") AND ((BillingRules.Deleted)=False)) ORDER B"
                        "Y BillingRules.RuleValue DESC;"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    Tag ="RentFee"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =13536
                    LayoutCachedTop =3516
                    LayoutCachedWidth =14856
                    LayoutCachedHeight =3792
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =12684
                            Top =3516
                            Width =840
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label538"
                            Caption ="Show Fee"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =12684
                            LayoutCachedTop =3516
                            LayoutCachedWidth =13524
                            LayoutCachedHeight =3792
                            RowStart =2
                            RowEnd =2
                            ColumnStart =8
                            ColumnEnd =8
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11592
                    Top =4212
                    Width =1092
                    Height =276
                    FontSize =10
                    TabIndex =19
                    BackColor =15921906
                    ForeColor =8210719
                    Name ="SplitFee"
                    ControlSource ="SplitFee"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"Split\") AND ((BillingRules.Deleted)=False)) ORDER BY B"
                        "illingRules.RuleValue;"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    Tag ="RentFee"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =11592
                    LayoutCachedTop =4212
                    LayoutCachedWidth =12684
                    LayoutCachedHeight =4488
                    RowStart =4
                    RowEnd =4
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeShade =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10824
                            Top =4212
                            Width =756
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label544"
                            Caption ="Split Bill"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =10824
                            LayoutCachedTop =4212
                            LayoutCachedWidth =11580
                            LayoutCachedHeight =4488
                            RowStart =4
                            RowEnd =4
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =13536
                    Top =4212
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =23
                    Name ="RefundFee"
                    ControlSource ="RefundFee"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"Refund\") AND ((BillingRules.Deleted)=False));"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    Tag ="RentFee"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =13536
                    LayoutCachedTop =4212
                    LayoutCachedWidth =14856
                    LayoutCachedHeight =4488
                    RowStart =4
                    RowEnd =4
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =12684
                            Top =4212
                            Width =840
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label550"
                            Caption ="Refund"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =12684
                            LayoutCachedTop =4212
                            LayoutCachedWidth =13524
                            LayoutCachedHeight =4488
                            RowStart =4
                            RowEnd =4
                            ColumnStart =8
                            ColumnEnd =8
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3708
                    Top =3540
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =8
                    Name ="PerDiemShowFeeColumn"
                    ControlSource ="PerDiemShowFeeColumn"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"ShowHide\") AND ((BillingRules.Deleted)=False)) ORDER B"
                        "Y BillingRules.RuleValue DESC;"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    Tag ="PDFee"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =3708
                    LayoutCachedTop =3540
                    LayoutCachedWidth =5028
                    LayoutCachedHeight =3816
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =2844
                            Top =3540
                            Width =852
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label556"
                            Caption ="Show Fee"
                            GroupTable =4
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2844
                            LayoutCachedTop =3540
                            LayoutCachedWidth =3696
                            LayoutCachedHeight =3816
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =4
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3708
                    Top =4236
                    Width =1320
                    Height =288
                    FontSize =10
                    TabIndex =10
                    Name ="PerDiemRefundFee"
                    ControlSource ="PerDiemRefundFee"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"YesNo\") AND ((BillingRules.Deleted)=False));"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    Tag ="PDFee"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =3708
                    LayoutCachedTop =4236
                    LayoutCachedWidth =5028
                    LayoutCachedHeight =4524
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =2844
                            Top =4236
                            Width =852
                            Height =288
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label558"
                            Caption ="Refund"
                            GroupTable =4
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2844
                            LayoutCachedTop =4236
                            LayoutCachedWidth =3696
                            LayoutCachedHeight =4524
                            RowStart =4
                            RowEnd =4
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =4
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =405
                    Top =30
                    Width =5100
                    Height =2850
                    FontSize =10
                    ForeColor =10040879
                    Name ="DescPerDiem"
                    TextFormat =1

                    LayoutCachedLeft =405
                    LayoutCachedTop =30
                    LayoutCachedWidth =5505
                    LayoutCachedHeight =2880
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin EmptyCell
                    Left =3708
                    Top =2832
                    Width =1320
                    Height =288
                    Name ="EmptyCell596"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =3708
                    LayoutCachedTop =2832
                    LayoutCachedWidth =5028
                    LayoutCachedHeight =3120
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =2844
                    Top =2832
                    Width =852
                    Height =288
                    Name ="EmptyCell603"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2844
                    LayoutCachedTop =2832
                    LayoutCachedWidth =3696
                    LayoutCachedHeight =3120
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =13536
                    Top =2832
                    Width =1320
                    Height =288
                    Name ="EmptyCell622"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13536
                    LayoutCachedTop =2832
                    LayoutCachedWidth =14856
                    LayoutCachedHeight =3120
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =12684
                    Top =2832
                    Width =840
                    Height =288
                    Name ="EmptyCell636"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12684
                    LayoutCachedTop =2832
                    LayoutCachedWidth =13524
                    LayoutCachedHeight =3120
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =9324
                    Top =2832
                    Width =1500
                    Height =288
                    Name ="EmptyCell649"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9324
                    LayoutCachedTop =2832
                    LayoutCachedWidth =10824
                    LayoutCachedHeight =3120
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =8496
                    Top =2832
                    Width =816
                    Height =288
                    Name ="EmptyCell663"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8496
                    LayoutCachedTop =2832
                    LayoutCachedWidth =9312
                    LayoutCachedHeight =3120
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1224
                    Top =4236
                    Width =1620
                    Height =288
                    FontSize =10
                    TabIndex =6
                    Name ="PerDiemFee"
                    ControlSource ="PerDiemFeeValue"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="65"
                    Tag ="RENT"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1224
                    LayoutCachedTop =4236
                    LayoutCachedWidth =2844
                    LayoutCachedHeight =4524
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    GroupTable =4
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =420
                            Top =4236
                            Width =795
                            Height =288
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label681"
                            Caption ="Amount"
                            GroupTable =4
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =420
                            LayoutCachedTop =4236
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =4524
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =4
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11592
                    Top =3516
                    Width =1092
                    Height =276
                    FontSize =10
                    TabIndex =17
                    BackColor =15921906
                    Name ="InvoiceFee"
                    ControlSource ="FeeValue"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="65"
                    Tag ="RentFee"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =11592
                    LayoutCachedTop =3516
                    LayoutCachedWidth =12684
                    LayoutCachedHeight =3792
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =6
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =10824
                            Top =3516
                            Width =756
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label688"
                            Caption ="Amount"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =10824
                            LayoutCachedTop =3516
                            LayoutCachedWidth =11580
                            LayoutCachedHeight =3792
                            RowStart =2
                            RowEnd =2
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin EmptyCell
                    Left =6216
                    Top =4212
                    Width =816
                    Height =276
                    Name ="EmptyCell731"
                    GroupTable =6
                    LayoutCachedLeft =6216
                    LayoutCachedTop =4212
                    LayoutCachedWidth =7032
                    LayoutCachedHeight =4488
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =7068
                    Top =4212
                    Width =828
                    Height =276
                    Name ="EmptyCell732"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =7068
                    LayoutCachedTop =4212
                    LayoutCachedWidth =7896
                    LayoutCachedHeight =4488
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =7896
                    Top =4212
                    Width =600
                    Height =276
                    Name ="EmptyCell733"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =7896
                    LayoutCachedTop =4212
                    LayoutCachedWidth =8496
                    LayoutCachedHeight =4488
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =8496
                    Top =4212
                    Width =816
                    Height =276
                    Name ="EmptyCell734"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8496
                    LayoutCachedTop =4212
                    LayoutCachedWidth =9312
                    LayoutCachedHeight =4488
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =9324
                    Top =4212
                    Width =1500
                    Height =276
                    Name ="EmptyCell735"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =9324
                    LayoutCachedTop =4212
                    LayoutCachedWidth =10824
                    LayoutCachedHeight =4488
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =10824
                    Top =4212
                    Width =0
                    Height =276
                    Name ="EmptyCell736"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =10824
                    LayoutCachedTop =4212
                    LayoutCachedWidth =10824
                    LayoutCachedHeight =4488
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14208
                    Top =4716
                    Width =660
                    Height =276
                    ColumnWidth =1440
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =28
                    Name ="BillingRuleId"
                    ControlSource ="BillingRuleId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14208
                    LayoutCachedTop =4716
                    LayoutCachedWidth =14868
                    LayoutCachedHeight =4992
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =85.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =13488
                            Top =4716
                            Width =708
                            Height =276
                            FontSize =9
                            FontWeight =700
                            Name ="BillingRuleId_Label"
                            Caption ="RuleId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =13488
                            LayoutCachedTop =4716
                            LayoutCachedWidth =14196
                            LayoutCachedHeight =4992
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =1
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            BorderShade =65.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =85.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9528
                    Top =4716
                    Width =660
                    Height =276
                    ColumnWidth =1530
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =25
                    Name ="ClientId"
                    ControlSource ="ClientId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9528
                    LayoutCachedTop =4716
                    LayoutCachedWidth =10188
                    LayoutCachedHeight =4992
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =85.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =8808
                            Top =4716
                            Width =708
                            Height =276
                            FontSize =9
                            FontWeight =700
                            Name ="ClientId_Label"
                            Caption ="ClientId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =8808
                            LayoutCachedTop =4716
                            LayoutCachedWidth =9516
                            LayoutCachedHeight =4992
                            LayoutGroup =1
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            BorderShade =65.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =85.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11112
                    Top =4716
                    Width =660
                    Height =276
                    ColumnWidth =1530
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =26
                    Name ="ScheduleId"
                    ControlSource ="ScheduleId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11112
                    LayoutCachedTop =4716
                    LayoutCachedWidth =11772
                    LayoutCachedHeight =4992
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =85.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Left =10200
                            Top =4716
                            Width =900
                            Height =276
                            FontSize =9
                            FontWeight =700
                            Name ="ScheduleId_Label"
                            Caption ="ScheduleId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =10200
                            LayoutCachedTop =4716
                            LayoutCachedWidth =11100
                            LayoutCachedHeight =4992
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            BorderShade =65.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =85.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12816
                    Top =4716
                    Width =660
                    Height =276
                    ColumnWidth =1530
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =27
                    Name ="UnitGroupId"
                    ControlSource ="UnitGroupId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12816
                    LayoutCachedTop =4716
                    LayoutCachedWidth =13476
                    LayoutCachedHeight =4992
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =85.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Left =11784
                            Top =4716
                            Width =1020
                            Height =276
                            FontSize =9
                            FontWeight =700
                            Name ="UnitGroupId_Label"
                            Caption ="UnitGroupId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =11784
                            LayoutCachedTop =4716
                            LayoutCachedWidth =12804
                            LayoutCachedHeight =4992
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            BorderShade =65.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =85.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =11280
                    Top =30
                    Width =4500
                    Height =2850
                    FontSize =10
                    TabIndex =2
                    ForeColor =8210719
                    Name ="SplitFeeDescRule"
                    Tag ="SplitFee"
                    TextFormat =1

                    LayoutCachedLeft =11280
                    LayoutCachedTop =30
                    LayoutCachedWidth =15780
                    LayoutCachedHeight =2880
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin EmptyCell
                    Left =7896
                    Top =2832
                    Width =600
                    Height =288
                    Name ="EmptyCell874"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7896
                    LayoutCachedTop =2832
                    LayoutCachedWidth =8496
                    LayoutCachedHeight =3120
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3708
                    Top =3888
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =9
                    Name ="PerDiemFeeTaxable"
                    ControlSource ="PerDiemFeeTaxable"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"YesNo\") AND ((BillingRules.Deleted)=False));"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    Tag ="PDFee"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =3708
                    LayoutCachedTop =3888
                    LayoutCachedWidth =5028
                    LayoutCachedHeight =4164
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =4
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =2844
                            Top =3888
                            Width =852
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label876"
                            Caption ="Tax Fee?"
                            GroupTable =4
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2844
                            LayoutCachedTop =3888
                            LayoutCachedWidth =3696
                            LayoutCachedHeight =4164
                            RowStart =3
                            RowEnd =3
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =4
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =13536
                    Top =3864
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =22
                    Name ="FeeTaxable"
                    ControlSource ="FeeTaxable"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BillingRules.RuleValue, BillingRules.RuleName FROM BillingRules WHERE ((("
                        "BillingRules.RuleType)=\"YesNo\") AND ((BillingRules.Deleted)=False));"
                    ColumnWidths ="0;2736"
                    OnClick ="[Event Procedure]"
                    Tag ="RentFee"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =13536
                    LayoutCachedTop =3864
                    LayoutCachedWidth =14856
                    LayoutCachedHeight =4140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =12684
                            Top =3864
                            Width =840
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label878"
                            Caption ="Tax Fee?"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =12684
                            LayoutCachedTop =3864
                            LayoutCachedWidth =13524
                            LayoutCachedHeight =4140
                            RowStart =3
                            RowEnd =3
                            ColumnStart =8
                            ColumnEnd =8
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "GroupBillingRules_subform.cls"
