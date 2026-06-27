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
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackColor =-1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
            BorderLineStyle =0
            BorderColor =-1
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
                    BorderColor =10921638
                    Name ="Image142"
                    Picture ="3_WhiteShadow"
                    GridlineColor =10921638
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
                    BorderColor =8355711
                    Name ="Label126"
                    Caption ="Billing and Fee Rules"
                    FontName ="Century Gothic"
                    GridlineColor =10921638
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
                    GridlineColor =10921638
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
                    GridlineColor =10921638
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
                    GridlineColor =10921638
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
                    Left =1215
                    Top =3180
                    Width =1620
                    Height =270
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =3
                    ForeColor =4138256
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1215
                    LayoutCachedTop =3180
                    LayoutCachedWidth =2835
                    LayoutCachedHeight =3450
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
                            Top =3180
                            Width =795
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label147"
                            Caption ="Bill Type"
                            GroupTable =4
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =420
                            LayoutCachedTop =3180
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =3450
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
                    Left =1215
                    Top =3510
                    Width =1620
                    Height =270
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =4
                    ForeColor =4138256
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1215
                    LayoutCachedTop =3510
                    LayoutCachedWidth =2835
                    LayoutCachedHeight =3780
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
                            Top =3510
                            Width =795
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="PerDiemStart_Label"
                            Caption ="Starts On"
                            GroupTable =4
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =420
                            LayoutCachedTop =3510
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =3780
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
                    Top =2835
                    Width =2415
                    Height =285
                    FontWeight =700
                    Name ="PerDiem_Label"
                    Caption ="Per Diem Rules"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =420
                    LayoutCachedTop =2835
                    LayoutCachedWidth =2835
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
                    Left =1215
                    Top =3840
                    Width =1620
                    Height =270
                    FontSize =10
                    TabIndex =5
                    ForeColor =4138256
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1215
                    LayoutCachedTop =3840
                    LayoutCachedWidth =2835
                    LayoutCachedHeight =4110
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
                            Top =3840
                            Width =795
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label161"
                            Caption ="Fee"
                            GroupTable =4
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =420
                            LayoutCachedTop =3840
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =4110
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
                    Left =3690
                    Top =3180
                    Width =1320
                    Height =270
                    FontSize =10
                    TabIndex =7
                    ForeColor =4138256
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =3690
                    LayoutCachedTop =3180
                    LayoutCachedWidth =5010
                    LayoutCachedHeight =3450
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
                            Left =2835
                            Top =3180
                            Width =855
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label173"
                            Caption ="Fee Label"
                            GroupTable =4
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =2835
                            LayoutCachedTop =3180
                            LayoutCachedWidth =3690
                            LayoutCachedHeight =3450
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
                    Left =7050
                    Top =3150
                    Width =1425
                    Height =285
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =7050
                    LayoutCachedTop =3150
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =3435
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
                            Left =6210
                            Top =3150
                            Width =810
                            Height =285
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label185"
                            Caption ="Bill Type"
                            GroupTable =6
                            GridlineColor =10921638
                            LayoutCachedLeft =6210
                            LayoutCachedTop =3150
                            LayoutCachedWidth =7020
                            LayoutCachedHeight =3435
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
                    Left =7050
                    Top =3495
                    Width =1425
                    Height =270
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =7050
                    LayoutCachedTop =3495
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =3765
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
                            Left =6210
                            Top =3495
                            Width =810
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label187"
                            Caption ="Rent Start"
                            GroupTable =6
                            GridlineColor =10921638
                            LayoutCachedLeft =6210
                            LayoutCachedTop =3495
                            LayoutCachedWidth =7020
                            LayoutCachedHeight =3765
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
                    Left =6210
                    Top =2835
                    Width =1662
                    Height =285
                    FontWeight =700
                    Name ="Label188"
                    Caption ="Rent Billing Rules"
                    GroupTable =6
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =6210
                    LayoutCachedTop =2835
                    LayoutCachedWidth =7872
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
                    Left =11535
                    Top =3150
                    Width =1095
                    Height =285
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =11535
                    LayoutCachedTop =3150
                    LayoutCachedWidth =12630
                    LayoutCachedHeight =3435
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
                            Left =10785
                            Top =3150
                            Width =750
                            Height =285
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label190"
                            Caption ="Fee"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =10785
                            LayoutCachedTop =3150
                            LayoutCachedWidth =11535
                            LayoutCachedHeight =3435
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
                    Left =13470
                    Top =3150
                    Width =1320
                    Height =285
                    FontSize =10
                    TabIndex =20
                    ForeColor =4138256
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =13470
                    LayoutCachedTop =3150
                    LayoutCachedWidth =14790
                    LayoutCachedHeight =3435
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
                            Left =12630
                            Top =3150
                            Width =840
                            Height =285
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label194"
                            Caption ="Fee Label"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =12630
                            LayoutCachedTop =3150
                            LayoutCachedWidth =13470
                            LayoutCachedHeight =3435
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
                    Left =9285
                    Top =3150
                    Width =1500
                    Height =285
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =9285
                    LayoutCachedTop =3150
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =3435
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
                            Left =8475
                            Top =3150
                            Width =810
                            Height =285
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label200"
                            Caption ="Remit To"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =8475
                            LayoutCachedTop =3150
                            LayoutCachedWidth =9285
                            LayoutCachedHeight =3435
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
                    Left =9285
                    Top =3495
                    Width =1500
                    Height =270
                    FontSize =10
                    TabIndex =24
                    BackColor =15921906
                    Name ="Remittee"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =9285
                    LayoutCachedTop =3495
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =3765
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
                End
                Begin EmptyCell
                    Left =11535
                    Top =2835
                    Width =1095
                    Height =285
                    Name ="EmptyCell220"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =11535
                    LayoutCachedTop =2835
                    LayoutCachedWidth =12630
                    LayoutCachedHeight =3120
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =10785
                    Top =2835
                    Width =750
                    Height =285
                    Name ="EmptyCell230"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =10785
                    LayoutCachedTop =2835
                    LayoutCachedWidth =11535
                    LayoutCachedHeight =3120
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =7875
                    Top =3825
                    Width =600
                    Height =270
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="DayofMthLabel"
                    Caption ="th  Day"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =7875
                    LayoutCachedTop =3825
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =4095
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
                    Left =11535
                    Top =3825
                    Width =1095
                    Height =270
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
                    GridlineColor =10921638

                    LayoutCachedLeft =11535
                    LayoutCachedTop =3825
                    LayoutCachedWidth =12630
                    LayoutCachedHeight =4095
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
                            Left =10785
                            Top =3825
                            Width =750
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label274"
                            Caption ="Term"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =10785
                            LayoutCachedTop =3825
                            LayoutCachedWidth =11535
                            LayoutCachedHeight =4095
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
                    Left =7050
                    Top =3825
                    Width =822
                    Height =270
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
                    GridlineColor =10921638

                    LayoutCachedLeft =7050
                    LayoutCachedTop =3825
                    LayoutCachedWidth =7872
                    LayoutCachedHeight =4095
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
                            Left =6210
                            Top =3825
                            Width =810
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label255"
                            Caption ="Bill On"
                            GroupTable =6
                            GridlineColor =10921638
                            LayoutCachedLeft =6210
                            LayoutCachedTop =3825
                            LayoutCachedWidth =7020
                            LayoutCachedHeight =4095
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =6
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =87
                    TextAlign =3
                    Left =8475
                    Top =3495
                    Width =810
                    Height =270
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =8355711
                    Name ="Label485"
                    Caption ="Assignee"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =8475
                    LayoutCachedTop =3495
                    LayoutCachedWidth =9285
                    LayoutCachedHeight =3765
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =6
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
                    BorderColor =10921638
                    ForeColor =10040879
                    Name ="BillDescRule"
                    GridlineColor =10921638
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
                    Left =9285
                    Top =3825
                    Width =1500
                    Height =270
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =9285
                    LayoutCachedTop =3825
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =4095
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
                            Left =8475
                            Top =3825
                            Width =810
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label531"
                            Caption ="Contract"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =8475
                            LayoutCachedTop =3825
                            LayoutCachedWidth =9285
                            LayoutCachedHeight =4095
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
                    Left =13470
                    Top =3495
                    Width =1320
                    Height =270
                    FontSize =10
                    TabIndex =21
                    ForeColor =4138256
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =13470
                    LayoutCachedTop =3495
                    LayoutCachedWidth =14790
                    LayoutCachedHeight =3765
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
                            Left =12630
                            Top =3495
                            Width =840
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label538"
                            Caption ="Show Fee"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =12630
                            LayoutCachedTop =3495
                            LayoutCachedWidth =13470
                            LayoutCachedHeight =3765
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
                    Left =11535
                    Top =4155
                    Width =1095
                    Height =270
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =11535
                    LayoutCachedTop =4155
                    LayoutCachedWidth =12630
                    LayoutCachedHeight =4425
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
                            Left =10785
                            Top =4155
                            Width =750
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label544"
                            Caption ="Split Bill"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =10785
                            LayoutCachedTop =4155
                            LayoutCachedWidth =11535
                            LayoutCachedHeight =4425
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
                    Left =13470
                    Top =4155
                    Width =1320
                    Height =270
                    FontSize =10
                    TabIndex =23
                    ForeColor =4138256
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =13470
                    LayoutCachedTop =4155
                    LayoutCachedWidth =14790
                    LayoutCachedHeight =4425
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
                            Left =12630
                            Top =4155
                            Width =840
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label550"
                            Caption ="Refund"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =12630
                            LayoutCachedTop =4155
                            LayoutCachedWidth =13470
                            LayoutCachedHeight =4425
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
                    Left =3690
                    Top =3510
                    Width =1320
                    Height =270
                    FontSize =10
                    TabIndex =8
                    ForeColor =4138256
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =3690
                    LayoutCachedTop =3510
                    LayoutCachedWidth =5010
                    LayoutCachedHeight =3780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =4
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3690
                    Top =4170
                    Width =1320
                    Height =285
                    FontSize =10
                    TabIndex =10
                    ForeColor =4138256
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =3690
                    LayoutCachedTop =4170
                    LayoutCachedWidth =5010
                    LayoutCachedHeight =4455
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    GroupTable =4
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
                    BorderColor =10921638
                    ForeColor =10040879
                    Name ="DescPerDiem"
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedLeft =405
                    LayoutCachedTop =30
                    LayoutCachedWidth =5505
                    LayoutCachedHeight =2880
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin EmptyCell
                    Left =3690
                    Top =2835
                    Width =1320
                    Height =285
                    Name ="EmptyCell596"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =3690
                    LayoutCachedTop =2835
                    LayoutCachedWidth =5010
                    LayoutCachedHeight =3120
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =2835
                    Top =2835
                    Width =855
                    Height =285
                    Name ="EmptyCell603"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2835
                    LayoutCachedTop =2835
                    LayoutCachedWidth =3690
                    LayoutCachedHeight =3120
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =13470
                    Top =2835
                    Width =1320
                    Height =285
                    Name ="EmptyCell622"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =13470
                    LayoutCachedTop =2835
                    LayoutCachedWidth =14790
                    LayoutCachedHeight =3120
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =12630
                    Top =2835
                    Width =840
                    Height =285
                    Name ="EmptyCell636"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =12630
                    LayoutCachedTop =2835
                    LayoutCachedWidth =13470
                    LayoutCachedHeight =3120
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =9285
                    Top =2835
                    Width =1500
                    Height =285
                    Name ="EmptyCell649"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =9285
                    LayoutCachedTop =2835
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =3120
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =8475
                    Top =2835
                    Width =810
                    Height =285
                    Name ="EmptyCell663"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =8475
                    LayoutCachedTop =2835
                    LayoutCachedWidth =9285
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
                    Left =1215
                    Top =4170
                    Width =1620
                    Height =285
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
                    GridlineColor =10921638

                    LayoutCachedLeft =1215
                    LayoutCachedTop =4170
                    LayoutCachedWidth =2835
                    LayoutCachedHeight =4455
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
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11535
                    Top =3495
                    Width =1095
                    Height =270
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
                    GridlineColor =10921638

                    LayoutCachedLeft =11535
                    LayoutCachedTop =3495
                    LayoutCachedWidth =12630
                    LayoutCachedHeight =3765
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
                            Left =10785
                            Top =3495
                            Width =750
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label688"
                            Caption ="Amount"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =10785
                            LayoutCachedTop =3495
                            LayoutCachedWidth =11535
                            LayoutCachedHeight =3765
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
                    Left =6210
                    Top =4155
                    Width =810
                    Height =270
                    Name ="EmptyCell731"
                    GroupTable =6
                    GridlineColor =10921638
                    LayoutCachedLeft =6210
                    LayoutCachedTop =4155
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =4425
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =7050
                    Top =4155
                    Width =822
                    Height =270
                    Name ="EmptyCell732"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =7050
                    LayoutCachedTop =4155
                    LayoutCachedWidth =7872
                    LayoutCachedHeight =4425
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =7875
                    Top =4155
                    Width =600
                    Height =270
                    Name ="EmptyCell733"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =7875
                    LayoutCachedTop =4155
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =4425
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =8475
                    Top =4155
                    Width =810
                    Height =270
                    Name ="EmptyCell734"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =8475
                    LayoutCachedTop =4155
                    LayoutCachedWidth =9285
                    LayoutCachedHeight =4425
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =9285
                    Top =4155
                    Width =1500
                    Height =270
                    Name ="EmptyCell735"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =9285
                    LayoutCachedTop =4155
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =4425
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =10785
                    Top =4155
                    Width =0
                    Height =270
                    Name ="EmptyCell736"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =10785
                    LayoutCachedTop =4155
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =4425
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =420
                    Top =4170
                    Width =795
                    Height =285
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =8355711
                    Name ="Label681"
                    Caption ="Amount"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =420
                    LayoutCachedTop =4170
                    LayoutCachedWidth =1215
                    LayoutCachedHeight =4455
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =2
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =3
                    Left =2835
                    Top =3510
                    Width =855
                    Height =270
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =8355711
                    Name ="Label556"
                    Caption ="Show Fee"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2835
                    LayoutCachedTop =3510
                    LayoutCachedWidth =3690
                    LayoutCachedHeight =3780
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =3
                    Left =2835
                    Top =4170
                    Width =855
                    Height =285
                    FontSize =9
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =8355711
                    Name ="Label558"
                    Caption ="Refund"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2835
                    LayoutCachedTop =4170
                    LayoutCachedWidth =3690
                    LayoutCachedHeight =4455
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14115
                    Top =4710
                    Width =660
                    Height =270
                    ColumnWidth =1440
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =28
                    BorderColor =10921638
                    Name ="BillingRuleId"
                    ControlSource ="BillingRuleId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =14115
                    LayoutCachedTop =4710
                    LayoutCachedWidth =14775
                    LayoutCachedHeight =4980
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
                            Left =13410
                            Top =4710
                            Width =705
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="BillingRuleId_Label"
                            Caption ="RuleId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =13410
                            LayoutCachedTop =4710
                            LayoutCachedWidth =14115
                            LayoutCachedHeight =4980
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
                    Left =9510
                    Top =4710
                    Width =660
                    Height =270
                    ColumnWidth =1530
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =25
                    BorderColor =10921638
                    Name ="ClientId"
                    ControlSource ="ClientId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =9510
                    LayoutCachedTop =4710
                    LayoutCachedWidth =10170
                    LayoutCachedHeight =4980
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
                            Left =8805
                            Top =4710
                            Width =705
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="ClientId_Label"
                            Caption ="ClientId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =8805
                            LayoutCachedTop =4710
                            LayoutCachedWidth =9510
                            LayoutCachedHeight =4980
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
                    Left =11070
                    Top =4710
                    Width =660
                    Height =270
                    ColumnWidth =1530
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =26
                    BorderColor =10921638
                    Name ="ScheduleId"
                    ControlSource ="ScheduleId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =11070
                    LayoutCachedTop =4710
                    LayoutCachedWidth =11730
                    LayoutCachedHeight =4980
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
                            Left =10170
                            Top =4710
                            Width =900
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="ScheduleId_Label"
                            Caption ="ScheduleId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =10170
                            LayoutCachedTop =4710
                            LayoutCachedWidth =11070
                            LayoutCachedHeight =4980
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
                    Left =12750
                    Top =4710
                    Width =660
                    Height =270
                    ColumnWidth =1530
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =27
                    BorderColor =10921638
                    Name ="UnitGroupId"
                    ControlSource ="UnitGroupId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =12750
                    LayoutCachedTop =4710
                    LayoutCachedWidth =13410
                    LayoutCachedHeight =4980
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
                            Left =11730
                            Top =4710
                            Width =1020
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="UnitGroupId_Label"
                            Caption ="UnitGroupId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =11730
                            LayoutCachedTop =4710
                            LayoutCachedWidth =12750
                            LayoutCachedHeight =4980
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
                    BorderColor =10921638
                    ForeColor =8210719
                    Name ="SplitFeeDescRule"
                    Tag ="SplitFee"
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedLeft =11280
                    LayoutCachedTop =30
                    LayoutCachedWidth =15780
                    LayoutCachedHeight =2880
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin EmptyCell
                    Left =7875
                    Top =2835
                    Width =600
                    Height =285
                    Name ="EmptyCell874"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =7875
                    LayoutCachedTop =2835
                    LayoutCachedWidth =8475
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
                    Left =3690
                    Top =3840
                    Width =1320
                    Height =270
                    FontSize =10
                    TabIndex =9
                    ForeColor =4138256
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =3690
                    LayoutCachedTop =3840
                    LayoutCachedWidth =5010
                    LayoutCachedHeight =4110
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
                            Left =2835
                            Top =3840
                            Width =855
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label876"
                            Caption ="Tax Fee?"
                            GroupTable =4
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =2835
                            LayoutCachedTop =3840
                            LayoutCachedWidth =3690
                            LayoutCachedHeight =4110
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
                    Left =13470
                    Top =3825
                    Width =1320
                    Height =270
                    FontSize =10
                    TabIndex =22
                    ForeColor =4138256
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
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =13470
                    LayoutCachedTop =3825
                    LayoutCachedWidth =14790
                    LayoutCachedHeight =4095
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
                            Left =12630
                            Top =3825
                            Width =840
                            Height =270
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =8355711
                            Name ="Label878"
                            Caption ="Tax Fee?"
                            GroupTable =6
                            LeftPadding =0
                            RightPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =12630
                            LayoutCachedTop =3825
                            LayoutCachedWidth =13470
                            LayoutCachedHeight =4095
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
