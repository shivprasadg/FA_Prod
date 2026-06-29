Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =16416
    DatasheetFontHeight =11
    ItemSuffix =917
    Left =600
    Top =9375
    Right =16935
    Bottom =15195
    AfterInsert ="[Event Procedure]"
    RecSrcDt = Begin
        0x3460a36f0c8fe440
    End
    RecordSource ="SELECT BillingRulesClient.* FROM BillingRulesClient;"
    Caption ="GroupBillingRules_subform"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
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
        Begin ListBox
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin ToggleButton
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
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
                    TabIndex =2
                End
                Begin Label
                    OverlapFlags =247
                    Left =120
                    Width =7860
                    Height =540
                    FontSize =20
                    Name ="Label126"
                    Caption ="Schedule Level Billing and Fee Rules"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =540
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =75.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14220
                    Top =75
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="SaveUpdate"
                    Caption ="Commit Update"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =14220
                    LayoutCachedTop =75
                    LayoutCachedWidth =15660
                    LayoutCachedHeight =345
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    UseTheme =1
                    Shape =0
                    Gradient =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =119
                    Left =11235
                    Top =90
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="SampleInvoice"
                    Caption ="Sample Invoice"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =11235
                    LayoutCachedTop =90
                    LayoutCachedWidth =12675
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    UseTheme =1
                    Shape =0
                    Gradient =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =5238
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    Visible = NotDefault
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
                    Left =180
                    Top =240
                    Width =285
                    Height =2736
                    Name ="Box866"
                    LayoutCachedLeft =180
                    LayoutCachedTop =240
                    LayoutCachedWidth =465
                    LayoutCachedHeight =2976
                    BorderThemeColorIndex =2
                End
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =93
                    Left =6000
                    Top =240
                    Width =285
                    Height =2736
                    Name ="Box865"
                    LayoutCachedLeft =6000
                    LayoutCachedTop =240
                    LayoutCachedWidth =6285
                    LayoutCachedHeight =2976
                    BorderThemeColorIndex =2
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1164
                    Top =3180
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

                    LayoutCachedLeft =1164
                    LayoutCachedTop =3180
                    LayoutCachedWidth =2784
                    LayoutCachedHeight =3456
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
                            Left =360
                            Top =3180
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
                            LayoutCachedLeft =360
                            LayoutCachedTop =3180
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =3456
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
                    Left =1164
                    Top =3528
                    Width =1620
                    Height =276
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =5
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

                    LayoutCachedLeft =1164
                    LayoutCachedTop =3528
                    LayoutCachedWidth =2784
                    LayoutCachedHeight =3804
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
                            Left =360
                            Top =3528
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
                            LayoutCachedLeft =360
                            LayoutCachedTop =3528
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =3804
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
                    Left =360
                    Top =2820
                    Width =2424
                    Height =285
                    FontWeight =700
                    Name ="PerDiem_Label"
                    Caption ="Per Diem Rules"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =360
                    LayoutCachedTop =2820
                    LayoutCachedWidth =2784
                    LayoutCachedHeight =3105
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
                    Left =1164
                    Top =3876
                    Width =1620
                    Height =276
                    FontSize =10
                    TabIndex =7
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

                    LayoutCachedLeft =1164
                    LayoutCachedTop =3876
                    LayoutCachedWidth =2784
                    LayoutCachedHeight =4152
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
                            Left =360
                            Top =3876
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
                            LayoutCachedLeft =360
                            LayoutCachedTop =3876
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =4152
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
                    Left =3648
                    Top =3180
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =4
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

                    LayoutCachedLeft =3648
                    LayoutCachedTop =3180
                    LayoutCachedWidth =4968
                    LayoutCachedHeight =3456
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
                            Left =2784
                            Top =3180
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
                            LayoutCachedLeft =2784
                            LayoutCachedTop =3180
                            LayoutCachedWidth =3636
                            LayoutCachedHeight =3456
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
                    Left =6972
                    Top =3144
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

                    LayoutCachedLeft =6972
                    LayoutCachedTop =3144
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =3432
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
                            Left =6120
                            Top =3144
                            Width =810
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
                            LayoutCachedLeft =6120
                            LayoutCachedTop =3144
                            LayoutCachedWidth =6930
                            LayoutCachedHeight =3432
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
                    Left =6972
                    Top =3504
                    Width =1428
                    Height =276
                    FontSize =10
                    TabIndex =15
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

                    LayoutCachedLeft =6972
                    LayoutCachedTop =3504
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =3780
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
                            Left =6120
                            Top =3504
                            Width =810
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
                            LayoutCachedLeft =6120
                            LayoutCachedTop =3504
                            LayoutCachedWidth =6930
                            LayoutCachedHeight =3780
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
                    Left =6120
                    Top =2820
                    Width =1680
                    Height =285
                    FontWeight =700
                    Name ="Label188"
                    Caption ="Rent Billing Rules"
                    GroupTable =6
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6120
                    LayoutCachedTop =2820
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =3105
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
                    Left =11496
                    Top =3144
                    Width =1092
                    Height =288
                    FontSize =10
                    TabIndex =13
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

                    LayoutCachedLeft =11496
                    LayoutCachedTop =3144
                    LayoutCachedWidth =12588
                    LayoutCachedHeight =3432
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
                            Left =10728
                            Top =3144
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
                            LayoutCachedLeft =10728
                            LayoutCachedTop =3144
                            LayoutCachedWidth =11484
                            LayoutCachedHeight =3432
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
                    Left =13440
                    Top =3144
                    Width =1320
                    Height =288
                    FontSize =10
                    TabIndex =14
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

                    LayoutCachedLeft =13440
                    LayoutCachedTop =3144
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =3432
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
                            Left =12588
                            Top =3144
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
                            LayoutCachedLeft =12588
                            LayoutCachedTop =3144
                            LayoutCachedWidth =13428
                            LayoutCachedHeight =3432
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
                    Left =9228
                    Top =3144
                    Width =1500
                    Height =288
                    FontSize =10
                    TabIndex =12
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

                    LayoutCachedLeft =9228
                    LayoutCachedTop =3144
                    LayoutCachedWidth =10728
                    LayoutCachedHeight =3432
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
                            Left =8400
                            Top =3144
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
                            LayoutCachedLeft =8400
                            LayoutCachedTop =3144
                            LayoutCachedWidth =9216
                            LayoutCachedHeight =3432
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
                    Visible = NotDefault
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9228
                    Top =3852
                    Width =1500
                    Height =276
                    FontSize =10
                    TabIndex =20
                    BackColor =15921906
                    Name ="Remittee"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =9228
                    LayoutCachedTop =3852
                    LayoutCachedWidth =10728
                    LayoutCachedHeight =4128
                    RowStart =3
                    RowEnd =3
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
                    Left =11496
                    Top =2820
                    Width =1092
                    Height =285
                    Name ="EmptyCell220"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11496
                    LayoutCachedTop =2820
                    LayoutCachedWidth =12588
                    LayoutCachedHeight =3105
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =10728
                    Top =2820
                    Width =756
                    Height =285
                    Name ="EmptyCell230"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10728
                    LayoutCachedTop =2820
                    LayoutCachedWidth =11484
                    LayoutCachedHeight =3105
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =7800
                    Top =3852
                    Width =600
                    Height =276
                    FontSize =9
                    FontWeight =700
                    Name ="DayofMthLabel"
                    Caption ="th  Day"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =7800
                    LayoutCachedTop =3852
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =4128
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
                    Left =11496
                    Top =3852
                    Width =1092
                    Height =276
                    FontSize =10
                    TabIndex =21
                    BackColor =15921906
                    Name ="FeeTerm"
                    ControlSource ="FeeTerm"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="25"
                    Tag ="RentFee"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =11496
                    LayoutCachedTop =3852
                    LayoutCachedWidth =12588
                    LayoutCachedHeight =4128
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
                            OverlapFlags =87
                            TextAlign =3
                            Left =10728
                            Top =3852
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
                            LayoutCachedLeft =10728
                            LayoutCachedTop =3852
                            LayoutCachedWidth =11484
                            LayoutCachedHeight =4128
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
                    Left =6972
                    Top =3852
                    Width =828
                    Height =276
                    FontSize =10
                    TabIndex =19
                    BackColor =15921906
                    Name ="DayOfMonth"
                    ControlSource ="DayOfMonth"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="25"
                    Tag ="RENT"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =6972
                    LayoutCachedTop =3852
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =4128
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
                            Left =6120
                            Top =3852
                            Width =810
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
                            LayoutCachedLeft =6120
                            LayoutCachedTop =3852
                            LayoutCachedWidth =6930
                            LayoutCachedHeight =4128
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
                    Visible = NotDefault
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =8400
                    Top =3852
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
                    LayoutCachedLeft =8400
                    LayoutCachedTop =3852
                    LayoutCachedWidth =9216
                    LayoutCachedHeight =4128
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
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
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9228
                    Top =3504
                    Width =1500
                    Height =276
                    FontSize =10
                    TabIndex =16
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

                    LayoutCachedLeft =9228
                    LayoutCachedTop =3504
                    LayoutCachedWidth =10728
                    LayoutCachedHeight =3780
                    DatasheetCaption ="`"
                    RowStart =2
                    RowEnd =2
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
                            Left =8400
                            Top =3504
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
                            LayoutCachedLeft =8400
                            LayoutCachedTop =3504
                            LayoutCachedWidth =9216
                            LayoutCachedHeight =3780
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
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =13440
                    Top =3504
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =18
                    Name ="ShowFeeColumn"
                    ControlSource ="ShowFeeColumn"
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

                    LayoutCachedLeft =13440
                    LayoutCachedTop =3504
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =3780
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
                            Left =12588
                            Top =3504
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
                            LayoutCachedLeft =12588
                            LayoutCachedTop =3504
                            LayoutCachedWidth =13428
                            LayoutCachedHeight =3780
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
                    Left =11496
                    Top =4200
                    Width =1092
                    Height =276
                    FontSize =10
                    TabIndex =25
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

                    LayoutCachedLeft =11496
                    LayoutCachedTop =4200
                    LayoutCachedWidth =12588
                    LayoutCachedHeight =4476
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
                            Left =10728
                            Top =4200
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
                            LayoutCachedLeft =10728
                            LayoutCachedTop =4200
                            LayoutCachedWidth =11484
                            LayoutCachedHeight =4476
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
                    Left =13440
                    Top =4200
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =26
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

                    LayoutCachedLeft =13440
                    LayoutCachedTop =4200
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =4476
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
                            Left =12588
                            Top =4200
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
                            LayoutCachedLeft =12588
                            LayoutCachedTop =4200
                            LayoutCachedWidth =13428
                            LayoutCachedHeight =4476
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
                    Left =3648
                    Top =3528
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =6
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

                    LayoutCachedLeft =3648
                    LayoutCachedTop =3528
                    LayoutCachedWidth =4968
                    LayoutCachedHeight =3804
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
                            Left =2784
                            Top =3528
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
                            LayoutCachedLeft =2784
                            LayoutCachedTop =3528
                            LayoutCachedWidth =3636
                            LayoutCachedHeight =3804
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
                    Left =3648
                    Top =4224
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

                    LayoutCachedLeft =3648
                    LayoutCachedTop =4224
                    LayoutCachedWidth =4968
                    LayoutCachedHeight =4512
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
                            OverlapFlags =93
                            TextAlign =3
                            Left =2784
                            Top =4224
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
                            LayoutCachedLeft =2784
                            LayoutCachedTop =4224
                            LayoutCachedWidth =3636
                            LayoutCachedHeight =4512
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
                    Locked = NotDefault
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
                    Left =3648
                    Top =2820
                    Width =1320
                    Height =285
                    Name ="EmptyCell596"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =3648
                    LayoutCachedTop =2820
                    LayoutCachedWidth =4968
                    LayoutCachedHeight =3105
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =2784
                    Top =2820
                    Width =852
                    Height =285
                    Name ="EmptyCell603"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2784
                    LayoutCachedTop =2820
                    LayoutCachedWidth =3636
                    LayoutCachedHeight =3105
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =13440
                    Top =2820
                    Width =1320
                    Height =285
                    Name ="EmptyCell622"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13440
                    LayoutCachedTop =2820
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =3105
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =12588
                    Top =2820
                    Width =840
                    Height =285
                    Name ="EmptyCell636"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12588
                    LayoutCachedTop =2820
                    LayoutCachedWidth =13428
                    LayoutCachedHeight =3105
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =9228
                    Top =2820
                    Width =1500
                    Height =285
                    Name ="EmptyCell649"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9228
                    LayoutCachedTop =2820
                    LayoutCachedWidth =10728
                    LayoutCachedHeight =3105
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =8400
                    Top =2820
                    Width =816
                    Height =285
                    Name ="EmptyCell663"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8400
                    LayoutCachedTop =2820
                    LayoutCachedWidth =9216
                    LayoutCachedHeight =3105
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1164
                    Top =4224
                    Width =1620
                    Height =288
                    FontSize =10
                    TabIndex =9
                    Name ="PerDiemFee"
                    ControlSource ="PerDiemFeeValue"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="65"
                    Tag ="RENT"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1164
                    LayoutCachedTop =4224
                    LayoutCachedWidth =2784
                    LayoutCachedHeight =4512
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
                            Left =360
                            Top =4224
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
                            LayoutCachedLeft =360
                            LayoutCachedTop =4224
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =4512
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
                    Left =11496
                    Top =3504
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

                    LayoutCachedLeft =11496
                    LayoutCachedTop =3504
                    LayoutCachedWidth =12588
                    LayoutCachedHeight =3780
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
                            Left =10728
                            Top =3504
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
                            LayoutCachedLeft =10728
                            LayoutCachedTop =3504
                            LayoutCachedWidth =11484
                            LayoutCachedHeight =3780
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
                    Left =10728
                    Top =4200
                    Width =0
                    Height =276
                    Name ="EmptyCell736"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =10728
                    LayoutCachedTop =4200
                    LayoutCachedWidth =10728
                    LayoutCachedHeight =4476
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
                    Left =14148
                    Top =4836
                    Width =660
                    Height =276
                    ColumnWidth =1440
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =31
                    Name ="BillingRuleId"
                    ControlSource ="BillingRuleId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120

                    LayoutCachedLeft =14148
                    LayoutCachedTop =4836
                    LayoutCachedWidth =14808
                    LayoutCachedHeight =5112
                    ColumnStart =8
                    ColumnEnd =8
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
                            Left =13428
                            Top =4836
                            Width =708
                            Height =276
                            FontSize =9
                            FontWeight =700
                            Name ="BillingRuleId_Label"
                            Caption ="RuleId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =15
                            RightPadding =0
                            BottomPadding =120
                            LayoutCachedLeft =13428
                            LayoutCachedTop =4836
                            LayoutCachedWidth =14136
                            LayoutCachedHeight =5112
                            ColumnStart =7
                            ColumnEnd =7
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
                    Left =9468
                    Top =4836
                    Width =660
                    Height =276
                    ColumnWidth =1530
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =28
                    Name ="ClientId"
                    ControlSource ="ClientId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120

                    LayoutCachedLeft =9468
                    LayoutCachedTop =4836
                    LayoutCachedWidth =10128
                    LayoutCachedHeight =5112
                    ColumnStart =2
                    ColumnEnd =2
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
                            Left =8748
                            Top =4836
                            Width =708
                            Height =276
                            FontSize =9
                            FontWeight =700
                            Name ="ClientId_Label"
                            Caption ="ClientId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =15
                            RightPadding =0
                            BottomPadding =120
                            LayoutCachedLeft =8748
                            LayoutCachedTop =4836
                            LayoutCachedWidth =9456
                            LayoutCachedHeight =5112
                            ColumnStart =1
                            ColumnEnd =1
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
                    Left =11052
                    Top =4836
                    Width =660
                    Height =276
                    ColumnWidth =1530
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =29
                    Name ="ScheduleId"
                    ControlSource ="ScheduleId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120

                    LayoutCachedLeft =11052
                    LayoutCachedTop =4836
                    LayoutCachedWidth =11712
                    LayoutCachedHeight =5112
                    ColumnStart =4
                    ColumnEnd =4
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
                            Left =10140
                            Top =4836
                            Width =900
                            Height =276
                            FontSize =9
                            FontWeight =700
                            Name ="ScheduleId_Label"
                            Caption ="ScheduleId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =15
                            RightPadding =0
                            BottomPadding =120
                            LayoutCachedLeft =10140
                            LayoutCachedTop =4836
                            LayoutCachedWidth =11040
                            LayoutCachedHeight =5112
                            ColumnStart =3
                            ColumnEnd =3
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
                    Left =12756
                    Top =4836
                    Width =660
                    Height =276
                    ColumnWidth =1530
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =30
                    Name ="UnitGroupId"
                    ControlSource ="UnitGroupId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120

                    LayoutCachedLeft =12756
                    LayoutCachedTop =4836
                    LayoutCachedWidth =13416
                    LayoutCachedHeight =5112
                    ColumnStart =6
                    ColumnEnd =6
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
                            Left =11724
                            Top =4836
                            Width =1020
                            Height =276
                            FontSize =9
                            FontWeight =700
                            Name ="UnitGroupId_Label"
                            Caption ="UnitGroupId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =15
                            RightPadding =0
                            BottomPadding =120
                            LayoutCachedLeft =11724
                            LayoutCachedTop =4836
                            LayoutCachedWidth =12744
                            LayoutCachedHeight =5112
                            ColumnStart =5
                            ColumnEnd =5
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
                    Left =7800
                    Top =2820
                    Width =600
                    Height =285
                    Name ="EmptyCell874"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7800
                    LayoutCachedTop =2820
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =3105
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
                    Left =3648
                    Top =3876
                    Width =1320
                    Height =276
                    FontSize =10
                    TabIndex =8
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

                    LayoutCachedLeft =3648
                    LayoutCachedTop =3876
                    LayoutCachedWidth =4968
                    LayoutCachedHeight =4152
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
                            Left =2784
                            Top =3876
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
                            LayoutCachedLeft =2784
                            LayoutCachedTop =3876
                            LayoutCachedWidth =3636
                            LayoutCachedHeight =4152
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
                    Left =13440
                    Top =3852
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

                    LayoutCachedLeft =13440
                    LayoutCachedTop =3852
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =4128
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
                            Left =12588
                            Top =3852
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
                            LayoutCachedLeft =12588
                            LayoutCachedTop =3852
                            LayoutCachedWidth =13428
                            LayoutCachedHeight =4128
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
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =14895
                    Top =3165
                    Height =1620
                    FontSize =9
                    TabIndex =27
                    Name ="GroupList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SGrpID, UnitGroup As Groups, Count(AssetID) AS Units, SchID From vw_SixKe"
                        "ys GROUP BY SGrpID, UnitGroup, SchID HAVING SchID=1273 Order BY UnitGroup; "
                    ColumnWidths ="0;864;576"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =14895
                    LayoutCachedTop =3165
                    LayoutCachedWidth =16335
                    LayoutCachedHeight =4785
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =14895
                            Top =2970
                            Width =765
                            Height =195
                            FontSize =9
                            FontWeight =700
                            Name ="Label900"
                            Caption ="Groups"
                            LayoutCachedLeft =14895
                            LayoutCachedTop =2970
                            LayoutCachedWidth =15660
                            LayoutCachedHeight =3165
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                        End
                    End
                End
                Begin EmptyCell
                    Left =7656
                    Top =4836
                    Width =1080
                    Height =276
                    Name ="EmptyCell901"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120
                    LayoutCachedLeft =7656
                    LayoutCachedTop =4836
                    LayoutCachedWidth =8736
                    LayoutCachedHeight =5112
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =8400
                    Top =4200
                    Width =816
                    Height =276
                    Name ="EmptyCell907"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8400
                    LayoutCachedTop =4200
                    LayoutCachedWidth =9216
                    LayoutCachedHeight =4476
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9228
                    Top =4200
                    Width =1500
                    Height =276
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =24
                    ForeColor =10040879
                    Name ="BillingScript"
                    ControlSource ="Script"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    TextFormat =1

                    LayoutCachedLeft =9228
                    LayoutCachedTop =4200
                    LayoutCachedWidth =10728
                    LayoutCachedHeight =4476
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =6
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6972
                    Top =4200
                    Width =1428
                    Height =276
                    FontSize =10
                    TabIndex =23
                    BackColor =15921906
                    Name ="CustomStartDate"
                    ControlSource ="CustomStartDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =6972
                    LayoutCachedTop =4200
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =4476
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =2
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
                            Left =6120
                            Top =4200
                            Width =810
                            Height =276
                            FontSize =9
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label914"
                            Caption ="Custom"
                            GroupTable =6
                            LayoutCachedLeft =6120
                            LayoutCachedTop =4200
                            LayoutCachedWidth =6930
                            LayoutCachedHeight =4476
                            RowStart =4
                            RowEnd =4
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
' See "ScheduleBillingRules_subform.cls"
