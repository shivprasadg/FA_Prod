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
            Height =5220
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
                    Left =1155
                    Top =3165
                    Width =1620
                    Height =270
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

                    LayoutCachedLeft =1155
                    LayoutCachedTop =3165
                    LayoutCachedWidth =2775
                    LayoutCachedHeight =3435
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
                            Top =3165
                            Width =795
                            Height =270
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
                            LayoutCachedTop =3165
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =3435
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
                    Left =1155
                    Top =3495
                    Width =1620
                    Height =270
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

                    LayoutCachedLeft =1155
                    LayoutCachedTop =3495
                    LayoutCachedWidth =2775
                    LayoutCachedHeight =3765
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
                            Top =3495
                            Width =795
                            Height =270
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
                            LayoutCachedTop =3495
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =3765
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
                    Width =2415
                    Height =285
                    FontWeight =700
                    Name ="PerDiem_Label"
                    Caption ="Per Diem Rules"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =360
                    LayoutCachedTop =2820
                    LayoutCachedWidth =2775
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
                    Left =1155
                    Top =3825
                    Width =1620
                    Height =270
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

                    LayoutCachedLeft =1155
                    LayoutCachedTop =3825
                    LayoutCachedWidth =2775
                    LayoutCachedHeight =4095
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
                            Top =3825
                            Width =795
                            Height =270
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
                            LayoutCachedTop =3825
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =4095
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
                    Left =3630
                    Top =3165
                    Width =1320
                    Height =270
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

                    LayoutCachedLeft =3630
                    LayoutCachedTop =3165
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =3435
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
                            Left =2775
                            Top =3165
                            Width =855
                            Height =270
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
                            LayoutCachedLeft =2775
                            LayoutCachedTop =3165
                            LayoutCachedWidth =3630
                            LayoutCachedHeight =3435
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
                    Left =6960
                    Top =3135
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
                    AllowValueListEdits =0

                    LayoutCachedLeft =6960
                    LayoutCachedTop =3135
                    LayoutCachedWidth =8385
                    LayoutCachedHeight =3420
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
                            Top =3135
                            Width =810
                            Height =285
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
                            LayoutCachedTop =3135
                            LayoutCachedWidth =6930
                            LayoutCachedHeight =3420
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
                    Left =6960
                    Top =3480
                    Width =1425
                    Height =270
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

                    LayoutCachedLeft =6960
                    LayoutCachedTop =3480
                    LayoutCachedWidth =8385
                    LayoutCachedHeight =3750
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
                            Top =3480
                            Width =810
                            Height =270
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
                            LayoutCachedTop =3480
                            LayoutCachedWidth =6930
                            LayoutCachedHeight =3750
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
                    Width =1662
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
                    LayoutCachedWidth =7782
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
                    Left =11445
                    Top =3135
                    Width =1095
                    Height =285
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

                    LayoutCachedLeft =11445
                    LayoutCachedTop =3135
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =3420
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
                            Left =10695
                            Top =3135
                            Width =750
                            Height =285
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
                            LayoutCachedLeft =10695
                            LayoutCachedTop =3135
                            LayoutCachedWidth =11445
                            LayoutCachedHeight =3420
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
                    Left =13380
                    Top =3135
                    Width =1320
                    Height =285
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

                    LayoutCachedLeft =13380
                    LayoutCachedTop =3135
                    LayoutCachedWidth =14700
                    LayoutCachedHeight =3420
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
                            Left =12540
                            Top =3135
                            Width =840
                            Height =285
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
                            LayoutCachedLeft =12540
                            LayoutCachedTop =3135
                            LayoutCachedWidth =13380
                            LayoutCachedHeight =3420
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
                    Left =9195
                    Top =3135
                    Width =1500
                    Height =285
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

                    LayoutCachedLeft =9195
                    LayoutCachedTop =3135
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =3420
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
                            Left =8385
                            Top =3135
                            Width =810
                            Height =285
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
                            LayoutCachedLeft =8385
                            LayoutCachedTop =3135
                            LayoutCachedWidth =9195
                            LayoutCachedHeight =3420
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
                    Left =9195
                    Top =3810
                    Width =1500
                    Height =270
                    FontSize =10
                    TabIndex =20
                    BackColor =15921906
                    Name ="Remittee"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =9195
                    LayoutCachedTop =3810
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =4080
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
                    Left =11445
                    Top =2820
                    Width =1095
                    Height =285
                    Name ="EmptyCell220"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11445
                    LayoutCachedTop =2820
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =3105
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =10695
                    Top =2820
                    Width =750
                    Height =285
                    Name ="EmptyCell230"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10695
                    LayoutCachedTop =2820
                    LayoutCachedWidth =11445
                    LayoutCachedHeight =3105
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =7785
                    Top =3810
                    Width =600
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="DayofMthLabel"
                    Caption ="th  Day"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =7785
                    LayoutCachedTop =3810
                    LayoutCachedWidth =8385
                    LayoutCachedHeight =4080
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
                    Left =11445
                    Top =3810
                    Width =1095
                    Height =270
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

                    LayoutCachedLeft =11445
                    LayoutCachedTop =3810
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =4080
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
                            Left =10695
                            Top =3810
                            Width =750
                            Height =270
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
                            LayoutCachedLeft =10695
                            LayoutCachedTop =3810
                            LayoutCachedWidth =11445
                            LayoutCachedHeight =4080
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
                    Left =6960
                    Top =3810
                    Width =822
                    Height =270
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

                    LayoutCachedLeft =6960
                    LayoutCachedTop =3810
                    LayoutCachedWidth =7782
                    LayoutCachedHeight =4080
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
                            Top =3810
                            Width =810
                            Height =270
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
                            LayoutCachedTop =3810
                            LayoutCachedWidth =6930
                            LayoutCachedHeight =4080
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
                    Left =8385
                    Top =3810
                    Width =810
                    Height =270
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
                    LayoutCachedLeft =8385
                    LayoutCachedTop =3810
                    LayoutCachedWidth =9195
                    LayoutCachedHeight =4080
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
                    Left =9195
                    Top =3480
                    Width =1500
                    Height =270
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

                    LayoutCachedLeft =9195
                    LayoutCachedTop =3480
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =3750
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
                            Left =8385
                            Top =3480
                            Width =810
                            Height =270
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
                            LayoutCachedLeft =8385
                            LayoutCachedTop =3480
                            LayoutCachedWidth =9195
                            LayoutCachedHeight =3750
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
                    Left =13380
                    Top =3480
                    Width =1320
                    Height =270
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

                    LayoutCachedLeft =13380
                    LayoutCachedTop =3480
                    LayoutCachedWidth =14700
                    LayoutCachedHeight =3750
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
                            Left =12540
                            Top =3480
                            Width =840
                            Height =270
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
                            LayoutCachedLeft =12540
                            LayoutCachedTop =3480
                            LayoutCachedWidth =13380
                            LayoutCachedHeight =3750
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
                    Left =11445
                    Top =4140
                    Width =1095
                    Height =270
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

                    LayoutCachedLeft =11445
                    LayoutCachedTop =4140
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =4410
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
                            Left =10695
                            Top =4140
                            Width =750
                            Height =270
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
                            LayoutCachedLeft =10695
                            LayoutCachedTop =4140
                            LayoutCachedWidth =11445
                            LayoutCachedHeight =4410
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
                    Left =13380
                    Top =4140
                    Width =1320
                    Height =270
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

                    LayoutCachedLeft =13380
                    LayoutCachedTop =4140
                    LayoutCachedWidth =14700
                    LayoutCachedHeight =4410
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
                            Left =12540
                            Top =4140
                            Width =840
                            Height =270
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
                            LayoutCachedLeft =12540
                            LayoutCachedTop =4140
                            LayoutCachedWidth =13380
                            LayoutCachedHeight =4410
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
                    Left =3630
                    Top =3495
                    Width =1320
                    Height =270
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

                    LayoutCachedLeft =3630
                    LayoutCachedTop =3495
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =3765
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
                            Left =2775
                            Top =3495
                            Width =855
                            Height =270
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
                            LayoutCachedLeft =2775
                            LayoutCachedTop =3495
                            LayoutCachedWidth =3630
                            LayoutCachedHeight =3765
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
                    Left =3630
                    Top =4155
                    Width =1320
                    Height =285
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

                    LayoutCachedLeft =3630
                    LayoutCachedTop =4155
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =4440
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
                            Left =2775
                            Top =4155
                            Width =855
                            Height =285
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
                            LayoutCachedLeft =2775
                            LayoutCachedTop =4155
                            LayoutCachedWidth =3630
                            LayoutCachedHeight =4440
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
                    Left =3630
                    Top =2820
                    Width =1320
                    Height =285
                    Name ="EmptyCell596"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =3630
                    LayoutCachedTop =2820
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =3105
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =2775
                    Top =2820
                    Width =855
                    Height =285
                    Name ="EmptyCell603"
                    GroupTable =4
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2775
                    LayoutCachedTop =2820
                    LayoutCachedWidth =3630
                    LayoutCachedHeight =3105
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =13380
                    Top =2820
                    Width =1320
                    Height =285
                    Name ="EmptyCell622"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13380
                    LayoutCachedTop =2820
                    LayoutCachedWidth =14700
                    LayoutCachedHeight =3105
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =12540
                    Top =2820
                    Width =840
                    Height =285
                    Name ="EmptyCell636"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12540
                    LayoutCachedTop =2820
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =3105
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =9195
                    Top =2820
                    Width =1500
                    Height =285
                    Name ="EmptyCell649"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9195
                    LayoutCachedTop =2820
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =3105
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =8385
                    Top =2820
                    Width =810
                    Height =285
                    Name ="EmptyCell663"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8385
                    LayoutCachedTop =2820
                    LayoutCachedWidth =9195
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
                    Left =1155
                    Top =4155
                    Width =1620
                    Height =285
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

                    LayoutCachedLeft =1155
                    LayoutCachedTop =4155
                    LayoutCachedWidth =2775
                    LayoutCachedHeight =4440
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
                            Top =4155
                            Width =795
                            Height =285
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
                            LayoutCachedTop =4155
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =4440
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
                    Left =11445
                    Top =3480
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

                    LayoutCachedLeft =11445
                    LayoutCachedTop =3480
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =3750
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
                            Left =10695
                            Top =3480
                            Width =750
                            Height =270
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
                            LayoutCachedLeft =10695
                            LayoutCachedTop =3480
                            LayoutCachedWidth =11445
                            LayoutCachedHeight =3750
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
                    Left =10695
                    Top =4140
                    Width =0
                    Height =270
                    Name ="EmptyCell736"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =10695
                    LayoutCachedTop =4140
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =4410
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
                    Left =14040
                    Top =4830
                    Width =660
                    Height =270
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

                    LayoutCachedLeft =14040
                    LayoutCachedTop =4830
                    LayoutCachedWidth =14700
                    LayoutCachedHeight =5100
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
                            Left =13335
                            Top =4830
                            Width =705
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="BillingRuleId_Label"
                            Caption ="RuleId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =15
                            RightPadding =0
                            BottomPadding =120
                            LayoutCachedLeft =13335
                            LayoutCachedTop =4830
                            LayoutCachedWidth =14040
                            LayoutCachedHeight =5100
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
                    Left =9435
                    Top =4830
                    Width =660
                    Height =270
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

                    LayoutCachedLeft =9435
                    LayoutCachedTop =4830
                    LayoutCachedWidth =10095
                    LayoutCachedHeight =5100
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
                            Left =8730
                            Top =4830
                            Width =705
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="ClientId_Label"
                            Caption ="ClientId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =15
                            RightPadding =0
                            BottomPadding =120
                            LayoutCachedLeft =8730
                            LayoutCachedTop =4830
                            LayoutCachedWidth =9435
                            LayoutCachedHeight =5100
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
                    Left =10995
                    Top =4830
                    Width =660
                    Height =270
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

                    LayoutCachedLeft =10995
                    LayoutCachedTop =4830
                    LayoutCachedWidth =11655
                    LayoutCachedHeight =5100
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
                            Left =10095
                            Top =4830
                            Width =900
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="ScheduleId_Label"
                            Caption ="ScheduleId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =15
                            RightPadding =0
                            BottomPadding =120
                            LayoutCachedLeft =10095
                            LayoutCachedTop =4830
                            LayoutCachedWidth =10995
                            LayoutCachedHeight =5100
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
                    Left =12675
                    Top =4830
                    Width =660
                    Height =270
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

                    LayoutCachedLeft =12675
                    LayoutCachedTop =4830
                    LayoutCachedWidth =13335
                    LayoutCachedHeight =5100
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
                            Left =11655
                            Top =4830
                            Width =1020
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="UnitGroupId_Label"
                            Caption ="UnitGroupId"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =15
                            RightPadding =0
                            BottomPadding =120
                            LayoutCachedLeft =11655
                            LayoutCachedTop =4830
                            LayoutCachedWidth =12675
                            LayoutCachedHeight =5100
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
                    Left =7785
                    Top =2820
                    Width =600
                    Height =285
                    Name ="EmptyCell874"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7785
                    LayoutCachedTop =2820
                    LayoutCachedWidth =8385
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
                    Left =3630
                    Top =3825
                    Width =1320
                    Height =270
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

                    LayoutCachedLeft =3630
                    LayoutCachedTop =3825
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =4095
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
                            Left =2775
                            Top =3825
                            Width =855
                            Height =270
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
                            LayoutCachedLeft =2775
                            LayoutCachedTop =3825
                            LayoutCachedWidth =3630
                            LayoutCachedHeight =4095
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
                    Left =13380
                    Top =3810
                    Width =1320
                    Height =270
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

                    LayoutCachedLeft =13380
                    LayoutCachedTop =3810
                    LayoutCachedWidth =14700
                    LayoutCachedHeight =4080
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
                            Left =12540
                            Top =3810
                            Width =840
                            Height =270
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
                            LayoutCachedLeft =12540
                            LayoutCachedTop =3810
                            LayoutCachedWidth =13380
                            LayoutCachedHeight =4080
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
                    Left =7650
                    Top =4830
                    Width =1080
                    Height =270
                    Name ="EmptyCell901"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120
                    LayoutCachedLeft =7650
                    LayoutCachedTop =4830
                    LayoutCachedWidth =8730
                    LayoutCachedHeight =5100
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =8385
                    Top =4140
                    Width =810
                    Height =270
                    Name ="EmptyCell907"
                    GroupTable =6
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8385
                    LayoutCachedTop =4140
                    LayoutCachedWidth =9195
                    LayoutCachedHeight =4410
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
                    Left =9195
                    Top =4140
                    Width =1500
                    Height =270
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

                    LayoutCachedLeft =9195
                    LayoutCachedTop =4140
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =4410
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
                    Left =6960
                    Top =4140
                    Width =1425
                    Height =270
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

                    LayoutCachedLeft =6960
                    LayoutCachedTop =4140
                    LayoutCachedWidth =8385
                    LayoutCachedHeight =4410
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
                            Top =4140
                            Width =810
                            Height =270
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
                            LayoutCachedTop =4140
                            LayoutCachedWidth =6930
                            LayoutCachedHeight =4410
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
