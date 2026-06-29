Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =204
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =1081
    Left =4665
    Top =4500
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0xf783f67156ade440
    End
    RecordSource ="GroupReviewSyndQ"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Rectangle
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
        End
        Begin ListBox
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="SGrpNO"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1080
            Name ="PageHeaderSection"
            Begin
                Begin Image
                    SizeMode =1
                    PictureType =2
                    Width =2034
                    Height =939
                    Name ="Auto_Logo0"
                    Picture ="Auto_Logo0"

                    LayoutCachedWidth =2034
                    LayoutCachedHeight =939
                    TabIndex =2
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1980
                    Top =60
                    Width =9480
                    Height =420
                    FontSize =16
                    FontWeight =700
                    Name ="Text0"
                    ControlSource ="=UCase(\" Group Review for \" & [mlno] & \" \" & [schno] & \" \" & [sgrpno])"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =60
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2460
                    Top =465
                    Width =9015
                    Height =315
                    FontSize =12
                    TabIndex =1
                    Name ="SchShortDesc"
                    ControlSource ="SchShortDesc"

                    LayoutCachedLeft =2460
                    LayoutCachedTop =465
                    LayoutCachedWidth =11475
                    LayoutCachedHeight =780
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =6540
            Name ="GroupHeader0"
            Begin
                Begin Rectangle
                    Left =360
                    Top =1380
                    Width =10800
                    Height =360
                    Name ="Box430"
                    LayoutCachedLeft =360
                    LayoutCachedTop =1380
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =1740
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1152
                    Top =4176
                    Width =1080
                    Height =228
                    FontSize =10
                    FontWeight =700
                    TabIndex =52
                    Name ="BLCD\015\012\015\012\015\012\015\012"
                    ControlSource ="=[BLCD]"
                    Format ="Short Date"
                    EventProcPrefix ="BLCD________"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1152
                    LayoutCachedTop =4176
                    LayoutCachedWidth =2232
                    LayoutCachedHeight =4404
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =360
                            Top =4176
                            Width =795
                            Height =228
                            FontSize =10
                            Name ="Label23"
                            Caption ="BLCD"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =360
                            LayoutCachedTop =4176
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =4404
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =7
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =360
                    Top =2244
                    Width =1884
                    Height =348
                    FontSize =12
                    FontWeight =700
                    TabIndex =21
                    Name ="Assignee_Bank\015\012\015\012\015\012\015\012"
                    ControlSource ="=[Assignee_Bank]"
                    StatusBarText ="select from Bank file"
                    EventProcPrefix ="Assignee_Bank________"
                    GroupTable =64
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =360
                    LayoutCachedTop =2244
                    LayoutCachedWidth =2244
                    LayoutCachedHeight =2592
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =64
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =360
                    Top =2604
                    Width =487
                    Height =432
                    FontSize =12
                    FontWeight =700
                    TabIndex =22
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="Assigned_Indicator\015\012\015\012\015\012\015\012"
                    ControlSource ="=[Assigned_Indicator]"
                    EventProcPrefix ="Assigned_Indicator________"
                    GroupTable =64
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =360
                    LayoutCachedTop =2604
                    LayoutCachedWidth =847
                    LayoutCachedHeight =3036
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =64
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1152
                    Top =3948
                    Width =1080
                    Height =228
                    FontSize =10
                    FontWeight =700
                    TabIndex =51
                    Name ="AssignmentDate\015\012"
                    ControlSource ="AssignmentDate"
                    Format ="Short Date"
                    StatusBarText ="bank assignment date"
                    EventProcPrefix ="AssignmentDate__"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1152
                    LayoutCachedTop =3948
                    LayoutCachedWidth =2232
                    LayoutCachedHeight =4176
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =360
                            Top =3948
                            Width =795
                            Height =228
                            FontSize =10
                            Name ="Label25"
                            Caption ="Assigned"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =360
                            LayoutCachedTop =3948
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =4176
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =7
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1152
                    Top =3720
                    Width =1080
                    Height =225
                    FontSize =10
                    FontWeight =700
                    TabIndex =50
                    Name ="FundDate"
                    ControlSource ="[FundDate]"
                    Format ="Short Date"
                    StatusBarText ="FundDate"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1152
                    LayoutCachedTop =3720
                    LayoutCachedWidth =2232
                    LayoutCachedHeight =3945
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =360
                            Top =3720
                            Width =795
                            Height =225
                            FontSize =10
                            Name ="Label31"
                            Caption ="Fund"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =360
                            LayoutCachedTop =3720
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =3945
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =7
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9012
                    Top =1980
                    Width =1452
                    Height =285
                    TabIndex =23
                    Name ="SalePrice\015\012\015\012\015\012"
                    ControlSource ="[SalePrice]"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="sale price to bank"
                    EventProcPrefix ="SalePrice______"
                    GroupTable =36

                    LayoutCachedLeft =9012
                    LayoutCachedTop =1980
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =2265
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =7500
                            Top =1980
                            Width =1470
                            Height =285
                            FontSize =10
                            Name ="Label32"
                            Caption ="Sale Price"
                            GroupTable =36
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =7500
                            LayoutCachedTop =1980
                            LayoutCachedWidth =8970
                            LayoutCachedHeight =2265
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =5
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =36
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =360
                    Top =1932
                    Width =1884
                    Height =300
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label24"
                    Caption ="Assignee"
                    GroupTable =64
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =360
                    LayoutCachedTop =1932
                    LayoutCachedWidth =2244
                    LayoutCachedHeight =2232
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =64
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9012
                    Top =3372
                    Width =1452
                    Height =288
                    FontWeight =700
                    TabIndex =31
                    Name ="NetFee"
                    ControlSource ="=CCur((Nz([SalePrice],0)-Nz([OrigEquipCost],0))+[StripsAmount])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001010000 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =36

                    LayoutCachedLeft =9012
                    LayoutCachedTop =3372
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =3660
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001010000ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =2460
                            Top =3372
                            Width =6510
                            Height =288
                            FontSize =10
                            Name ="Label65"
                            Caption ="Gross Bank Fee + Strips"
                            GroupTable =36
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2460
                            LayoutCachedTop =3372
                            LayoutCachedWidth =8970
                            LayoutCachedHeight =3660
                            RowStart =4
                            RowEnd =4
                            ColumnEnd =1
                            LayoutGroup =5
                            BorderThemeColorIndex =0
                            BorderTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =36
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9012
                    Top =3024
                    Width =1452
                    Height =288
                    TabIndex =29
                    Name ="StripsAmount"
                    ControlSource ="StripsAmount"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =36

                    LayoutCachedLeft =9012
                    LayoutCachedTop =3024
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =3312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =7500
                            Top =3024
                            Width =1470
                            Height =288
                            FontSize =10
                            Name ="Label67"
                            Caption ="+ Strips Amount"
                            GroupTable =36
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =7500
                            LayoutCachedTop =3024
                            LayoutCachedWidth =8970
                            LayoutCachedHeight =3312
                            RowStart =3
                            RowEnd =3
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =5
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =36
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9012
                    Top =4392
                    Width =1452
                    Height =312
                    TabIndex =38
                    Name ="UTALine"
                    ControlSource ="=[SgrpSynRebate]*[UnitsInGroup]"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =36
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9012
                    LayoutCachedTop =4392
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =4704
                    RowStart =7
                    RowEnd =7
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =7500
                            Top =4392
                            Width =1470
                            Height =312
                            FontSize =10
                            Name ="Label70"
                            Caption ="+ Actual UTA"
                            GroupTable =36
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =7500
                            LayoutCachedTop =4392
                            LayoutCachedWidth =8970
                            LayoutCachedHeight =4704
                            RowStart =7
                            RowEnd =7
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =5
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =36
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =238
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9012
                    Top =5676
                    Width =1452
                    Height =312
                    FontWeight =700
                    TabIndex =47
                    Name ="AllInLine"
                    ControlSource ="=[NetBankFee]+[UTALine]+[PerDiem]+[ServFeePVLine]-[FundingExpense]"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001010000 ,
                        0xff000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =36

                    LayoutCachedLeft =9012
                    LayoutCachedTop =5676
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =5988
                    RowStart =11
                    RowEnd =11
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    BorderThemeColorIndex =2
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001010000ff000000ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =7500
                            Top =5676
                            Width =1470
                            Height =312
                            FontSize =10
                            FontWeight =700
                            Name ="Label91"
                            Caption ="All In Fee"
                            GroupTable =36
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =7500
                            LayoutCachedTop =5676
                            LayoutCachedWidth =8970
                            LayoutCachedHeight =5988
                            RowStart =11
                            RowEnd =11
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =5
                            BorderThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =36
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10524
                    Top =5676
                    Width =780
                    Height =312
                    FontSize =10
                    TabIndex =48
                    Name ="AllinPct"
                    ControlSource ="=[AllInLine]/[OrigEquipCost]"
                    Format ="Percent"
                    GroupTable =36

                    LayoutCachedLeft =10524
                    LayoutCachedTop =5676
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =5988
                    RowStart =11
                    RowEnd =11
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10524
                    Top =4068
                    Width =780
                    Height =288
                    FontSize =10
                    TabIndex =36
                    Name ="Text96"
                    ControlSource ="=([SalePrice]-[OrigEquipCost]+Nz([StripsAmount],0)+Nz([SyndPenalty],0))/[OrigEqu"
                        "ipCost]"
                    Format ="Percent"
                    GroupTable =36

                    LayoutCachedLeft =10524
                    LayoutCachedTop =4068
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =4356
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6852
                    Top =564
                    Width =876
                    Height =312
                    TabIndex =6
                    Name ="SGrpBonus"
                    ControlSource ="SGrpBonus"
                    StatusBarText ="Bonus Depreciation"
                    GroupTable =9

                    LayoutCachedLeft =6852
                    LayoutCachedTop =564
                    LayoutCachedWidth =7728
                    LayoutCachedHeight =876
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1380
                    Top =564
                    Width =840
                    Height =312
                    FontWeight =700
                    Name ="SGrpIndex"
                    ControlSource ="SGrpIndex"
                    GroupTable =9

                    LayoutCachedLeft =1380
                    LayoutCachedTop =564
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =876
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                    Begin
                        Begin Label
                            BackStyle =1
                            TextAlign =2
                            Left =360
                            Top =564
                            Width =945
                            Height =312
                            ForeColor =16777215
                            Name ="Label86"
                            Caption ="Synd"
                            GroupTable =9
                            LayoutCachedLeft =360
                            LayoutCachedTop =564
                            LayoutCachedWidth =1305
                            LayoutCachedHeight =876
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            BackThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =9
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2292
                    Top =564
                    Width =1092
                    Height =312
                    FontWeight =700
                    TabIndex =1
                    Name ="SGrpIndDt"
                    ControlSource ="SGrpIndDt"
                    Format ="mm/dd/yy"
                    GroupTable =9
                    ShowDatePicker =1

                    LayoutCachedLeft =2292
                    LayoutCachedTop =564
                    LayoutCachedWidth =3384
                    LayoutCachedHeight =876
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3456
                    Top =564
                    Width =612
                    Height =312
                    FontWeight =700
                    TabIndex =2
                    Name ="SGprIndYr"
                    ControlSource ="SGprIndYr"
                    GroupTable =9

                    LayoutCachedLeft =3456
                    LayoutCachedTop =564
                    LayoutCachedWidth =4068
                    LayoutCachedHeight =876
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5952
                    Top =564
                    Width =828
                    Height =312
                    FontWeight =700
                    TabIndex =5
                    Name ="SgrpYield"
                    ControlSource ="SgrpYield"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =5952
                    LayoutCachedTop =564
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =876
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =6852
                    Top =180
                    Width =876
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label33"
                    Caption ="Bonus"
                    GroupTable =9
                    LayoutCachedLeft =6852
                    LayoutCachedTop =180
                    LayoutCachedWidth =7728
                    LayoutCachedHeight =495
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =1380
                    Top =180
                    Width =840
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label34"
                    Caption ="Index"
                    GroupTable =9
                    LayoutCachedLeft =1380
                    LayoutCachedTop =180
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =495
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =2292
                    Top =180
                    Width =1092
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label35"
                    Caption ="Date"
                    GroupTable =9
                    LayoutCachedLeft =2292
                    LayoutCachedTop =180
                    LayoutCachedWidth =3384
                    LayoutCachedHeight =495
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =3456
                    Top =180
                    Width =612
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label36"
                    Caption ="Year"
                    GroupTable =9
                    LayoutCachedLeft =3456
                    LayoutCachedTop =180
                    LayoutCachedWidth =4068
                    LayoutCachedHeight =495
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =5952
                    Top =180
                    Width =828
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label37"
                    Caption ="Yield"
                    GroupTable =9
                    LayoutCachedLeft =5952
                    LayoutCachedTop =180
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =495
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4968
                    Top =564
                    Width =912
                    Height =312
                    FontWeight =700
                    TabIndex =4
                    Name ="Text54"
                    ControlSource ="=[sgrpyield]-[IndexRate]"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =4968
                    LayoutCachedTop =564
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =876
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4140
                    Top =564
                    Width =756
                    Height =312
                    FontWeight =700
                    TabIndex =3
                    Name ="IndexRate"
                    ControlSource ="IndexRate"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =4140
                    LayoutCachedTop =564
                    LayoutCachedWidth =4896
                    LayoutCachedHeight =876
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6852
                    Top =948
                    Width =876
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =16
                    ForeColor =9211020
                    Name ="SGrpSalesBonus"
                    ControlSource ="SGrpSalesBonus"
                    StatusBarText ="Bonus Depreciation"
                    GroupTable =9

                    LayoutCachedLeft =6852
                    LayoutCachedTop =948
                    LayoutCachedWidth =7728
                    LayoutCachedHeight =1224
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1380
                    Top =948
                    Width =840
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =10
                    ForeColor =9211020
                    Name ="SGrpSalesIndex"
                    ControlSource ="SGrpSalesIndex"
                    GroupTable =9

                    LayoutCachedLeft =1380
                    LayoutCachedTop =948
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =1224
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OldBorderStyle =1
                            TextAlign =2
                            TextFontFamily =34
                            Left =360
                            Top =948
                            Width =945
                            Height =276
                            ForeColor =9211020
                            Name ="Label87"
                            Caption ="LO Target"
                            GroupTable =9
                            LayoutCachedLeft =360
                            LayoutCachedTop =948
                            LayoutCachedWidth =1305
                            LayoutCachedHeight =1224
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            BorderThemeColorIndex =0
                            BorderTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =9
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2292
                    Top =948
                    Width =1092
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =11
                    ForeColor =9211020
                    Name ="SGrpSalesIndDt"
                    ControlSource ="SGrpSalesIndDt"
                    Format ="mm/dd/yy"
                    GroupTable =9
                    ShowDatePicker =1

                    LayoutCachedLeft =2292
                    LayoutCachedTop =948
                    LayoutCachedWidth =3384
                    LayoutCachedHeight =1224
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3456
                    Top =948
                    Width =612
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =12
                    ForeColor =9211020
                    Name ="SGprSalesIndYr"
                    ControlSource ="SGprSalesIndYr"
                    GroupTable =9

                    LayoutCachedLeft =3456
                    LayoutCachedTop =948
                    LayoutCachedWidth =4068
                    LayoutCachedHeight =1224
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4140
                    Top =948
                    Width =756
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =13
                    ForeColor =9211020
                    Name ="SgrpSalesIndRt"
                    ControlSource ="SgrpSalesIndRt"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =4140
                    LayoutCachedTop =948
                    LayoutCachedWidth =4896
                    LayoutCachedHeight =1224
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5952
                    Top =948
                    Width =828
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =15
                    ForeColor =9211020
                    Name ="SgrpSalesYield"
                    ControlSource ="SgrpSalesYield"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =5952
                    LayoutCachedTop =948
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =1224
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4968
                    Top =948
                    Width =912
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =14
                    ForeColor =9211020
                    Name ="Text84"
                    ControlSource ="=[SgrpSalesYield]-[SgrpSalesIndRt]"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =4968
                    LayoutCachedTop =948
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =1224
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1152
                    Top =3492
                    Width =1080
                    Height =228
                    FontSize =10
                    FontWeight =700
                    TabIndex =49
                    Name ="SaleDate\015\012"
                    ControlSource ="[SaleDate]"
                    Format ="Short Date"
                    StatusBarText ="date sold to assignee"
                    EventProcPrefix ="SaleDate__"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1152
                    LayoutCachedTop =3492
                    LayoutCachedWidth =2232
                    LayoutCachedHeight =3720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =360
                            Top =3492
                            Width =795
                            Height =228
                            FontSize =10
                            Name ="Label109"
                            Caption ="Sale"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =360
                            LayoutCachedTop =3492
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =3720
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =7
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9012
                    Top =3720
                    Width =1452
                    Height =288
                    TabIndex =33
                    ForeColor =2366701
                    Name ="SynPenalty"
                    ControlSource ="=CDbl(Nz([SgrpSynPenalty],0))"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =36

                    LayoutCachedLeft =9012
                    LayoutCachedTop =3720
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =4008
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =2460
                            Top =3720
                            Width =6510
                            Height =288
                            FontSize =10
                            Name ="Label113"
                            Caption ="- Termination Penalty"
                            GroupTable =36
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2460
                            LayoutCachedTop =3720
                            LayoutCachedWidth =8970
                            LayoutCachedHeight =4008
                            RowStart =5
                            RowEnd =5
                            ColumnEnd =1
                            LayoutGroup =5
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =36
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9012
                    Top =2676
                    Width =1452
                    Height =288
                    TabIndex =26
                    Name ="GrossFee"
                    ControlSource ="=[SalePrice]-[OrigEquipCost]"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =36

                    LayoutCachedLeft =9012
                    LayoutCachedTop =2676
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =2964
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =7500
                            Top =2676
                            Width =1470
                            Height =288
                            FontSize =10
                            Name ="Label119"
                            Caption ="Gross Bank Fee"
                            GroupTable =36
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =7500
                            LayoutCachedTop =2676
                            LayoutCachedWidth =8970
                            LayoutCachedHeight =2964
                            RowStart =2
                            RowEnd =2
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =5
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =36
                        End
                    End
                End
                Begin EmptyCell
                    Left =360
                    Top =180
                    Width =945
                    Height =315
                    Name ="EmptyCell163"
                    GroupTable =9
                    LayoutCachedLeft =360
                    LayoutCachedTop =180
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =495
                    LayoutGroup =1
                    GroupTable =9
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =4140
                    Top =180
                    Width =756
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label56"
                    Caption ="Rate"
                    GroupTable =9
                    LayoutCachedLeft =4140
                    LayoutCachedTop =180
                    LayoutCachedWidth =4896
                    LayoutCachedHeight =495
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =4968
                    Top =180
                    Width =912
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label55"
                    Caption ="Spread"
                    GroupTable =9
                    LayoutCachedLeft =4968
                    LayoutCachedTop =180
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =495
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    TextFontFamily =34
                    Left =864
                    Top =2604
                    Width =1380
                    Height =432
                    FontSize =8
                    FontWeight =700
                    Name ="Label419"
                    Caption =" P = Projected\015\012 F = Funded"
                    GroupTable =64
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =864
                    LayoutCachedTop =2604
                    LayoutCachedWidth =2244
                    LayoutCachedHeight =3036
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =64
                End
                Begin Line
                    Top =6360
                    Width =11355
                    Name ="Line108"
                    LayoutCachedTop =6360
                    LayoutCachedWidth =11355
                    LayoutCachedHeight =6360
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =168
                    Top =6180
                    Width =2160
                    Height =345
                    FontSize =12
                    FontWeight =700
                    Name ="Label477"
                    Caption ="Unit Payment Detail"
                    GroupTable =81
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =168
                    LayoutCachedTop =6180
                    LayoutCachedWidth =2328
                    LayoutCachedHeight =6525
                    ColumnEnd =1
                    LayoutGroup =4
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =81
                End
                Begin EmptyCell
                    Left =10524
                    Top =2328
                    Width =780
                    Height =288
                    Name ="EmptyCell484"
                    GroupTable =36
                    LayoutCachedLeft =10524
                    LayoutCachedTop =2328
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =2616
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    GroupTable =36
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9012
                    Top =2328
                    Width =1452
                    Height =288
                    TabIndex =24
                    Name ="OrigEquipCost"
                    ControlSource ="OrigEquipCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =36

                    LayoutCachedLeft =9012
                    LayoutCachedTop =2328
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =2616
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =7500
                            Top =2328
                            Width =1470
                            Height =288
                            FontSize =10
                            Name ="Label486"
                            Caption =" - Cost (OEC)"
                            GroupTable =36
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =7500
                            LayoutCachedTop =2328
                            LayoutCachedWidth =8970
                            LayoutCachedHeight =2616
                            RowStart =1
                            RowEnd =1
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =5
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =36
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9012
                    Top =4068
                    Width =1452
                    Height =288
                    FontWeight =700
                    TabIndex =35
                    Name ="NetBankFee"
                    ControlSource ="=CCur((Nz([SalePrice],0)-Nz([OrigEquipCost],0))+Nz([StripsAmount],0)+Nz([SyndPen"
                        "alty],0))"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001010000 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =36

                    LayoutCachedLeft =9012
                    LayoutCachedTop =4068
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =4356
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001010000ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =7500
                            Top =4068
                            Width =1470
                            Height =288
                            FontSize =10
                            Name ="Label602"
                            Caption ="Net Bank Fee"
                            GroupTable =36
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =7500
                            LayoutCachedTop =4068
                            LayoutCachedWidth =8970
                            LayoutCachedHeight =4356
                            RowStart =6
                            RowEnd =6
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =5
                            BorderThemeColorIndex =0
                            BorderTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =36
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10524
                    Top =3372
                    Width =780
                    Height =288
                    FontSize =10
                    TabIndex =32
                    Name ="Text625"
                    ControlSource ="=([SalePrice]-[OrigEquipCost]+Nz([StripsAmount],0))/[OrigEquipCost]"
                    Format ="Percent"
                    GroupTable =36

                    LayoutCachedLeft =10524
                    LayoutCachedTop =3372
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =3660
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7800
                    Top =564
                    Width =1200
                    Height =312
                    FontWeight =700
                    TabIndex =7
                    Name ="Text627"
                    ControlSource ="UnitSynResid"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =9

                    LayoutCachedLeft =7800
                    LayoutCachedTop =564
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =876
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7800
                    Top =948
                    Width =1200
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =17
                    ForeColor =9211020
                    Name ="Text628"
                    ControlSource ="unitsalesresid"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =9

                    LayoutCachedLeft =7800
                    LayoutCachedTop =948
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =1224
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                    CurrencySymbol ="$"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =7800
                    Top =180
                    Width =1200
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label629"
                    Caption ="Residual"
                    GroupTable =9
                    LayoutCachedLeft =7800
                    LayoutCachedTop =180
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =495
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9072
                    Top =564
                    Width =1392
                    Height =312
                    FontWeight =700
                    TabIndex =8
                    Name ="Text647"
                    ControlSource ="=CCur((Nz([SalePrice],0)-Nz([OrigEquipCost],0))+[StripsAmount])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =9

                    LayoutCachedLeft =9072
                    LayoutCachedTop =564
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =876
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9072
                    Top =948
                    Width =1392
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =18
                    ForeColor =9211020
                    Name ="Text648"
                    ControlSource ="[Fee_Projected]"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =9

                    LayoutCachedLeft =9072
                    LayoutCachedTop =948
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =1224
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                    CurrencySymbol ="$"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =9072
                    Top =180
                    Width =1392
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label649"
                    Caption ="Fee"
                    GroupTable =9
                    LayoutCachedLeft =9072
                    LayoutCachedTop =180
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =495
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2460
                    Top =3024
                    Width =5040
                    Height =288
                    FontSize =10
                    TabIndex =28
                    Name ="Label612"
                    ControlSource ="=IIf([StripsAmount]>0,\"(Assignee Term @ <B>\" & [LeaseTermPrimary]-[StripsNum] "
                        "& \"</b> mths, <B>\" & [StripsNum] & \"</b> Strips @ <b>\" & Format(([SalesRent]"
                        "*[unitsingroup]),\"$#,###.00\") & \" </b>Ea)\",\"\")"
                    GroupTable =36
                    LeftPadding =0
                    RightPadding =0
                    TextFormat =1

                    LayoutCachedLeft =2460
                    LayoutCachedTop =3024
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =3312
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =5
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2460
                    Top =2676
                    Width =5040
                    Height =288
                    FontSize =10
                    TabIndex =25
                    Name ="Label124"
                    ControlSource ="=IIf([WithHeldPymts]>0,\"<B>\" & [WithHeldPymts] & \"</b> Withheld Payments  (<B"
                        ">\" & Format([LeasePymtForGroup]*[WithHeldPymts],\"$#,000\") & \"</b>)\",\"\")"
                    GroupTable =36
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =2460
                    LayoutCachedTop =2676
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =2964
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =5
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
                Begin Line
                    BorderWidth =1
                    Left =9000
                    Top =4032
                    Width =1380
                    Name ="Line665"
                    LayoutCachedLeft =9000
                    LayoutCachedTop =4032
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =4032
                    BorderThemeColorIndex =0
                End
                Begin Line
                    BorderWidth =1
                    Left =9000
                    Top =2640
                    Width =1380
                    Name ="Line674"
                    LayoutCachedLeft =9000
                    LayoutCachedTop =2640
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =2640
                    BorderThemeColorIndex =0
                End
                Begin Line
                    BorderWidth =1
                    Left =9000
                    Top =5645
                    Width =1380
                    Name ="Line675"
                    LayoutCachedLeft =9000
                    LayoutCachedTop =5645
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =5645
                    BorderThemeColorIndex =0
                End
                Begin Line
                    BorderWidth =1
                    Left =9000
                    Top =3300
                    Width =1380
                    Name ="Line676"
                    LayoutCachedLeft =9000
                    LayoutCachedTop =3300
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =3300
                    BorderThemeColorIndex =0
                End
                Begin EmptyCell
                    Left =10524
                    Top =1980
                    Width =780
                    Height =285
                    Name ="EmptyCell681"
                    GroupTable =36
                    LayoutCachedLeft =10524
                    LayoutCachedTop =1980
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =2265
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    GroupTable =36
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9012
                    Top =4704
                    Width =1452
                    Height =312
                    TabIndex =40
                    Name ="PerDiem"
                    ControlSource ="PerDiemGroup"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =36
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9012
                    LayoutCachedTop =4704
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =5016
                    RowStart =8
                    RowEnd =8
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =7500
                            Top =4704
                            Width =1470
                            Height =312
                            FontSize =10
                            Name ="Label682"
                            Caption ="+ Actual Perdiem"
                            GroupTable =36
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =7500
                            LayoutCachedTop =4704
                            LayoutCachedWidth =8970
                            LayoutCachedHeight =5016
                            RowStart =8
                            RowEnd =8
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =5
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =36
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2460
                    Top =4392
                    Width =5040
                    Height =312
                    FontSize =10
                    TabIndex =37
                    Name ="Text693"
                    ControlSource ="=IIf([SgrpSynRebate]>0,\"(<B>\" & [UnitsInGroup] & \"</b> Units @ <b>\" & Format"
                        "([SgrpSynRebate],\"$#,###.00\") & \" </b>Ea)\",\"\")"
                    GroupTable =36
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =2460
                    LayoutCachedTop =4392
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =4704
                    RowStart =7
                    RowEnd =7
                    LayoutGroup =5
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =420
                    Top =1425
                    Width =10680
                    Height =285
                    TabIndex =20
                    Name ="Text712"
                    ControlSource ="=\"Primary Term of: <B>\" & [LeaseTermPrimary] & \"</b> months with payments @ <"
                        "B> \" & Format([LeasePymtForGroup],\"$#,###.00\") & \"</b> each to be paid <B>\""
                        " & [PayFreq] & \"</b> in <B>\" & [AdvanceArrears] & \"</b>\""
                    TextFormat =1

                    LayoutCachedLeft =420
                    LayoutCachedTop =1425
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =1710
                    ColumnEnd =3
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =360
                    Top =3192
                    Width =1872
                    Height =300
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label749"
                    Caption ="Dates"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =360
                    LayoutCachedTop =3192
                    LayoutCachedWidth =2232
                    LayoutCachedHeight =3492
                    ColumnEnd =1
                    LayoutGroup =7
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =2460
                    Top =1980
                    Width =5040
                    Height =285
                    Name ="EmptyCell940"
                    GroupTable =36
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2460
                    LayoutCachedTop =1980
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =2265
                    LayoutGroup =5
                    GroupTable =36
                End
                Begin EmptyCell
                    Left =2460
                    Top =2328
                    Width =5040
                    Height =288
                    Name ="EmptyCell941"
                    GroupTable =36
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2460
                    LayoutCachedTop =2328
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =2616
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =5
                    GroupTable =36
                End
                Begin EmptyCell
                    Left =2460
                    Top =4068
                    Width =5040
                    Height =288
                    Name ="EmptyCell946"
                    GroupTable =36
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2460
                    LayoutCachedTop =4068
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =4356
                    RowStart =6
                    RowEnd =6
                    LayoutGroup =5
                    GroupTable =36
                End
                Begin EmptyCell
                    Left =2460
                    Top =5676
                    Width =5040
                    Height =312
                    Name ="EmptyCell949"
                    GroupTable =36
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2460
                    LayoutCachedTop =5676
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =5988
                    RowStart =11
                    RowEnd =11
                    LayoutGroup =5
                    GroupTable =36
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10524
                    Top =3720
                    Width =780
                    Height =288
                    FontSize =10
                    TabIndex =34
                    ForeColor =1643706
                    Name ="Text953"
                    ControlSource ="=IIf([SyndPenalty]<>0,[SyndPenalty]/[OrigEquipCost],Null)"
                    Format ="Percent"
                    GroupTable =36

                    LayoutCachedLeft =10524
                    LayoutCachedTop =3720
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =4008
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10524
                    Top =3024
                    Width =780
                    Height =288
                    FontSize =10
                    TabIndex =30
                    Name ="Text955"
                    ControlSource ="=IIf([StripsAmount]<>0,[StripsAmount]/[OrigEquipCost],Null)"
                    Format ="Percent"
                    GroupTable =36

                    LayoutCachedLeft =10524
                    LayoutCachedTop =3024
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =3312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10524
                    Top =4704
                    Width =780
                    Height =312
                    FontSize =10
                    TabIndex =41
                    Name ="Text957"
                    ControlSource ="=IIf([PerDiemGroup]<>0,[PerDiemGroup]/[OrigEquipCost],Null)"
                    Format ="Percent"
                    GroupTable =36
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10524
                    LayoutCachedTop =4704
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =5016
                    RowStart =8
                    RowEnd =8
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10524
                    Top =4392
                    Width =780
                    Height =312
                    FontSize =10
                    TabIndex =39
                    Name ="Text959"
                    ControlSource ="=IIf([SgrpSynRebate]*[UnitsInGroup]>0,([SgrpSynRebate]*[UnitsInGroup])/[OrigEqui"
                        "pCost],Null)"
                    Format ="Percent"
                    GroupTable =36
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10524
                    LayoutCachedTop =4392
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =4704
                    RowStart =7
                    RowEnd =7
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10524
                    Top =2676
                    Width =780
                    Height =288
                    FontSize =10
                    TabIndex =27
                    Name ="Text961"
                    ControlSource ="=IIf([SalePrice]-[OrigEquipCost]<>0,([SalePrice]-[OrigEquipCost])/[OrigEquipCost"
                        "],Null)"
                    Format ="Percent"
                    GroupTable =36

                    LayoutCachedLeft =10524
                    LayoutCachedTop =2676
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =2964
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10536
                    Top =564
                    Width =768
                    Height =312
                    FontWeight =700
                    TabIndex =9
                    Name ="Text964"
                    ControlSource ="=([SalePrice]-[OrigEquipCost]+Nz([StripsAmount],0))/[OrigEquipCost]"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =10536
                    LayoutCachedTop =564
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =876
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =10536
                    Top =180
                    Width =768
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label965"
                    Caption ="Fee%"
                    GroupTable =9
                    LayoutCachedLeft =10536
                    LayoutCachedTop =180
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =495
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10536
                    Top =948
                    Width =768
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =19
                    ForeColor =9211020
                    Name ="Text966"
                    ControlSource ="=[Fee_Projected]/[OrigEquipCost]"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =10536
                    LayoutCachedTop =948
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =1224
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9012
                    Top =5016
                    Width =1452
                    Height =312
                    TabIndex =43
                    Name ="ServFeePVLine"
                    ControlSource ="ServFeePV"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =36
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9012
                    LayoutCachedTop =5016
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =5328
                    RowStart =9
                    RowEnd =9
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =7500
                            Top =5016
                            Width =1470
                            Height =312
                            FontSize =10
                            Name ="Label977"
                            Caption ="+ Service Fee (PV)"
                            GroupTable =36
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =7500
                            LayoutCachedTop =5016
                            LayoutCachedWidth =8970
                            LayoutCachedHeight =5328
                            RowStart =9
                            RowEnd =9
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =5
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =36
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10524
                    Top =5016
                    Width =780
                    Height =312
                    FontSize =10
                    TabIndex =44
                    Name ="Text981"
                    ControlSource ="=IIf([ServFeePV]>0,[ServFeePV]/[OrigEquipCost],Null)"
                    Format ="Percent"
                    GroupTable =36
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10524
                    LayoutCachedTop =5016
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =5328
                    RowStart =9
                    RowEnd =9
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2460
                    Top =5016
                    Width =5040
                    Height =312
                    FontSize =10
                    TabIndex =42
                    Name ="Text983"
                    ControlSource ="=IIf([FullServiceFee]>0,\"(PV Rate: <B>\" & Format(Nz([PVRate_Default_5pct],0.05"
                        "),\"Percent\") & \"</b> applied on <b>\" & Format([FullServiceFee],\"$#,###.00\""
                        ") & \"</b>)\",\"\")"
                    GroupTable =36
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =2460
                    LayoutCachedTop =5016
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =5328
                    RowStart =9
                    RowEnd =9
                    LayoutGroup =5
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
                Begin Rectangle
                    BackStyle =0
                    Left =360
                    Top =3495
                    Width =1875
                    Height =960
                    Name ="Box1022"
                    LayoutCachedLeft =360
                    LayoutCachedTop =3495
                    LayoutCachedWidth =2235
                    LayoutCachedHeight =4455
                    BorderThemeColorIndex =0
                End
                Begin EmptyCell
                    Left =2460
                    Top =4704
                    Width =5040
                    Height =312
                    Name ="EmptyCell1026"
                    GroupTable =36
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2460
                    LayoutCachedTop =4704
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =5016
                    RowStart =8
                    RowEnd =8
                    LayoutGroup =5
                    GroupTable =36
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9012
                    Top =5328
                    Width =1452
                    Height =312
                    TabIndex =45
                    ForeColor =2366701
                    Name ="FundingExpense"
                    ControlSource ="=CDbl(Nz([IntrimFundingExpense],0))"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =36
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9012
                    LayoutCachedTop =5328
                    LayoutCachedWidth =10464
                    LayoutCachedHeight =5640
                    RowStart =10
                    RowEnd =10
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =2460
                            Top =5328
                            Width =6510
                            Height =312
                            FontSize =10
                            Name ="Label1042"
                            Caption ="- Est. Intrim Funding Expense"
                            GroupTable =36
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =2460
                            LayoutCachedTop =5328
                            LayoutCachedWidth =8970
                            LayoutCachedHeight =5640
                            RowStart =10
                            RowEnd =10
                            ColumnEnd =1
                            LayoutGroup =5
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =36
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10524
                    Top =5328
                    Width =780
                    Height =312
                    FontSize =10
                    TabIndex =46
                    ForeColor =1643706
                    Name ="Text1052"
                    ControlSource ="=IIf([FundingExpense]<>0,[FundingExpense]/[OrigEquipCost],Null)"
                    Format ="Percent"
                    GroupTable =36
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10524
                    LayoutCachedTop =5328
                    LayoutCachedWidth =11304
                    LayoutCachedHeight =5640
                    RowStart =10
                    RowEnd =10
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =36
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =315
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2328
                    Width =1296
                    Height =315
                    FontWeight =700
                    TabIndex =1
                    Name ="AvgOfUNITRENT"
                    ControlSource ="RENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =70
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2328
                    LayoutCachedWidth =3624
                    LayoutCachedHeight =315
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =70
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =1680
                            Width =570
                            Height =315
                            Name ="Label103"
                            Caption ="Rent"
                            GroupTable =70
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1680
                            LayoutCachedWidth =2250
                            LayoutCachedHeight =315
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =70
                        End
                    End
                End
                Begin TextBox
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4140
                    Width =960
                    Height =315
                    FontWeight =700
                    TabIndex =2
                    Name ="AvgOfUnitTax"
                    ControlSource ="=Nz([UnitTax],0)"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =70
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4140
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =315
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =70
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =1
                            TextFontFamily =34
                            Left =3684
                            Width =396
                            Height =315
                            Name ="Label106"
                            Caption ="Tax"
                            GroupTable =70
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =3684
                            LayoutCachedWidth =4080
                            LayoutCachedHeight =315
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =70
                        End
                    End
                End
                Begin TextBox
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6096
                    Width =996
                    Height =315
                    FontWeight =700
                    TabIndex =3
                    Name ="fees"
                    ControlSource ="ServiceFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="per unit FSL pmts"
                    GroupTable =70
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6096
                    LayoutCachedWidth =7092
                    LayoutCachedHeight =315
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =70
                    Begin
                        Begin Label
                            TextAlign =1
                            TextFontFamily =34
                            Left =5160
                            Width =876
                            Height =315
                            Name ="Label107"
                            Caption ="Serv Fee"
                            GroupTable =70
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =5160
                            LayoutCachedWidth =6036
                            LayoutCachedHeight =315
                            ColumnStart =5
                            ColumnEnd =5
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =70
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =480
                    Width =1140
                    Height =315
                    FontWeight =700
                    Name ="Text100"
                    ControlSource ="=[countOfunitvin] & \" Units\""
                    GroupTable =70
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =480
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =315
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =70
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9084
                    Height =315
                    FontWeight =700
                    TabIndex =4
                    Name ="Text709"
                    ControlSource ="=[Rent]+[ServiceFee]+Nz([UnitTax],0)"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =70
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9084
                    LayoutCachedWidth =10524
                    LayoutCachedHeight =315
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =3
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =70
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =7152
                            Width =1860
                            Height =315
                            Name ="Label698"
                            Caption ="Unit Payment"
                            GroupTable =70
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =7152
                            LayoutCachedWidth =9012
                            LayoutCachedHeight =315
                            ColumnStart =7
                            ColumnEnd =7
                            LayoutGroup =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =70
                        End
                    End
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2040
            Name ="GroupFooter2"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =9660
                    Top =1500
                    Width =1440
                    Height =315
                    FontWeight =700
                    Name ="Label60"
                    Caption ="Date      "
                    GroupTable =95
                    LayoutCachedLeft =9660
                    LayoutCachedTop =1500
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =1815
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =95
                End
                Begin Line
                    Left =4815
                    Top =1440
                    Width =6315
                    Name ="Line61"
                    LayoutCachedLeft =4815
                    LayoutCachedTop =1440
                    LayoutCachedWidth =11130
                    LayoutCachedHeight =1440
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2310
                    Width =1290
                    Height =315
                    FontWeight =700
                    Name ="Text110"
                    ControlSource ="=Sum([rent]*[countOfunitvin])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =2310
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =315
                    BorderThemeColorIndex =0
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =870
                            Width =1380
                            Height =315
                            Name ="Label111"
                            Caption ="Group Rent"
                            LayoutCachedLeft =870
                            LayoutCachedWidth =2250
                            LayoutCachedHeight =315
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4800
                    Top =1500
                    Width =4800
                    Height =315
                    FontWeight =700
                    TabIndex =1
                    Name ="Text986"
                    ControlSource ="FullUserName"
                    GroupTable =95

                    LayoutCachedLeft =4800
                    LayoutCachedTop =1500
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =1815
                    LayoutGroup =6
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =95
                End
            End
        End
        Begin PageFooter
            Height =315
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7635
                    Width =3600
                    Height =315
                    FontSize =9
                    Name ="Text39"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =7635
                    LayoutCachedWidth =11235
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    TextFontCharSet =163
                    BackStyle =0
                    IMESentenceMode =3
                    Width =3390
                    Height =315
                    FontSize =9
                    TabIndex =1
                    Name ="Text40"
                    ControlSource ="=\"printed on \" & Date() & \" @\" & Time()"

                    LayoutCachedWidth =3390
                    LayoutCachedHeight =315
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
