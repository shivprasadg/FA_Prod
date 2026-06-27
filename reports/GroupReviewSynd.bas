Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    AutoCenter = NotDefault
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
    ItemSuffix =1445
    Left =5550
    Top =3165
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x0cc85e8f8c4fe540
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
                    Name ="Text0"
                    ControlSource ="=\" Group Check for Lease: <b>\" & [MasterLease] & \" </b>Sch: <b>\" & [Schedule"
                        "] & \" </b>Group: <b>\" & [GroupNum]"
                    TextFormat =1

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
                    ControlSource ="GroupDesc"

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
            Height =8040
            Name ="GroupHeader0"
            Begin
                Begin Rectangle
                    Left =300
                    Top =1380
                    Width =10800
                    Height =360
                    Name ="Box430"
                    LayoutCachedLeft =300
                    LayoutCachedTop =1380
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =1740
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1035
                    Top =4755
                    Width =1080
                    Height =225
                    FontSize =10
                    FontWeight =700
                    TabIndex =39
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

                    LayoutCachedLeft =1035
                    LayoutCachedTop =4755
                    LayoutCachedWidth =2115
                    LayoutCachedHeight =4980
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =2160
                    Width =1860
                    Height =345
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

                    LayoutCachedLeft =240
                    LayoutCachedTop =2160
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =2505
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
                    IMESentenceMode =3
                    Left =240
                    Top =2715
                    Width =487
                    Height =390
                    FontSize =12
                    FontWeight =700
                    TabIndex =23
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

                    LayoutCachedLeft =240
                    LayoutCachedTop =2715
                    LayoutCachedWidth =727
                    LayoutCachedHeight =3105
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =2
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =64
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1035
                    Top =4530
                    Width =1080
                    Height =225
                    FontSize =10
                    FontWeight =700
                    TabIndex =38
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

                    LayoutCachedLeft =1035
                    LayoutCachedTop =4530
                    LayoutCachedWidth =2115
                    LayoutCachedHeight =4755
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1035
                    Top =4305
                    Width =1080
                    Height =225
                    FontSize =10
                    FontWeight =700
                    TabIndex =37
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

                    LayoutCachedLeft =1035
                    LayoutCachedTop =4305
                    LayoutCachedWidth =2115
                    LayoutCachedHeight =4530
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8925
                    Top =1860
                    Width =1530
                    Height =285
                    TabIndex =25
                    Name ="SalePrice\015\012\015\012\015\012"
                    ControlSource ="[SalePrice]"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="sale price to bank"
                    EventProcPrefix ="SalePrice______"

                    LayoutCachedLeft =8925
                    LayoutCachedTop =1860
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =2145
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =240
                    Top =4755
                    Width =795
                    Height =225
                    FontSize =10
                    Name ="Label23"
                    Caption ="BLCD"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =240
                    LayoutCachedTop =4755
                    LayoutCachedWidth =1035
                    LayoutCachedHeight =4980
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =240
                    Top =1860
                    Width =1860
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
                    LayoutCachedLeft =240
                    LayoutCachedTop =1860
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =2160
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =64
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =240
                    Top =4530
                    Width =795
                    Height =225
                    FontSize =10
                    Name ="Label25"
                    Caption ="Assigned"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =240
                    LayoutCachedTop =4530
                    LayoutCachedWidth =1035
                    LayoutCachedHeight =4755
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =240
                    Top =4305
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
                    LayoutCachedLeft =240
                    LayoutCachedTop =4305
                    LayoutCachedWidth =1035
                    LayoutCachedHeight =4530
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =7065
                    Top =1860
                    Width =1800
                    Height =285
                    FontSize =10
                    Name ="Label32"
                    Caption ="Sale Price"
                    LayoutCachedLeft =7065
                    LayoutCachedTop =1860
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =2145
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8925
                    Top =3240
                    Width =1530
                    Height =285
                    FontWeight =700
                    TabIndex =32
                    Name ="NetFee"
                    ControlSource ="=CCur((Nz([SalePrice],0)-Nz([OrigEquipCost],0))+[StripsAmount])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001010000 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =8925
                    LayoutCachedTop =3240
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =3525
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001010000ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =6780
                            Top =3240
                            Width =2085
                            Height =285
                            FontSize =10
                            FontWeight =700
                            Name ="Label65"
                            Caption ="Gross Bank Fee + Strips"
                            LayoutCachedLeft =6780
                            LayoutCachedTop =3240
                            LayoutCachedWidth =8865
                            LayoutCachedHeight =3525
                            RowStart =4
                            RowEnd =4
                            ColumnEnd =1
                            BorderThemeColorIndex =0
                            BorderTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8925
                    Top =2895
                    Width =1530
                    Height =285
                    TabIndex =30
                    Name ="StripsAmount"
                    ControlSource ="StripsAmount"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8925
                    LayoutCachedTop =2895
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =3180
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8925
                    Top =4275
                    Width =1530
                    Height =285
                    TabIndex =43
                    Name ="UTALine"
                    ControlSource ="=IIf([RebateUnitActual]>0,[RebateUnitActual],[RebateUnitProjected])*[UnitsInGrou"
                        "p]"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8925
                    LayoutCachedTop =4275
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =4560
                    RowStart =7
                    RowEnd =7
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =238
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8925
                    Top =5655
                    Width =1530
                    Height =285
                    FontWeight =700
                    TabIndex =53
                    Name ="AllInLine"
                    ControlSource ="=[NetBankFee]+[UTALine]+[PerDiemLine]+[ServFeePVLine]+[FundingExpense]"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001010000 ,
                        0xff000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End

                    LayoutCachedLeft =8925
                    LayoutCachedTop =5655
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =5940
                    RowStart =11
                    RowEnd =11
                    ColumnStart =2
                    ColumnEnd =2
                    BorderThemeColorIndex =2
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001010000ff000000ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =7065
                            Top =5655
                            Width =1800
                            Height =285
                            FontSize =10
                            FontWeight =700
                            Name ="Label91"
                            Caption ="Sub total"
                            LayoutCachedLeft =7065
                            LayoutCachedTop =5655
                            LayoutCachedWidth =8865
                            LayoutCachedHeight =5940
                            RowStart =11
                            RowEnd =11
                            ColumnStart =1
                            ColumnEnd =1
                            BorderThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =5655
                    Width =885
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =54
                    Name ="AllinPct"
                    ControlSource ="=[AllInLine]/[OrigEquipCost]"
                    Format ="Percent"

                    LayoutCachedLeft =10515
                    LayoutCachedTop =5655
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =5940
                    RowStart =11
                    RowEnd =11
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =3930
                    Width =885
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =41
                    Name ="Text96"
                    ControlSource ="=([SalePrice]-[OrigEquipCost]+Nz([StripsAmount],0)+Nz([SyndPenalty],0))/[OrigEqu"
                        "ipCost]"
                    Format ="Percent"

                    LayoutCachedLeft =10515
                    LayoutCachedTop =3930
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =4215
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6705
                    Top =555
                    Width =870
                    Height =315
                    TabIndex =6
                    Name ="SGrpBonus"
                    ControlSource ="Bonus"
                    StatusBarText ="Bonus Depreciation"
                    GroupTable =9

                    LayoutCachedLeft =6705
                    LayoutCachedTop =555
                    LayoutCachedWidth =7575
                    LayoutCachedHeight =870
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
                    Left =1305
                    Top =555
                    Width =840
                    Height =315
                    FontWeight =700
                    Name ="SGrpIndex"
                    ControlSource ="Index"
                    GroupTable =9

                    LayoutCachedLeft =1305
                    LayoutCachedTop =555
                    LayoutCachedWidth =2145
                    LayoutCachedHeight =870
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
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
                    Left =2205
                    Top =555
                    Width =1095
                    Height =315
                    FontWeight =700
                    TabIndex =1
                    Name ="SGrpIndDt"
                    ControlSource ="IndexDate"
                    Format ="mm/dd/yy"
                    GroupTable =9
                    ShowDatePicker =1

                    LayoutCachedLeft =2205
                    LayoutCachedTop =555
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =870
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
                    Left =3360
                    Top =555
                    Width =615
                    Height =315
                    FontWeight =700
                    TabIndex =2
                    Name ="SGprIndYr"
                    ControlSource ="IndexYear"
                    GroupTable =9

                    LayoutCachedLeft =3360
                    LayoutCachedTop =555
                    LayoutCachedWidth =3975
                    LayoutCachedHeight =870
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
                    Left =5820
                    Top =555
                    Width =825
                    Height =315
                    FontWeight =700
                    TabIndex =5
                    Name ="SgrpYield"
                    ControlSource ="Yield"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =5820
                    LayoutCachedTop =555
                    LayoutCachedWidth =6645
                    LayoutCachedHeight =870
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
                    Left =6705
                    Top =180
                    Width =870
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label33"
                    Caption ="Bonus"
                    GroupTable =9
                    LayoutCachedLeft =6705
                    LayoutCachedTop =180
                    LayoutCachedWidth =7575
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
                    Left =1305
                    Top =180
                    Width =840
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label34"
                    Caption ="Index"
                    GroupTable =9
                    LayoutCachedLeft =1305
                    LayoutCachedTop =180
                    LayoutCachedWidth =2145
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
                    Left =2205
                    Top =180
                    Width =1095
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label35"
                    Caption ="Date"
                    GroupTable =9
                    LayoutCachedLeft =2205
                    LayoutCachedTop =180
                    LayoutCachedWidth =3300
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
                    Left =3360
                    Top =180
                    Width =615
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label36"
                    Caption ="Year"
                    GroupTable =9
                    LayoutCachedLeft =3360
                    LayoutCachedTop =180
                    LayoutCachedWidth =3975
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
                    Left =5820
                    Top =180
                    Width =825
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label37"
                    Caption ="Yield"
                    GroupTable =9
                    LayoutCachedLeft =5820
                    LayoutCachedTop =180
                    LayoutCachedWidth =6645
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
                    Left =4845
                    Top =555
                    Width =915
                    Height =315
                    FontWeight =700
                    TabIndex =4
                    Name ="Text54"
                    ControlSource ="=[Yield]-[IndexRate]"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =4845
                    LayoutCachedTop =555
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =870
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
                    Left =4035
                    Top =555
                    Width =750
                    Height =315
                    FontWeight =700
                    TabIndex =3
                    Name ="IndexRate"
                    ControlSource ="IndexRate"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =4035
                    LayoutCachedTop =555
                    LayoutCachedWidth =4785
                    LayoutCachedHeight =870
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
                    Left =6705
                    Top =930
                    Width =870
                    Height =259
                    FontSize =9
                    FontWeight =700
                    TabIndex =16
                    ForeColor =9211020
                    Name ="SGrpSalesBonus"
                    ControlSource ="SalesBonus"
                    StatusBarText ="Bonus Depreciation"
                    GroupTable =9

                    LayoutCachedLeft =6705
                    LayoutCachedTop =930
                    LayoutCachedWidth =7575
                    LayoutCachedHeight =1189
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
                    Left =1305
                    Top =930
                    Width =840
                    Height =259
                    FontSize =9
                    FontWeight =700
                    TabIndex =10
                    ForeColor =9211020
                    Name ="SGrpSalesIndex"
                    ControlSource ="SalesIndex"
                    GroupTable =9

                    LayoutCachedLeft =1305
                    LayoutCachedTop =930
                    LayoutCachedWidth =2145
                    LayoutCachedHeight =1189
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
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2205
                    Top =930
                    Width =1095
                    Height =259
                    FontSize =9
                    FontWeight =700
                    TabIndex =11
                    ForeColor =9211020
                    Name ="SGrpSalesIndDt"
                    ControlSource ="SalesIndexDate"
                    Format ="mm/dd/yy"
                    GroupTable =9
                    ShowDatePicker =1

                    LayoutCachedLeft =2205
                    LayoutCachedTop =930
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =1189
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
                    Left =3360
                    Top =930
                    Width =615
                    Height =259
                    FontSize =9
                    FontWeight =700
                    TabIndex =12
                    ForeColor =9211020
                    Name ="SGprSalesIndYr"
                    ControlSource ="SalesIndexYear"
                    GroupTable =9

                    LayoutCachedLeft =3360
                    LayoutCachedTop =930
                    LayoutCachedWidth =3975
                    LayoutCachedHeight =1189
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
                    Left =4035
                    Top =930
                    Width =750
                    Height =259
                    FontSize =9
                    FontWeight =700
                    TabIndex =13
                    ForeColor =9211020
                    Name ="SgrpSalesIndRt"
                    ControlSource ="SalesIndexRate"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =4035
                    LayoutCachedTop =930
                    LayoutCachedWidth =4785
                    LayoutCachedHeight =1189
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
                    Left =5820
                    Top =930
                    Width =825
                    Height =259
                    FontSize =9
                    FontWeight =700
                    TabIndex =15
                    ForeColor =9211020
                    Name ="SgrpSalesYield"
                    ControlSource ="SalesYield"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =5820
                    LayoutCachedTop =930
                    LayoutCachedWidth =6645
                    LayoutCachedHeight =1189
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
                    Left =4845
                    Top =930
                    Width =915
                    Height =259
                    FontSize =9
                    FontWeight =700
                    TabIndex =14
                    ForeColor =9211020
                    Name ="Text84"
                    ControlSource ="=[SalesYield]-[SalesIndexRate]"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =4845
                    LayoutCachedTop =930
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =1189
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
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =240
                    Top =555
                    Width =1005
                    Height =315
                    ForeColor =16777215
                    Name ="Label86"
                    Caption ="Synd"
                    GroupTable =9
                    LayoutCachedLeft =240
                    LayoutCachedTop =555
                    LayoutCachedWidth =1245
                    LayoutCachedHeight =870
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin Label
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =240
                    Top =930
                    Width =1005
                    Height =259
                    FontSize =10
                    FontWeight =700
                    ForeColor =9211020
                    Name ="Label87"
                    Caption ="LO Target"
                    GroupTable =9
                    LayoutCachedLeft =240
                    LayoutCachedTop =930
                    LayoutCachedWidth =1245
                    LayoutCachedHeight =1189
                    RowStart =2
                    RowEnd =2
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
                    Left =1035
                    Top =4080
                    Width =1080
                    Height =225
                    FontSize =10
                    FontWeight =700
                    TabIndex =36
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

                    LayoutCachedLeft =1035
                    LayoutCachedTop =4080
                    LayoutCachedWidth =2115
                    LayoutCachedHeight =4305
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =240
                            Top =4080
                            Width =795
                            Height =225
                            FontSize =10
                            Name ="Label109"
                            Caption ="Sale"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =240
                            LayoutCachedTop =4080
                            LayoutCachedWidth =1035
                            LayoutCachedHeight =4305
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =4
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
                    Left =8925
                    Top =3585
                    Width =1530
                    Height =285
                    TabIndex =34
                    ForeColor =2366701
                    Name ="SynPenalty"
                    ControlSource ="=CDbl(Nz([SgrpSynPenalty],0))"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8925
                    LayoutCachedTop =3585
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =3870
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =6780
                    Top =3585
                    Width =2085
                    Height =285
                    FontSize =10
                    Name ="Label113"
                    Caption ="- Termination Penalty"
                    LayoutCachedLeft =6780
                    LayoutCachedTop =3585
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =3870
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =7065
                    Top =2550
                    Width =1800
                    Height =285
                    FontSize =10
                    Name ="Label119"
                    Caption ="Gross Bank Fee"
                    LayoutCachedLeft =7065
                    LayoutCachedTop =2550
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =2835
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8925
                    Top =2550
                    Width =1530
                    Height =285
                    TabIndex =27
                    Name ="GrossFee"
                    ControlSource ="=[SalePrice]-[OrigEquipCost]"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8925
                    LayoutCachedTop =2550
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =2835
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin EmptyCell
                    Left =240
                    Top =180
                    Width =1005
                    Height =315
                    Name ="EmptyCell163"
                    GroupTable =9
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =180
                    LayoutCachedWidth =1245
                    LayoutCachedHeight =495
                    LayoutGroup =1
                    GroupTable =9
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =4035
                    Top =180
                    Width =750
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label56"
                    Caption ="Rate"
                    GroupTable =9
                    LayoutCachedLeft =4035
                    LayoutCachedTop =180
                    LayoutCachedWidth =4785
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
                    Left =4845
                    Top =180
                    Width =915
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label55"
                    Caption ="Spread"
                    GroupTable =9
                    LayoutCachedLeft =4845
                    LayoutCachedTop =180
                    LayoutCachedWidth =5760
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
                    Left =720
                    Top =2505
                    Width =1380
                    Height =795
                    FontSize =8
                    FontWeight =700
                    Name ="Label419"
                    Caption =" F   Final (Synd)\015\012 H  Hold (FA)\015\012 I    Invetory (FA)\015\012 P   Pr"
                        "ojected"
                    GroupTable =64
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =720
                    LayoutCachedTop =2505
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =3300
                    RowStart =2
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =64
                End
                Begin Line
                    Top =7500
                    Width =11355
                    Name ="Line108"
                    LayoutCachedTop =7500
                    LayoutCachedWidth =11355
                    LayoutCachedHeight =7500
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =165
                    Top =7320
                    Width =2145
                    Height =345
                    FontSize =12
                    FontWeight =700
                    Name ="Label477"
                    Caption ="Unit Payment Detail"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =165
                    LayoutCachedTop =7320
                    LayoutCachedWidth =2310
                    LayoutCachedHeight =7665
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =7065
                    Top =2205
                    Width =1800
                    Height =285
                    FontSize =10
                    Name ="Label486"
                    Caption =" - Cost (OEC)"
                    LayoutCachedLeft =7065
                    LayoutCachedTop =2205
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =2490
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8925
                    Top =2205
                    Width =1530
                    Height =285
                    TabIndex =26
                    Name ="OrigEquipCost"
                    ControlSource ="OrigEquipCost"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8925
                    LayoutCachedTop =2205
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =2490
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8925
                    Top =3930
                    Width =1530
                    Height =285
                    FontWeight =700
                    TabIndex =40
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

                    LayoutCachedLeft =8925
                    LayoutCachedTop =3930
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =4215
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001010000ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =7065
                            Top =3930
                            Width =1800
                            Height =285
                            FontSize =10
                            FontWeight =700
                            Name ="Label602"
                            Caption ="Net Synd Fee"
                            LayoutCachedLeft =7065
                            LayoutCachedTop =3930
                            LayoutCachedWidth =8865
                            LayoutCachedHeight =4215
                            RowStart =6
                            RowEnd =6
                            ColumnStart =1
                            ColumnEnd =1
                            BorderThemeColorIndex =0
                            BorderTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =7065
                    Top =2895
                    Width =1800
                    Height =285
                    FontSize =10
                    Name ="Label67"
                    Caption ="+ Strips Amount"
                    LayoutCachedLeft =7065
                    LayoutCachedTop =2895
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =3180
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =3240
                    Width =885
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =33
                    Name ="Text625"
                    ControlSource ="=([SalePrice]-[OrigEquipCost]+Nz([StripsAmount],0))/[OrigEquipCost]"
                    Format ="Percent"

                    LayoutCachedLeft =10515
                    LayoutCachedTop =3240
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =3525
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7635
                    Top =555
                    Width =1200
                    Height =315
                    FontWeight =700
                    TabIndex =7
                    Name ="Text627"
                    ControlSource ="UnitResidual"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =9

                    LayoutCachedLeft =7635
                    LayoutCachedTop =555
                    LayoutCachedWidth =8835
                    LayoutCachedHeight =870
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
                    Left =7635
                    Top =930
                    Width =1200
                    Height =259
                    FontSize =9
                    FontWeight =700
                    TabIndex =17
                    ForeColor =9211020
                    Name ="Text628"
                    ControlSource ="SalesUnitResidual"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =9

                    LayoutCachedLeft =7635
                    LayoutCachedTop =930
                    LayoutCachedWidth =8835
                    LayoutCachedHeight =1189
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
                    Left =7635
                    Top =180
                    Width =1200
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label629"
                    Caption ="Residual"
                    GroupTable =9
                    LayoutCachedLeft =7635
                    LayoutCachedTop =180
                    LayoutCachedWidth =8835
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
                    Left =8895
                    Top =555
                    Width =1395
                    Height =315
                    FontWeight =700
                    TabIndex =8
                    Name ="Text647"
                    ControlSource ="=CCur((Nz([SalePrice],0)-Nz([OrigEquipCost],0)))"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =9

                    LayoutCachedLeft =8895
                    LayoutCachedTop =555
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =870
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
                    Left =8895
                    Top =930
                    Width =1395
                    Height =259
                    FontSize =9
                    FontWeight =700
                    TabIndex =18
                    ForeColor =9211020
                    Name ="Text648"
                    ControlSource ="[Fee_Projected]"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =9

                    LayoutCachedLeft =8895
                    LayoutCachedTop =930
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =1189
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
                    Left =8895
                    Top =180
                    Width =1395
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label649"
                    Caption ="Fee"
                    GroupTable =9
                    LayoutCachedLeft =8895
                    LayoutCachedTop =180
                    LayoutCachedWidth =10290
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
                    Left =2160
                    Top =2895
                    Width =4875
                    Height =285
                    FontSize =10
                    TabIndex =29
                    Name ="Label612"
                    ControlSource ="=IIf([StripsAmount]>0,\"(Assignee Term @ <B>\" & [LeaseTermPrimary]-[StripsNum] "
                        "& \"</b> mths, <B>\" & [StripsNum] & \"</b> Strips @ <b>\" & Format(([SalesRent]"
                        "*[unitsingroup]),\"$#,###.00\") & \" </b>Ea)\",\"\")"
                    RightPadding =0
                    TextFormat =1

                    LayoutCachedLeft =2160
                    LayoutCachedTop =2895
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =3180
                    RowStart =3
                    RowEnd =3
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2160
                    Top =6000
                    Width =4875
                    Height =285
                    FontSize =10
                    TabIndex =55
                    Name ="Label124"
                    ControlSource ="=IIf([WithHeldPymts]>0,\" (<B>\" & [WithHeldPymts] & \"</b> Witheld Payments @ <"
                        "b>\" & Format([LeasePymtForGroup],\"$#,###.00\") & \"</b> each)\",Null)"
                    RightPadding =0
                    TextFormat =1

                    LayoutCachedLeft =2160
                    LayoutCachedTop =6000
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =6285
                    RowStart =12
                    RowEnd =12
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Line
                    BorderWidth =1
                    Left =9060
                    Top =3874
                    Width =1380
                    Name ="Line665"
                    LayoutCachedLeft =9060
                    LayoutCachedTop =3874
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =3874
                    BorderThemeColorIndex =0
                End
                Begin Line
                    BorderWidth =1
                    Left =9060
                    Top =2505
                    Width =1380
                    Name ="Line674"
                    LayoutCachedLeft =9060
                    LayoutCachedTop =2505
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =2505
                    BorderThemeColorIndex =0
                End
                Begin Line
                    BorderWidth =1
                    Left =9060
                    Top =3165
                    Width =1380
                    Name ="Line676"
                    LayoutCachedLeft =9060
                    LayoutCachedTop =3165
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =3165
                    BorderThemeColorIndex =0
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =7065
                    Top =4620
                    Width =1800
                    Height =285
                    FontSize =10
                    Name ="Label682"
                    Caption ="+ Est. Perdiem"
                    LayoutCachedLeft =7065
                    LayoutCachedTop =4620
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =4905
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8925
                    Top =4620
                    Width =1530
                    Height =285
                    TabIndex =45
                    Name ="PerDiemLine"
                    ControlSource ="PerDiemToDisplay"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8925
                    LayoutCachedTop =4620
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =4905
                    RowStart =8
                    RowEnd =8
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =360
                    Top =1425
                    Width =10680
                    Height =285
                    TabIndex =20
                    Name ="Text712"
                    ControlSource ="=\"Committed Term of: <B>\" & [LeaseTermPrimary] & \"</b> months with payments @"
                        " <B> \" & Format((Sum([rent]*[UnitCount])),\"$#,###.00\") & \"</b> each to be pa"
                        "id <B>\" & [PayFreq] & \"</b> in <B>\" & [AdvanceArrears] & \"</b>\""
                    TextFormat =1

                    LayoutCachedLeft =360
                    LayoutCachedTop =1425
                    LayoutCachedWidth =11040
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
                    Left =240
                    Top =3780
                    Width =1875
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
                    LayoutCachedLeft =240
                    LayoutCachedTop =3780
                    LayoutCachedWidth =2115
                    LayoutCachedHeight =4080
                    ColumnEnd =1
                    LayoutGroup =4
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =3585
                    Width =885
                    Height =285
                    FontSize =10
                    TabIndex =35
                    ForeColor =1643706
                    Name ="Text953"
                    ControlSource ="=IIf([SyndPenalty]<>0,[SyndPenalty]/[OrigEquipCost],Null)"
                    Format ="Percent"

                    LayoutCachedLeft =10515
                    LayoutCachedTop =3585
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =3870
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =2895
                    Width =885
                    Height =285
                    FontSize =10
                    TabIndex =31
                    Name ="Text955"
                    ControlSource ="=IIf([StripsAmount]<>0,[StripsAmount]/[OrigEquipCost],Null)"
                    Format ="Percent"

                    LayoutCachedLeft =10515
                    LayoutCachedTop =2895
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =3180
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =4620
                    Width =885
                    Height =285
                    FontSize =10
                    TabIndex =46
                    Name ="Text957"
                    ControlSource ="=IIf([PerDiemLine]<>0,[PerDiemLine]/[OrigEquipCost],Null)"
                    Format ="Percent"

                    LayoutCachedLeft =10515
                    LayoutCachedTop =4620
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =4905
                    RowStart =8
                    RowEnd =8
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =4275
                    Width =885
                    Height =285
                    FontSize =10
                    TabIndex =44
                    Name ="Text959"
                    ControlSource ="=IIf([RebateUnitActual]>0,([RebateUnitActual]*[UnitsInGroup])/[OrigEquipCost],Nu"
                        "ll)"
                    Format ="Percent"

                    LayoutCachedLeft =10515
                    LayoutCachedTop =4275
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =4560
                    RowStart =7
                    RowEnd =7
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =2550
                    Width =885
                    Height =285
                    FontSize =10
                    TabIndex =28
                    Name ="Text961"
                    ControlSource ="=IIf([SalePrice]-[OrigEquipCost]<>0,([SalePrice]-[OrigEquipCost])/[OrigEquipCost"
                        "],Null)"
                    Format ="Percent"

                    LayoutCachedLeft =10515
                    LayoutCachedTop =2550
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =2835
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10350
                    Top =555
                    Width =765
                    Height =315
                    FontWeight =700
                    TabIndex =9
                    Name ="Text964"
                    ControlSource ="=([SalePrice]-[OrigEquipCost]+Nz([StripsAmount],0))/[OrigEquipCost]"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =10350
                    LayoutCachedTop =555
                    LayoutCachedWidth =11115
                    LayoutCachedHeight =870
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
                    Left =10350
                    Top =180
                    Width =765
                    Height =315
                    BackColor =12349952
                    ForeColor =16777215
                    Name ="Label965"
                    Caption ="Fee%"
                    GroupTable =9
                    LayoutCachedLeft =10350
                    LayoutCachedTop =180
                    LayoutCachedWidth =11115
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
                    Left =10350
                    Top =930
                    Width =765
                    Height =259
                    FontSize =9
                    FontWeight =700
                    TabIndex =19
                    ForeColor =9211020
                    Name ="Text966"
                    ControlSource ="=[Fee_Projected]/[OrigEquipCost]"
                    Format ="Percent"
                    GroupTable =9

                    LayoutCachedLeft =10350
                    LayoutCachedTop =930
                    LayoutCachedWidth =11115
                    LayoutCachedHeight =1189
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
                    Left =8925
                    Top =5295
                    Width =1530
                    Height =285
                    TabIndex =50
                    Name ="ServFeePVLine"
                    ControlSource ="ServFeePV"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8925
                    LayoutCachedTop =5295
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =5580
                    RowStart =9
                    RowEnd =9
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =5295
                    Width =885
                    Height =285
                    FontSize =10
                    TabIndex =51
                    Name ="Text981"
                    ControlSource ="=IIf([ServFeePV]>0,[ServFeePV]/[OrigEquipCost],Null)"
                    Format ="Percent"

                    LayoutCachedLeft =10515
                    LayoutCachedTop =5295
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =5580
                    RowStart =9
                    RowEnd =9
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Rectangle
                    BackStyle =0
                    Left =240
                    Top =4080
                    Width =1875
                    Height =960
                    Name ="Box1022"
                    LayoutCachedLeft =240
                    LayoutCachedTop =4080
                    LayoutCachedWidth =2115
                    LayoutCachedHeight =5040
                    BorderThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8925
                    Top =4950
                    Width =1530
                    Height =285
                    TabIndex =48
                    ForeColor =2366701
                    Name ="FundingExpense"
                    ControlSource ="=CDbl(Nz([InterimFundingExpense],0))"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8925
                    LayoutCachedTop =4950
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =5235
                    RowStart =10
                    RowEnd =10
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =4950
                    Width =885
                    Height =285
                    FontSize =10
                    TabIndex =49
                    ForeColor =1643706
                    Name ="Text1052"
                    ControlSource ="=IIf([FundingExpense]<>0,[FundingExpense]/[OrigEquipCost],Null)"
                    Format ="Percent"

                    LayoutCachedLeft =10515
                    LayoutCachedTop =4950
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =5235
                    RowStart =10
                    RowEnd =10
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =2085
                    Top =7680
                    Width =1395
                    Height =285
                    FontSize =10
                    Name ="Label103"
                    Caption ="Rent"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2085
                    LayoutCachedTop =7680
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =7965
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =8760
                    Top =7680
                    Width =1125
                    Height =285
                    FontSize =10
                    Name ="Label106"
                    Caption ="Sales Tax"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8760
                    LayoutCachedTop =7680
                    LayoutCachedWidth =9885
                    LayoutCachedHeight =7965
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =9960
                    Top =7680
                    Width =1425
                    Height =285
                    FontSize =10
                    Name ="Label698"
                    Caption ="Total Payment"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9960
                    LayoutCachedTop =7680
                    LayoutCachedWidth =11385
                    LayoutCachedHeight =7965
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =1515
                    Top =7680
                    Width =510
                    Height =285
                    FontSize =10
                    Name ="Label1161"
                    Caption ="Units"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1515
                    LayoutCachedTop =7680
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =7965
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4950
                    Top =7680
                    Width =1245
                    Height =285
                    FontSize =10
                    TabIndex =71
                    Name ="Label1093"
                    ControlSource ="=IIf(IsNull([Forms]![GroupForm]![AncillaryFeeLabel]),\"Ancillary Fee\",[Forms]!["
                        "GroupForm]![AncillaryFeeLabel])"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4950
                    LayoutCachedTop =7680
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =7965
                    ColumnStart =4
                    ColumnEnd =4
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3540
                    Top =7680
                    Width =1350
                    Height =285
                    FontSize =10
                    TabIndex =70
                    Name ="Label107"
                    ControlSource ="=[Forms]![GroupForm]![ServFeeLabel]"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3540
                    LayoutCachedTop =7680
                    LayoutCachedWidth =4890
                    LayoutCachedHeight =7965
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =6255
                    Top =7680
                    Width =1425
                    Height =285
                    FontSize =10
                    Name ="Label1183"
                    Caption ="Payment"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6255
                    LayoutCachedTop =7680
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =7965
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =6780
                    Top =6345
                    Width =2085
                    Height =285
                    FontSize =10
                    Name ="Label1197"
                    Caption ="- Holdback Residual"
                    LayoutCachedLeft =6780
                    LayoutCachedTop =6345
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =6630
                    RowStart =13
                    RowEnd =13
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8925
                    Top =6345
                    Width =1530
                    Height =285
                    TabIndex =58
                    Name ="HBResidual"
                    ControlSource ="HBResidual"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8925
                    LayoutCachedTop =6345
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =6630
                    RowStart =13
                    RowEnd =13
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =6345
                    Width =885
                    Height =285
                    FontSize =10
                    TabIndex =59
                    Name ="Text1199"
                    ControlSource ="=IIf([HBResidual]>0,[HBResidual]/[OrigEquipCost],Null)"
                    Format ="Percent"

                    LayoutCachedLeft =10515
                    LayoutCachedTop =6345
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =6630
                    RowStart =13
                    RowEnd =13
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =6780
                    Top =6690
                    Width =2085
                    Height =285
                    FontSize =10
                    Name ="Label1207"
                    Caption ="- Holdback Conversion"
                    LayoutCachedLeft =6780
                    LayoutCachedTop =6690
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =6975
                    RowStart =14
                    RowEnd =14
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8925
                    Top =6690
                    Width =1530
                    Height =285
                    TabIndex =60
                    Name ="HBConversion"
                    ControlSource ="HBConversion"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8925
                    LayoutCachedTop =6690
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =6975
                    RowStart =14
                    RowEnd =14
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =6690
                    Width =885
                    Height =285
                    FontSize =10
                    TabIndex =61
                    Name ="Text1209"
                    ControlSource ="=IIf([HBConversion]>0,[HBConversion]/[OrigEquipCost],Null)"
                    Format ="Percent"

                    LayoutCachedLeft =10515
                    LayoutCachedTop =6690
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =6975
                    RowStart =14
                    RowEnd =14
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8940
                    Top =7080
                    Width =1620
                    Height =285
                    FontWeight =700
                    TabIndex =68
                    Name ="FinalFee"
                    ControlSource ="=[AllInLine]-([WitheldLine]+[HBResidual]+[HBConversion])+Nz([SyndMiscValue_1],0)"
                        "+Nz([SyndMiscValue_2],0)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001010000 ,
                        0xff000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =96

                    LayoutCachedLeft =8940
                    LayoutCachedTop =7080
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =7365
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    BorderThemeColorIndex =2
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001010000ff000000ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =96
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =5370
                            Top =7080
                            Width =3510
                            Height =285
                            FontSize =10
                            FontWeight =700
                            Name ="Label1218"
                            Caption ="Group Total Fee"
                            GroupTable =96
                            LayoutCachedLeft =5370
                            LayoutCachedTop =7080
                            LayoutCachedWidth =8880
                            LayoutCachedHeight =7365
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =5
                            BorderThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =96
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10620
                    Top =7080
                    Width =810
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =69
                    Name ="Text1219"
                    ControlSource ="=[FinalFee]/[OrigEquipCost]"
                    Format ="Percent"
                    GroupTable =96

                    LayoutCachedLeft =10620
                    LayoutCachedTop =7080
                    LayoutCachedWidth =11430
                    LayoutCachedHeight =7365
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =96
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =7065
                    Top =6000
                    Width =1800
                    Height =285
                    FontSize =10
                    Name ="Label1244"
                    Caption ="- Witheld Amount"
                    LayoutCachedLeft =7065
                    LayoutCachedTop =6000
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =6285
                    RowStart =12
                    RowEnd =12
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8925
                    Top =6000
                    Width =1530
                    Height =285
                    TabIndex =56
                    Name ="WitheldLine"
                    ControlSource ="=CCur(Nz([WithHeldAmount],0))"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8925
                    LayoutCachedTop =6000
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =6285
                    RowStart =12
                    RowEnd =12
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =6000
                    Width =885
                    Height =285
                    FontSize =10
                    TabIndex =57
                    Name ="Text1246"
                    ControlSource ="=IIf([WitheldLine]<>0,[WitheldLine]/[OrigEquipCost],Null)"
                    Format ="Percent"

                    LayoutCachedLeft =10515
                    LayoutCachedTop =6000
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =6285
                    RowStart =12
                    RowEnd =12
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =7740
                    Top =7680
                    Width =390
                    Height =285
                    FontSize =10
                    Name ="Label1366"
                    Caption ="ST"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7740
                    LayoutCachedTop =7680
                    LayoutCachedWidth =8130
                    LayoutCachedHeight =7965
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =8205
                    Top =7680
                    Width =491
                    Height =285
                    FontSize =10
                    Name ="Label1369"
                    Caption ="RT"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8205
                    LayoutCachedTop =7680
                    LayoutCachedWidth =8696
                    LayoutCachedHeight =7965
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3225
                    Top =5310
                    Width =5640
                    FontSize =10
                    TabIndex =52
                    Name ="txtServiceFeeDynamicLabel"
                    ControlSource ="=IIf([Forms]![groupform]![PVFeesYN]=True And [ServFeePV]>0,\"Service Fee: <B>\" "
                        "& Format(Nz([SGrpFSLPmt],0)+Nz([SGrpRSLTPmt],0),\"$#,###.00\") & \" For \" & [Fe"
                        "eTerm] & \" months</B> PV @\" & Format(Nz([PVFeeRate],0.05),\"Percent\") & \" = "
                        " Service Fee \",\"+ Service Fee \")"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =3225
                    LayoutCachedTop =5310
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =5550
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Rectangle
                    Left =6720
                    Top =6960
                    Width =4680
                    Height =15
                    Name ="Box880"
                    LayoutCachedLeft =6720
                    LayoutCachedTop =6960
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =6975
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8940
                    Top =7020
                    Width =1620
                    Height =0
                    FontSize =9
                    TabIndex =63
                    Name ="SyndMiscValue_1"
                    ControlSource ="=IIf(Nz([SyndMiscValue1],0)<>0,[SyndMiscValue1],Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =96

                    LayoutCachedLeft =8940
                    LayoutCachedTop =7020
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =7020
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =96
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10620
                    Top =7020
                    Width =810
                    Height =0
                    FontSize =9
                    TabIndex =64
                    Name ="SyndMiscValue_1Pct"
                    ControlSource ="=IIf(Nz([SyndMiscValue1],0)<>0,[SyndMiscValue1]/[TotalCost],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =96

                    LayoutCachedLeft =10620
                    LayoutCachedTop =7020
                    LayoutCachedWidth =11430
                    LayoutCachedHeight =7020
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =96
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5370
                    Top =7020
                    Width =3510
                    Height =0
                    FontSize =9
                    TabIndex =62
                    ForeColor =10040879
                    Name ="SyndMiscLabel_1"
                    ControlSource ="=IIf(Nz([SyndMiscValue1],0)<>0,[SyndMiscLabel1],Null)"
                    FontName ="Segoe UI"
                    GroupTable =96

                    LayoutCachedLeft =5370
                    LayoutCachedTop =7020
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =7020
                    LayoutGroup =5
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =96
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8940
                    Top =7050
                    Width =1620
                    Height =0
                    TabIndex =66
                    Name ="SyndMiscValue_2"
                    ControlSource ="=IIf(Nz([SyndMiscValue2],0)<>0,[SyndMiscValue2],Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =96
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8940
                    LayoutCachedTop =7050
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =7050
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =96
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10620
                    Top =7050
                    Width =810
                    Height =0
                    FontSize =9
                    TabIndex =67
                    Name ="SyndMiscValue_2Pct"
                    ControlSource ="=IIf(Nz([SyndMiscValue2],0)<>0,[SyndMiscValue2]/[TotalCost],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =96
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10620
                    LayoutCachedTop =7050
                    LayoutCachedWidth =11430
                    LayoutCachedHeight =7050
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =96
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5370
                    Top =7050
                    Width =3510
                    Height =0
                    FontSize =9
                    TabIndex =65
                    ForeColor =10040879
                    Name ="SyndMiscLabel_2"
                    ControlSource ="=IIf(Nz([SyndMiscValue2],0)<>0,[SyndMiscLabel2],Null)"
                    FontName ="Segoe UI"
                    GroupTable =96
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5370
                    LayoutCachedTop =7050
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =7050
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =5
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =96
                End
                Begin Rectangle
                    Left =6720
                    Top =5610
                    Width =4680
                    Height =15
                    Name ="Box1437"
                    LayoutCachedLeft =6720
                    LayoutCachedTop =5610
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =5625
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6120
                    Top =4275
                    Width =2745
                    Height =285
                    FontSize =10
                    TabIndex =42
                    Name ="Label70"
                    ControlSource ="=\"+<b>\" & [UnitsInGroup] & \"</b> UTA @ \" & Format([UTAPer],\"$#,##0.00\") & "
                        "\" Each\""
                    TextFormat =1

                    LayoutCachedLeft =6120
                    LayoutCachedTop =4275
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =4560
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2160
                    Top =4950
                    Width =6705
                    Height =285
                    FontSize =10
                    TabIndex =47
                    Name ="Label1042"
                    ControlSource ="=IIf([FundingExpense]<>0,\"- Est Interim Expense for \" & [InterimLoanDays] & \""
                        " days (avg) @ \" & [InterimIntRate] & \"%\",\"- Est. Interim Funding Expense\")"

                    LayoutCachedLeft =2160
                    LayoutCachedTop =4950
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =5235
                    RowStart =10
                    RowEnd =10
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =2505
                    Width =487
                    Height =210
                    FontSize =12
                    FontWeight =700
                    TabIndex =22
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="Text1441"
                    ControlSource ="=\"\""
                    GroupTable =64
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =240
                    LayoutCachedTop =2505
                    LayoutCachedWidth =727
                    LayoutCachedHeight =2715
                    RowStart =2
                    RowEnd =2
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
                    Left =240
                    Top =3105
                    Width =487
                    Height =195
                    FontSize =12
                    FontWeight =700
                    TabIndex =24
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="Text1443"
                    ControlSource ="=\"\""
                    GroupTable =64
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =240
                    LayoutCachedTop =3105
                    LayoutCachedWidth =727
                    LayoutCachedHeight =3300
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =2
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =64
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =230
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2085
                    Width =1395
                    Height =230
                    FontSize =9
                    TabIndex =1
                    Name ="AvgOfUNITRENT"
                    ControlSource ="RENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2085
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =230
                    ColumnStart =2
                    ColumnEnd =2
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8760
                    Width =1125
                    Height =230
                    FontSize =9
                    TabIndex =5
                    Name ="AvgOfUnitTax"
                    ControlSource ="=Nz([UnitTax],0)"
                    Format ="$#,##0.00;($#,##0.00)"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8760
                    LayoutCachedWidth =9885
                    LayoutCachedHeight =230
                    ColumnStart =5
                    ColumnEnd =5
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3540
                    Width =1350
                    Height =230
                    FontSize =9
                    TabIndex =2
                    Name ="ServFee"
                    ControlSource ="ServiceFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="per unit FSL pmts"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3540
                    LayoutCachedWidth =4890
                    LayoutCachedHeight =230
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1515
                    Width =510
                    Height =230
                    FontSize =9
                    Name ="Text100"
                    ControlSource ="=[UnitCount]"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1515
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =230
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9960
                    Width =1425
                    Height =230
                    FontSize =9
                    TabIndex =6
                    Name ="Text709"
                    ControlSource ="=[Rent]+[ServFee]+[AncillaryFee]+[AvgOfUnitTax]"
                    Format ="$#,##0.00;($#,##0.00)"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9960
                    LayoutCachedWidth =11385
                    LayoutCachedHeight =230
                    ColumnStart =6
                    ColumnEnd =6
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =300
                    Width =1138
                    Height =230
                    FontSize =9
                    BorderColor =14211288
                    Name ="Label1089"
                    Caption ="Per →"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =300
                    LayoutCachedWidth =1438
                    LayoutCachedHeight =230
                    BackThemeColorIndex =1
                    BackShade =95.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4950
                    Width =1245
                    Height =230
                    FontSize =9
                    TabIndex =3
                    Name ="AncillaryFee"
                    ControlSource ="AncillaryFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="per unit FSL pmts"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4950
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =230
                    ColumnStart =4
                    ColumnEnd =4
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6255
                    Width =1425
                    Height =230
                    FontSize =9
                    TabIndex =4
                    Name ="Text1185"
                    ControlSource ="=[Rent]+[ServiceFee]+[AncillaryFee]"
                    Format ="$#,##0.00;($#,##0.00)"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6255
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =230
                    ColumnStart =6
                    ColumnEnd =6
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7740
                    Width =390
                    Height =230
                    FontSize =9
                    TabIndex =7
                    Name ="LocState"
                    ControlSource ="LocState"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7740
                    LayoutCachedWidth =8130
                    LayoutCachedHeight =230
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8205
                    Width =491
                    Height =230
                    FontSize =9
                    TabIndex =8
                    Name ="LocTaxRate"
                    ControlSource ="LTaxRate"
                    Format ="Percent"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8205
                    LayoutCachedWidth =8696
                    LayoutCachedHeight =230
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
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
                    Top =1680
                    Width =1440
                    Height =315
                    FontWeight =700
                    Name ="Label60"
                    Caption ="Date      "
                    GroupTable =95
                    LayoutCachedLeft =9660
                    LayoutCachedTop =1680
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =1995
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =95
                End
                Begin Line
                    Left =4815
                    Top =1620
                    Width =6315
                    Name ="Line61"
                    LayoutCachedLeft =4815
                    LayoutCachedTop =1620
                    LayoutCachedWidth =11130
                    LayoutCachedHeight =1620
                End
                Begin TextBox
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4800
                    Top =1680
                    Width =4800
                    Height =315
                    FontWeight =700
                    Name ="Text986"
                    ControlSource ="=\"Marc Gingold\""
                    GroupTable =95

                    LayoutCachedLeft =4800
                    LayoutCachedTop =1680
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =1995
                    LayoutGroup =3
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =95
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2085
                    Top =45
                    Width =1395
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="GRent"
                    ControlSource ="=Sum([rent]*[UnitCount])"
                    Format ="$#,##0.00;($#,##0.00)"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2085
                    LayoutCachedTop =45
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1515
                    Top =45
                    Width =510
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="Text1081"
                    ControlSource ="=Sum([UnitCount])"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1515
                    LayoutCachedTop =45
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8760
                    Top =45
                    Width =1125
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    Name ="GTaxes"
                    ControlSource ="=Sum(Nz([UnitTax],0)*[UnitCount])"
                    Format ="$#,##0.00;($#,##0.00)"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8760
                    LayoutCachedTop =45
                    LayoutCachedWidth =9885
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3540
                    Top =45
                    Width =1350
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="GServFee"
                    ControlSource ="=([ServiceFee])*Sum([UnitCount])"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="per unit FSL pmts"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3540
                    LayoutCachedTop =45
                    LayoutCachedWidth =4890
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9960
                    Top =45
                    Width =1425
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =7
                    Name ="Text1085"
                    ControlSource ="=[GRent]+[GServFee]+[GAncillaryFee]+[GTaxes]"
                    Format ="$#,##0.00;($#,##0.00)"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9960
                    LayoutCachedTop =45
                    LayoutCachedWidth =11385
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4950
                    Top =45
                    Width =1245
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    Name ="GAncillaryFee"
                    ControlSource ="=[AncillaryFee]*Sum([UnitCount])"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="per unit FSL pmts"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4950
                    LayoutCachedTop =45
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    TextFontFamily =34
                    Left =300
                    Top =45
                    Width =1138
                    Height =255
                    FontSize =9
                    BorderColor =14211288
                    Name ="Label1141"
                    Caption ="Group Totals"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =300
                    LayoutCachedTop =45
                    LayoutCachedWidth =1438
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    BackThemeColorIndex =1
                    BackShade =95.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6255
                    Top =45
                    Width =1425
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    Name ="RentNoTax"
                    ControlSource ="=([GRent]+[GServFee]+[GAncillaryFee])"
                    Format ="$#,##0.00;($#,##0.00)"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6255
                    LayoutCachedTop =45
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =300
                    Top =600
                    Width =11040
                    Height =780
                    FontSize =9
                    TabIndex =8
                    Name ="Text2171"
                    ControlSource ="=DLookUp(\"SynMemo\",\"schgrp\",\"sgrpId=\" & [GroupID])"

                    LayoutCachedLeft =300
                    LayoutCachedTop =600
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =1380
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            TextAlign =1
                            TextFontFamily =34
                            Left =300
                            Top =360
                            Width =615
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2172"
                            Caption ="Notes:"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =300
                            LayoutCachedTop =360
                            LayoutCachedWidth =915
                            LayoutCachedHeight =600
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =0
                            BorderTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
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
            Height =240
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Width =7080
                    FontSize =8
                    BorderColor =12632256
                    Name ="LastUpdate"
                    ControlSource ="=\"Last Modified by \" & [LastUpdatedBy] & \" on \" & [LastUpdatedDate]"
                    StatusBarText ="upadj"
                    FontName ="Segoe UI"
                    ControlTipText ="Rate Adjustment - If Upward"
                    ShowDatePicker =1

                    LayoutCachedWidth =7080
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
    End
End
