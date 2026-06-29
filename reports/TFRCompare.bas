Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =19680
    DatasheetFontHeight =11
    ItemSuffix =800
    Left =6375
    Top =3165
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x58778fa50da5e540
    End
    RecordSource ="TFRCompare"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
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
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            SortOrder = NotDefault
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =2
            ControlSource ="FundingStatus"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="AccountName"
        End
        Begin BreakLevel
            ControlSource ="Sch_Grp"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="TargetFundDate"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =824
            Name ="GroupHeader0"
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15
                    Top =60
                    Width =6030
                    Height =439
                    FontSize =14
                    BackColor =10319446
                    BorderColor =10319446
                    Name ="Text374"
                    ControlSource ="=\"<B>\" & [FundingStatus] & \"</b>\""
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =15
                    LayoutCachedTop =60
                    LayoutCachedWidth =6045
                    LayoutCachedHeight =499
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =19092
                    Top =516
                    Width =576
                    Height =264
                    FontSize =8
                    TabIndex =20
                    ForeColor =8872517
                    Name ="Text375"
                    ControlSource ="=Format([FundingNetFeeCom]/[FundingOEC],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =19092
                    LayoutCachedTop =516
                    LayoutCachedWidth =19668
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =18012
                    Top =516
                    Width =1080
                    Height =264
                    FontSize =9
                    TabIndex =19
                    ForeColor =0
                    Name ="FundingNetFeeCom"
                    ControlSource ="=Sum([NetTotalFee])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =18012
                    LayoutCachedTop =516
                    LayoutCachedWidth =19092
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16596
                    Top =516
                    Width =576
                    Height =264
                    FontSize =8
                    TabIndex =17
                    ForeColor =8872517
                    Name ="Text377"
                    ControlSource ="=Format([FundingNetFee]/[FundingOEC],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =16596
                    LayoutCachedTop =516
                    LayoutCachedWidth =17172
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15516
                    Top =516
                    Width =1080
                    Height =264
                    FontSize =9
                    TabIndex =16
                    ForeColor =0
                    Name ="FundingNetFee"
                    ControlSource ="=Sum([TotalFee])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =15516
                    LayoutCachedTop =516
                    LayoutCachedWidth =16596
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17172
                    Top =516
                    Width =840
                    Height =264
                    FontSize =9
                    FontWeight =700
                    TabIndex =18
                    BorderColor =9211020
                    ForeColor =0
                    Name ="Text379"
                    ControlSource ="=Sum([NetSyndFee])*[DealCommission]"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =17172
                    LayoutCachedTop =516
                    LayoutCachedWidth =18012
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15000
                    Top =516
                    Width =510
                    Height =264
                    FontSize =8
                    TabIndex =15
                    ForeColor =8872517
                    Name ="Text380"
                    ControlSource ="=Format([FundingUTA]/[FundingOEC],\"Percent\")"
                    Format ="Percent"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =15000
                    LayoutCachedTop =516
                    LayoutCachedWidth =15510
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9972
                    Top =516
                    Width =600
                    Height =264
                    FontSize =8
                    TabIndex =8
                    ForeColor =8872517
                    Name ="Text381"
                    ControlSource ="=Format([FundingTermPenalty]/[FundingOEC],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =9972
                    LayoutCachedTop =516
                    LayoutCachedWidth =10572
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8916
                    Top =516
                    Width =1056
                    Height =264
                    FontSize =9
                    FontWeight =700
                    TabIndex =7
                    BorderColor =9211020
                    ForeColor =0
                    Name ="FundingTermPenalty"
                    ControlSource ="=Sum([TermPenalty])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8916
                    LayoutCachedTop =516
                    LayoutCachedWidth =9972
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14232
                    Top =516
                    Width =768
                    Height =264
                    FontSize =9
                    FontWeight =700
                    TabIndex =14
                    BorderColor =9211020
                    ForeColor =0
                    Name ="FundingUTA"
                    ControlSource ="=Sum([UTA])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14232
                    LayoutCachedTop =516
                    LayoutCachedWidth =15000
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13692
                    Top =516
                    Width =540
                    Height =264
                    FontSize =9
                    TabIndex =13
                    ForeColor =0
                    Name ="FundingDays"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13692
                    LayoutCachedTop =516
                    LayoutCachedWidth =14232
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13152
                    Top =516
                    Width =540
                    Height =264
                    FontSize =8
                    TabIndex =12
                    ForeColor =8872517
                    Name ="Text385"
                    ControlSource ="=Format([FundingPerDiem]/[FundingOEC],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =13152
                    LayoutCachedTop =516
                    LayoutCachedWidth =13692
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12252
                    Top =516
                    Width =900
                    Height =264
                    FontSize =9
                    FontWeight =700
                    TabIndex =11
                    BorderColor =9211020
                    ForeColor =0
                    Name ="FundingPerDiem"
                    ControlSource ="=Sum([PDRent])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12252
                    LayoutCachedTop =516
                    LayoutCachedWidth =13152
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8328
                    Top =516
                    Width =588
                    Height =264
                    FontSize =8
                    TabIndex =6
                    ForeColor =8872517
                    Name ="Text387"
                    ControlSource ="=Format([FundingSyndFee]/[FundingOEC],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8328
                    LayoutCachedTop =516
                    LayoutCachedWidth =8916
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7248
                    Top =516
                    Width =1080
                    Height =264
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    BorderColor =9211020
                    ForeColor =0
                    Name ="FundingSyndFee"
                    ControlSource ="=Sum([SyndFee])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7248
                    LayoutCachedTop =516
                    LayoutCachedWidth =8328
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6036
                    Top =516
                    Width =1212
                    Height =264
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    BorderColor =9211020
                    ForeColor =0
                    Name ="FundingOEC"
                    ControlSource ="=Sum([TotalVendorCost])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6036
                    LayoutCachedTop =516
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4500
                    Top =516
                    Width =1536
                    Height =264
                    FontSize =10
                    TabIndex =3
                    ForeColor =5855577
                    Name ="Text390"
                    ControlSource ="=\"← Totals →\""
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4500
                    LayoutCachedTop =516
                    LayoutCachedWidth =6036
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4008
                    Top =516
                    Width =492
                    Height =264
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    BorderColor =9211020
                    ForeColor =0
                    Name ="FundingUnits"
                    ControlSource ="=Sum([UnitCount])"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4008
                    LayoutCachedTop =516
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10572
                    Top =516
                    Width =1080
                    Height =264
                    FontSize =9
                    FontWeight =700
                    TabIndex =9
                    BorderColor =9211020
                    ForeColor =0
                    Name ="FundingNetSyndFee"
                    ControlSource ="=Sum([NetSyndFee])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10572
                    LayoutCachedTop =516
                    LayoutCachedWidth =11652
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11652
                    Top =516
                    Width =600
                    Height =264
                    FontSize =8
                    TabIndex =10
                    ForeColor =8872517
                    Name ="Text393"
                    ControlSource ="=Format([FundingNetSyndFee]/[FundingOEC],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =11652
                    LayoutCachedTop =516
                    LayoutCachedWidth =12252
                    LayoutCachedHeight =780
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =4008
                    Top =60
                    Width =492
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    BorderColor =0
                    Name ="Label394"
                    Caption ="\015\012Units"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4008
                    LayoutCachedTop =60
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =510
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =4500
                    Top =60
                    Width =765
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label395"
                    Caption =" "
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4500
                    LayoutCachedTop =60
                    LayoutCachedWidth =5265
                    LayoutCachedHeight =510
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =5268
                    Top =60
                    Width =768
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label396"
                    Caption =" "
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5268
                    LayoutCachedTop =60
                    LayoutCachedWidth =6036
                    LayoutCachedHeight =510
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =6036
                    Top =60
                    Width =1212
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label397"
                    Caption ="Vendor Cost\015\012+ Parts"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6036
                    LayoutCachedTop =60
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =510
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =7248
                    Top =60
                    Width =1080
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label398"
                    Caption ="   Gross  \015\012    Synd Fee $  "
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7248
                    LayoutCachedTop =60
                    LayoutCachedWidth =8328
                    LayoutCachedHeight =510
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =8328
                    Top =60
                    Width =588
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label399"
                    Caption ="\015\012%"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8328
                    LayoutCachedTop =60
                    LayoutCachedWidth =8916
                    LayoutCachedHeight =510
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =12252
                    Top =60
                    Width =900
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label400"
                    Caption ="\015\012PerDiem  "
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12252
                    LayoutCachedTop =60
                    LayoutCachedWidth =13152
                    LayoutCachedHeight =510
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =13152
                    Top =60
                    Width =540
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label401"
                    Caption ="\015\012%"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13152
                    LayoutCachedTop =60
                    LayoutCachedWidth =13692
                    LayoutCachedHeight =510
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =13692
                    Top =60
                    Width =540
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label402"
                    Caption ="PD Days"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13692
                    LayoutCachedTop =60
                    LayoutCachedWidth =14232
                    LayoutCachedHeight =510
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =15000
                    Top =60
                    Width =510
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label403"
                    Caption ="\015\012%"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15000
                    LayoutCachedTop =60
                    LayoutCachedWidth =15510
                    LayoutCachedHeight =510
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =14232
                    Top =60
                    Width =768
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label404"
                    Caption ="\015\012UTA"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14232
                    LayoutCachedTop =60
                    LayoutCachedWidth =15000
                    LayoutCachedHeight =510
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =8916
                    Top =60
                    Width =1056
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label405"
                    Caption ="Term Penalty$"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8916
                    LayoutCachedTop =60
                    LayoutCachedWidth =9972
                    LayoutCachedHeight =510
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =9972
                    Top =60
                    Width =600
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label406"
                    Caption ="\015\012%"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9972
                    LayoutCachedTop =60
                    LayoutCachedWidth =10572
                    LayoutCachedHeight =510
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =10572
                    Top =60
                    Width =1080
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label407"
                    Caption ="   Net  \015\012    Synd Fee $  "
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10572
                    LayoutCachedTop =60
                    LayoutCachedWidth =11652
                    LayoutCachedHeight =510
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =11652
                    Top =60
                    Width =600
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label408"
                    Caption ="\015\012%"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11652
                    LayoutCachedTop =60
                    LayoutCachedWidth =12252
                    LayoutCachedHeight =510
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =17172
                    Top =60
                    Width =840
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label409"
                    Caption ="BDEX\015\012Comm  "
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17172
                    LayoutCachedTop =60
                    LayoutCachedWidth =18012
                    LayoutCachedHeight =510
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =15516
                    Top =60
                    Width =1080
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label410"
                    Caption ="     \015\012     Revenue  "
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15516
                    LayoutCachedTop =60
                    LayoutCachedWidth =16596
                    LayoutCachedHeight =510
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =16596
                    Top =60
                    Width =576
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label411"
                    Caption ="\015\012%"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16596
                    LayoutCachedTop =60
                    LayoutCachedWidth =17172
                    LayoutCachedHeight =510
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =18012
                    Top =60
                    Width =1080
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label412"
                    Caption ="   Net  \015\012     Revenue  "
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18012
                    LayoutCachedTop =60
                    LayoutCachedWidth =19092
                    LayoutCachedHeight =510
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =19092
                    Top =60
                    Width =576
                    Height =450
                    FontSize =9
                    BackColor =10319446
                    Name ="Label413"
                    Caption ="\015\012%"
                    GroupTable =18
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19092
                    LayoutCachedTop =60
                    LayoutCachedWidth =19668
                    LayoutCachedHeight =510
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =18
                End
                Begin Rectangle
                    OldBorderStyle =0
                    Top =795
                    Width =19512
                    Height =29
                    BackColor =9605778
                    Name ="Box452"
                    LayoutCachedTop =795
                    LayoutCachedWidth =19512
                    LayoutCachedHeight =824
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Top =495
                    Width =3960
                    Height =288
                    TabIndex =21
                    BackColor =15527148
                    ForeColor =0
                    Name ="Text453"
                    ControlSource ="=IIf([FundingStatus] Like \"*Syndicated*\",\"<i> \" & [OpenArgs] & \"</i>      \""
                        ",\"\") & \" Totals →\""
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedTop =495
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =783
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Left =4821
                    Top =144
                    Width =1080
                    Height =278
                    FontSize =10
                    TabIndex =1
                    ForeColor =0
                    Name ="ArchiveReport"
                    Caption ="Archive"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000b17d4a1bb17d4affb17d4affb17d4a48 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000b17d4a21b17d4a27b17d4affb17d4aff ,
                        0xb17d4a4882c2ea0f82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2ea6000000000b17d4a8db17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4aff0000000082c2eabd82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaa500000000b17d4af3b17d4aa200000000b17d4a42b17d4affb17d4aff ,
                        0xb17d4a3082c2ea0f82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eae400000000b17d4afc00000000b17d4a2db17d4affb17d4affb17d4a30 ,
                        0x82c2ea0f82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2ea2a000000000000000000000000000000000000000082c2ea0f ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2ea6982c2eaff0000000082c2ea7e82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaae82c2eaff0000000082c2ea1882c2eaf982c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaea82c2eaff00000000727272ff82c2ea6682c2eaf982c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eab482c2eaff00000000727272ff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000082c2eaff00000000727272ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0x727272ff0000000082c2ea6300000000727272ffffffffffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffffffffffffb3b3b3ffffffffffffffffffffffffff ,
                        0x727272ff000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffffffffffffb3b3b3ffffffffffffffffffffffffff ,
                        0x727272ff000000000000000000000000b8824dffb8824dffb8824dffb8824dff ,
                        0xb8824dffb8824dffb8824dffb8824dffb8824dffb8824dffb8824dffb8824dff ,
                        0xb8824dff000000000000000000000000b8824dffb8824dffb8824dffb8824dff ,
                        0xb8824dffb8824dffb8824dffb8824dffb8824dffb8824dffb8824dffb8824dff ,
                        0xb8824dff000000000000000000000000b8824dffb8824dffb8824dffb8824dff ,
                        0xb8824dffb8824dffb8824dffb8824dffb8824dffb8824dffb8824dffb8824dff ,
                        0xb8824dff00000000
                    End

                    LayoutCachedLeft =4821
                    LayoutCachedTop =144
                    LayoutCachedWidth =5901
                    LayoutCachedHeight =422
                    PictureCaptionArrangement =5
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =11830108
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    QuickStyle =7
                    QuickStyleMask =-193
                    Overlaps =1
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =780
            BreakLevel =1
            Name ="GroupHeader1"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =1
                    Top =348
                    Width =749
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label6"
                    Caption ="\015\012Awarded"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =348
                    LayoutCachedWidth =749
                    LayoutCachedHeight =780
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =1980
                    Top =348
                    Width =2028
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label20"
                    Caption ="\015\012OEM and Type"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1980
                    LayoutCachedTop =348
                    LayoutCachedWidth =4008
                    LayoutCachedHeight =780
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =744
                    Top =348
                    Width =1236
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label13"
                    Caption ="\015\012Sch Grp"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =744
                    LayoutCachedTop =348
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =780
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =4008
                    Top =348
                    Width =492
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    BorderColor =0
                    ForeColor =8210719
                    Name ="Label27"
                    Caption ="\015\012Units"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4008
                    LayoutCachedTop =348
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =780
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =4500
                    Top =348
                    Width =765
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label80"
                    Caption ="Fund\015\012Mth"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4500
                    LayoutCachedTop =348
                    LayoutCachedWidth =5265
                    LayoutCachedHeight =780
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =5268
                    Top =348
                    Width =768
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label112"
                    Caption ="\015\012Bank"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5268
                    LayoutCachedTop =348
                    LayoutCachedWidth =6036
                    LayoutCachedHeight =780
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =6036
                    Top =348
                    Width =1212
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label114"
                    Caption ="Vendor Cost\015\012+ Parts"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6036
                    LayoutCachedTop =348
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =780
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =7248
                    Top =348
                    Width =1080
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label126"
                    Caption ="   Gross  \015\012    Synd Fee $  "
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7248
                    LayoutCachedTop =348
                    LayoutCachedWidth =8328
                    LayoutCachedHeight =780
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =8328
                    Top =348
                    Width =588
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8872517
                    Name ="Label133"
                    Caption ="\015\012%"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8328
                    LayoutCachedTop =348
                    LayoutCachedWidth =8916
                    LayoutCachedHeight =780
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =12252
                    Top =348
                    Width =900
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label149"
                    Caption ="Net\015\012PerDiem  "
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12252
                    LayoutCachedTop =348
                    LayoutCachedWidth =13152
                    LayoutCachedHeight =780
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =13152
                    Top =348
                    Width =540
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8872517
                    Name ="Label151"
                    Caption ="\015\012%"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13152
                    LayoutCachedTop =348
                    LayoutCachedWidth =13692
                    LayoutCachedHeight =780
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =13692
                    Top =348
                    Width =540
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label158"
                    Caption ="PD Days"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13692
                    LayoutCachedTop =348
                    LayoutCachedWidth =14232
                    LayoutCachedHeight =780
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =15000
                    Top =348
                    Width =495
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8872517
                    Name ="Label174"
                    Caption ="\015\012%"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15000
                    LayoutCachedTop =348
                    LayoutCachedWidth =15495
                    LayoutCachedHeight =780
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =14232
                    Top =348
                    Width =768
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label176"
                    Caption ="\015\012UTA"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14232
                    LayoutCachedTop =348
                    LayoutCachedWidth =15000
                    LayoutCachedHeight =780
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Top =60
                    Width =4005
                    Height =288
                    BackColor =15527148
                    ForeColor =0
                    Name ="Text102"
                    ControlSource ="=\"<b>\" & [AccountName] & \"</b>  <i>\" & [FundingStatus]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedTop =60
                    LayoutCachedWidth =4005
                    LayoutCachedHeight =348
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =8916
                    Top =348
                    Width =1056
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label244"
                    Caption ="Term Penalty"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8916
                    LayoutCachedTop =348
                    LayoutCachedWidth =9972
                    LayoutCachedHeight =780
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =9972
                    Top =348
                    Width =600
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8872517
                    Name ="Label264"
                    Caption ="\015\012%"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9972
                    LayoutCachedTop =348
                    LayoutCachedWidth =10572
                    LayoutCachedHeight =780
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =10572
                    Top =348
                    Width =1080
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label278"
                    Caption ="   Net  \015\012    Synd Fee $  "
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10572
                    LayoutCachedTop =348
                    LayoutCachedWidth =11652
                    LayoutCachedHeight =780
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =11652
                    Top =348
                    Width =600
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8872517
                    Name ="Label291"
                    Caption ="\015\012%"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11652
                    LayoutCachedTop =348
                    LayoutCachedWidth =12252
                    LayoutCachedHeight =780
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =17148
                    Top =348
                    Width =840
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label311"
                    Caption ="BDEX\015\012Comm  "
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17148
                    LayoutCachedTop =348
                    LayoutCachedWidth =17988
                    LayoutCachedHeight =780
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =15492
                    Top =348
                    Width =1080
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label342"
                    Caption ="     \015\012     Revenue "
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15492
                    LayoutCachedTop =348
                    LayoutCachedWidth =16572
                    LayoutCachedHeight =780
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =16572
                    Top =348
                    Width =576
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8872517
                    Name ="Label350"
                    Caption ="\015\012%"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16572
                    LayoutCachedTop =348
                    LayoutCachedWidth =17148
                    LayoutCachedHeight =780
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =17988
                    Top =348
                    Width =1080
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8210719
                    Name ="Label361"
                    Caption ="   Net  \015\012     Revenue  "
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17988
                    LayoutCachedTop =348
                    LayoutCachedWidth =19068
                    LayoutCachedHeight =780
                    ColumnStart =22
                    ColumnEnd =22
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =19068
                    Top =348
                    Width =576
                    Height =432
                    FontSize =9
                    BackColor =15527148
                    ForeColor =8872517
                    Name ="Label369"
                    Caption ="\015\012%"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19068
                    LayoutCachedTop =348
                    LayoutCachedWidth =19644
                    LayoutCachedHeight =780
                    ColumnStart =23
                    ColumnEnd =23
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =240
            Name ="Detail"
            AlternateBackColor =16249583
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Width =749
                    FontSize =9
                    ForeColor =0
                    Name ="LPAF Signed"
                    ControlSource ="LPAFMonth"
                    EventProcPrefix ="LPAF_Signed"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =749
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1980
                    Width =2028
                    ColumnWidth =2790
                    FontSize =9
                    TabIndex =2
                    ForeColor =0
                    Name ="UnitDesc"
                    ControlSource ="UnitDesc"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1980
                    LayoutCachedWidth =4008
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4008
                    Width =492
                    FontSize =9
                    TabIndex =3
                    BorderColor =0
                    ForeColor =0
                    Name ="UnitCount"
                    ControlSource ="UnitCount"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4008
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =744
                    Width =1236
                    FontSize =9
                    TabIndex =1
                    ForeColor =0
                    Name ="Sch Grp"
                    ControlSource ="Sch_Grp"
                    OnDblClick ="[Event Procedure]"
                    EventProcPrefix ="Sch_Grp"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =744
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4500
                    Width =765
                    FontSize =9
                    TabIndex =4
                    ForeColor =0
                    Name ="TargetFundDate"
                    ControlSource ="TargetFundMonth"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =4500
                    LayoutCachedWidth =5265
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5268
                    Width =768
                    FontSize =9
                    TabIndex =5
                    ForeColor =0
                    Name ="Text107"
                    ControlSource ="Bank"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =5268
                    LayoutCachedWidth =6036
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6036
                    Width =1212
                    FontSize =9
                    TabIndex =6
                    ForeColor =0
                    Name ="Text119"
                    ControlSource ="TotalVendorCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =6036
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7248
                    Width =1080
                    FontSize =9
                    TabIndex =7
                    ForeColor =0
                    Name ="SyndFee"
                    ControlSource ="SyndFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =7248
                    LayoutCachedWidth =8328
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8328
                    Width =588
                    FontSize =8
                    TabIndex =8
                    ForeColor =8872517
                    Name ="Text128"
                    ControlSource ="=Format([SyndFeePct],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x010000007a000000010000000100000000000000000000000c00000001000100 ,
                        0xed1c2400ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00530079006e0064004600650065005d003c00300000000000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8328
                    LayoutCachedWidth =8916
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000100ed1c2400ffffff000b0000005b00 ,
                        0x530079006e0064004600650065005d003c003000000000000000000000000000 ,
                        0x000000000000000000
                    End
                    GroupTable =7
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12252
                    Width =900
                    FontSize =9
                    TabIndex =13
                    ForeColor =0
                    Name ="PerDiemGroup"
                    ControlSource ="PDRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =12252
                    LayoutCachedWidth =13152
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13152
                    Width =540
                    FontSize =8
                    TabIndex =14
                    ForeColor =8872517
                    Name ="Text144"
                    ControlSource ="PDRentPct"
                    Format ="Percent"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =13152
                    LayoutCachedWidth =13692
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13692
                    Width =540
                    FontSize =9
                    TabIndex =15
                    ForeColor =0
                    Name ="Text153"
                    ControlSource ="PDDaysAvg"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13692
                    LayoutCachedWidth =14232
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14232
                    Width =768
                    FontSize =9
                    TabIndex =16
                    ForeColor =0
                    Name ="Text164"
                    ControlSource ="UTA"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =14232
                    LayoutCachedWidth =15000
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15000
                    Width =495
                    FontSize =8
                    TabIndex =17
                    ForeColor =8872517
                    Name ="Text165"
                    ControlSource ="UTAPct"
                    Format ="Percent"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =15000
                    LayoutCachedWidth =15495
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8916
                    Width =1056
                    FontSize =9
                    TabIndex =9
                    ForeColor =0
                    Name ="Text245"
                    ControlSource ="TermPenalty"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8916
                    LayoutCachedWidth =9972
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9972
                    Width =600
                    FontSize =8
                    TabIndex =10
                    ForeColor =8872517
                    Name ="Text258"
                    ControlSource ="=Format([TermPenaltyPct],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =9972
                    LayoutCachedWidth =10572
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10572
                    Width =1080
                    FontSize =9
                    TabIndex =11
                    ForeColor =0
                    Name ="Text279NetSyndFeeDelta"
                    ControlSource ="NetSyndFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =10572
                    LayoutCachedWidth =11652
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11652
                    Width =600
                    FontSize =8
                    TabIndex =12
                    ForeColor =8872517
                    Name ="Text292"
                    ControlSource ="=[NetSyndPct]"
                    Format ="Percent"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =11652
                    LayoutCachedWidth =12252
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17148
                    Width =840
                    FontSize =9
                    TabIndex =20
                    ForeColor =0
                    Name ="Text312"
                    ControlSource ="BDEXComm"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =17148
                    LayoutCachedWidth =17988
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15492
                    Width =1080
                    FontSize =9
                    TabIndex =18
                    ForeColor =0
                    Name ="Text338"
                    ControlSource ="TotalFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =15492
                    LayoutCachedWidth =16572
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16572
                    Width =576
                    FontSize =8
                    TabIndex =19
                    ForeColor =8872517
                    Name ="Text340"
                    ControlSource ="TotalFeePct"
                    Format ="Percent"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =16572
                    LayoutCachedWidth =17148
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17988
                    Width =1080
                    FontSize =9
                    TabIndex =21
                    ForeColor =0
                    Name ="NetTotalFee"
                    ControlSource ="NetTotalFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =17988
                    LayoutCachedWidth =19068
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =22
                    ColumnEnd =22
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =19068
                    Width =576
                    FontSize =8
                    TabIndex =22
                    ForeColor =8872517
                    Name ="Text363"
                    ControlSource ="NetTotalPct"
                    Format ="Percent"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =19068
                    LayoutCachedWidth =19644
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =23
                    ColumnEnd =23
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =7
                End
                Begin EmptyCell
                    Left =1980
                    Top =240
                    Width =1995
                    Height =0
                    Name ="EmptyCell615"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1980
                    LayoutCachedTop =240
                    LayoutCachedWidth =3975
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =7
                End
                Begin EmptyCell
                    Left =3972
                    Top =240
                    Width =36
                    Height =0
                    Name ="EmptyCell616"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3972
                    LayoutCachedTop =240
                    LayoutCachedWidth =4008
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =7
                End
                Begin EmptyCell
                    Left =4008
                    Top =240
                    Width =492
                    Height =0
                    Name ="EmptyCell617"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4008
                    LayoutCachedTop =240
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =7
                End
                Begin EmptyCell
                    Left =4500
                    Top =240
                    Width =765
                    Height =0
                    Name ="EmptyCell618"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4500
                    LayoutCachedTop =240
                    LayoutCachedWidth =5265
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =7
                End
                Begin EmptyCell
                    Left =5268
                    Top =240
                    Width =768
                    Height =0
                    Name ="EmptyCell619"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5268
                    LayoutCachedTop =240
                    LayoutCachedWidth =6036
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =7
                End
                Begin EmptyCell
                    Left =6036
                    Top =240
                    Width =1212
                    Height =0
                    Name ="EmptyCell620"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6036
                    LayoutCachedTop =240
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =7
                End
                Begin EmptyCell
                    Left =13692
                    Top =240
                    Width =540
                    Height =0
                    Name ="EmptyCell629"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13692
                    LayoutCachedTop =240
                    LayoutCachedWidth =14232
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =7
                End
                Begin EmptyCell
                    Left =17148
                    Top =240
                    Width =840
                    Height =0
                    Name ="EmptyCell634"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17148
                    LayoutCachedTop =240
                    LayoutCachedWidth =17988
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    GroupTable =7
                End
                Begin EmptyCell
                    Left =17988
                    Top =240
                    Width =1080
                    Height =0
                    Name ="EmptyCell635"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17988
                    LayoutCachedTop =240
                    LayoutCachedWidth =19068
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =22
                    ColumnEnd =22
                    LayoutGroup =1
                    GroupTable =7
                End
                Begin EmptyCell
                    Left =19068
                    Top =240
                    Width =576
                    Height =0
                    Name ="EmptyCell636"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19068
                    LayoutCachedTop =240
                    LayoutCachedWidth =19644
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =23
                    ColumnEnd =23
                    LayoutGroup =1
                    GroupTable =7
                End
                Begin EmptyCell
                    Top =240
                    Width =749
                    Height =0
                    Name ="EmptyCell637"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =240
                    LayoutCachedWidth =749
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =7
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8916
                    Top =240
                    Width =1056
                    Height =0
                    FontSize =9
                    TabIndex =26
                    ForeColor =4144959
                    Name ="TermPenaltyx"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",[TermPenaltyProj],Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8916
                    LayoutCachedTop =240
                    LayoutCachedWidth =9972
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10572
                    Top =240
                    Width =1080
                    Height =0
                    FontSize =9
                    TabIndex =28
                    ForeColor =4144959
                    Name ="NetSyndFeeD"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",[NetSyndProj],Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =10572
                    LayoutCachedTop =240
                    LayoutCachedWidth =11652
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12252
                    Top =240
                    Width =900
                    Height =0
                    FontSize =9
                    TabIndex =30
                    ForeColor =4144959
                    Name ="PerDiemDelta"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",[PDRentProj],Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =12252
                    LayoutCachedTop =240
                    LayoutCachedWidth =13152
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7248
                    Top =240
                    Width =1080
                    Height =0
                    FontSize =9
                    TabIndex =24
                    ForeColor =4144959
                    Name ="SyndFeeD"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",[SyndFeeProj],Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =7248
                    LayoutCachedTop =240
                    LayoutCachedWidth =8328
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8328
                    Top =240
                    Width =588
                    Height =0
                    FontSize =8
                    TabIndex =25
                    ForeColor =8872517
                    Name ="SyndFeeV"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Format([SyndFeeVar],\"#.##%;-#.##%\"),Null)"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x0100000084000000010000000100000000000000000000001100000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00530079006e006400460065006500440065006c00740061005d003c003000 ,
                        0x00000000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8328
                    LayoutCachedTop =240
                    LayoutCachedWidth =8916
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000100ba141900ffffff00100000005b00 ,
                        0x530079006e006400460065006500440065006c00740061005d003c0030000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    GroupTable =7
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9972
                    Top =240
                    Width =600
                    Height =0
                    FontSize =8
                    TabIndex =27
                    ForeColor =8872517
                    Name ="TermPenaltyV"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Format([TermPenaltyVar],\"#.##%;-#.##%\"),Null)"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b005400650072006d00500065006e0061006c00740079005600610072005d00 ,
                        0x3c00300000000000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =9972
                    LayoutCachedTop =240
                    LayoutCachedWidth =10572
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000100ba141900ffffff00120000005b00 ,
                        0x5400650072006d00500065006e0061006c00740079005600610072005d003c00 ,
                        0x3000000000000000000000000000000000000000000000
                    End
                    GroupTable =7
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11652
                    Top =240
                    Width =600
                    Height =0
                    FontSize =8
                    TabIndex =29
                    ForeColor =8872517
                    Name ="NetSyndFeeV"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Format([NetSyndFeeVar],\"#.##%;-#.##%\"),Null)"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x010000008a000000010000000100000000000000000000001400000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004e0065007400530079006e006400460065006500440065006c0074006100 ,
                        0x5d003c00300000000000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =11652
                    LayoutCachedTop =240
                    LayoutCachedWidth =12252
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000100ba141900ffffff00130000005b00 ,
                        0x4e0065007400530079006e006400460065006500440065006c00740061005d00 ,
                        0x3c003000000000000000000000000000000000000000000000
                    End
                    GroupTable =7
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14232
                    Top =240
                    Width =768
                    Height =0
                    FontSize =9
                    TabIndex =32
                    ForeColor =4144959
                    Name ="UTAD"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",[UTAProj],Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =14232
                    LayoutCachedTop =240
                    LayoutCachedWidth =15000
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15000
                    Top =240
                    Width =495
                    Height =0
                    FontSize =8
                    TabIndex =33
                    ForeColor =8872517
                    Name ="UtAV"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Format([UtAVar],\"#.##%;-#.##%\"),Null)"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x010000007c000000010000000100000000000000000000000d00000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00550054004100440065006c00740061005d003c00300000000000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =15000
                    LayoutCachedTop =240
                    LayoutCachedWidth =15495
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000100ba141900ffffff000c0000005b00 ,
                        0x550054004100440065006c00740061005d003c00300000000000000000000000 ,
                        0x0000000000000000000000
                    End
                    GroupTable =7
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15492
                    Top =240
                    Width =1080
                    Height =0
                    FontSize =9
                    TabIndex =34
                    ForeColor =4144959
                    Name ="TotalFeeD"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",[TotalFeeProj],Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =15492
                    LayoutCachedTop =240
                    LayoutCachedWidth =16572
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =7
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16572
                    Top =240
                    Width =576
                    Height =0
                    FontSize =8
                    TabIndex =35
                    ForeColor =8872517
                    Name ="TotalFeeV"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Format([TotalFeeVar],\"#.##%;-#.##%\"),Null)"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x0100000086000000010000000100000000000000000000001200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0054006f00740061006c00460065006500440065006c00740061005d003c00 ,
                        0x300000000000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =16572
                    LayoutCachedTop =240
                    LayoutCachedWidth =17148
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000100ba141900ffffff00110000005b00 ,
                        0x54006f00740061006c00460065006500440065006c00740061005d003c003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =7
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13152
                    Top =240
                    Width =540
                    Height =0
                    FontSize =8
                    TabIndex =31
                    ForeColor =8872517
                    Name ="PDRentV"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Format([PDRentVar],\"#.##%;-#.##%\"),Null)"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x0100000082000000010000000100000000000000000000001000000001000100 ,
                        0xff000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0050004400520065006e007400440065006c00740061005d003c0030000000 ,
                        0x0000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =13152
                    LayoutCachedTop =240
                    LayoutCachedWidth =13692
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000100ff000000ffffff000f0000005b00 ,
                        0x50004400520065006e007400440065006c00740061005d003c00300000000000 ,
                        0x0000000000000000000000000000000000
                    End
                    GroupTable =7
                End
                Begin TextBox
                    Visible = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =744
                    Top =240
                    Width =1236
                    Height =0
                    FontSize =9
                    TabIndex =23
                    ForeColor =4144959
                    Name ="SGrpID"
                    ControlSource ="SGrpID"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =744
                    LayoutCachedTop =240
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =240
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =7
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =300
            BreakLevel =1
            Name ="GroupFooter0"
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =19068
                    Top =12
                    Width =576
                    Height =276
                    FontSize =8
                    TabIndex =18
                    ForeColor =8872517
                    Name ="Text371"
                    ControlSource ="=Format([ClientNetFeeCom]/[ClientOEC],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =19068
                    LayoutCachedTop =12
                    LayoutCachedWidth =19644
                    LayoutCachedHeight =288
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17988
                    Top =12
                    Width =1080
                    Height =276
                    FontSize =9
                    TabIndex =17
                    ForeColor =0
                    Name ="ClientNetFeeCom"
                    ControlSource ="=Sum([NetTotalFee])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =17988
                    LayoutCachedTop =12
                    LayoutCachedWidth =19068
                    LayoutCachedHeight =288
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16572
                    Top =12
                    Width =576
                    Height =276
                    FontSize =8
                    TabIndex =15
                    ForeColor =8872517
                    Name ="Text347"
                    ControlSource ="=Format([ClientNetFee]/[ClientOEC],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =16572
                    LayoutCachedTop =12
                    LayoutCachedWidth =17148
                    LayoutCachedHeight =288
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15492
                    Top =12
                    Width =1080
                    Height =276
                    FontSize =9
                    TabIndex =14
                    ForeColor =0
                    Name ="ClientNetFee"
                    ControlSource ="=Sum([TotalFee])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =15492
                    LayoutCachedTop =12
                    LayoutCachedWidth =16572
                    LayoutCachedHeight =288
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17148
                    Top =12
                    Width =840
                    Height =276
                    FontSize =9
                    TabIndex =16
                    BorderColor =9211020
                    ForeColor =0
                    Name ="ClientComm"
                    ControlSource ="=Sum([BDEXComm])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =17148
                    LayoutCachedTop =12
                    LayoutCachedWidth =17988
                    LayoutCachedHeight =288
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =15000
                    Top =12
                    Width =495
                    Height =276
                    FontSize =8
                    TabIndex =13
                    ForeColor =8872517
                    Name ="Text219"
                    ControlSource ="=Format([ClientUTA]/[ClientOEC],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =15000
                    LayoutCachedTop =12
                    LayoutCachedWidth =15495
                    LayoutCachedHeight =288
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =9972
                    Top =12
                    Width =600
                    Height =276
                    FontSize =8
                    TabIndex =6
                    ForeColor =8872517
                    Name ="Text255"
                    ControlSource ="=Format([ClientTermPenalty]/[ClientOEC],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =9972
                    LayoutCachedTop =12
                    LayoutCachedWidth =10572
                    LayoutCachedHeight =288
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8916
                    Top =12
                    Width =1056
                    Height =276
                    FontSize =9
                    TabIndex =5
                    BorderColor =9211020
                    ForeColor =0
                    Name ="ClientTermPenalty"
                    ControlSource ="=Sum([TermPenalty])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8916
                    LayoutCachedTop =12
                    LayoutCachedWidth =9972
                    LayoutCachedHeight =288
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14232
                    Top =12
                    Width =768
                    Height =276
                    FontSize =9
                    TabIndex =12
                    BorderColor =9211020
                    ForeColor =0
                    Name ="ClientUTA"
                    ControlSource ="=Sum([UTA])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14232
                    LayoutCachedTop =12
                    LayoutCachedWidth =15000
                    LayoutCachedHeight =288
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13692
                    Top =12
                    Width =540
                    Height =276
                    FontSize =9
                    TabIndex =11
                    ForeColor =0
                    Name ="Text215"
                    ControlSource ="=Sum([PDDaysAvg])"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13692
                    LayoutCachedTop =12
                    LayoutCachedWidth =14232
                    LayoutCachedHeight =288
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =13152
                    Top =12
                    Width =540
                    Height =276
                    FontSize =8
                    TabIndex =10
                    ForeColor =8872517
                    Name ="Text213"
                    ControlSource ="=Format([ClientPerDiemGroup]/[ClientOEC],\"Percent\")"
                    Format ="Percent"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =13152
                    LayoutCachedTop =12
                    LayoutCachedWidth =13692
                    LayoutCachedHeight =288
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12252
                    Top =12
                    Width =900
                    Height =276
                    FontSize =9
                    TabIndex =9
                    BorderColor =9211020
                    ForeColor =0
                    Name ="ClientPerDiemGroup"
                    ControlSource ="=Sum([PDRent])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12252
                    LayoutCachedTop =12
                    LayoutCachedWidth =13152
                    LayoutCachedHeight =288
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =8328
                    Top =12
                    Width =588
                    Height =276
                    FontSize =8
                    TabIndex =4
                    ForeColor =8872517
                    Name ="Text203"
                    ControlSource ="=Format([ClientGSynFee]/[ClientOEC],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8328
                    LayoutCachedTop =12
                    LayoutCachedWidth =8916
                    LayoutCachedHeight =288
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7248
                    Top =12
                    Width =1080
                    Height =276
                    FontSize =9
                    TabIndex =3
                    BorderColor =9211020
                    ForeColor =0
                    Name ="ClientGSynFee"
                    ControlSource ="=Sum([SyndFee])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7248
                    LayoutCachedTop =12
                    LayoutCachedWidth =8328
                    LayoutCachedHeight =288
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6036
                    Top =12
                    Width =1212
                    Height =276
                    FontSize =9
                    TabIndex =2
                    BorderColor =9211020
                    ForeColor =0
                    Name ="ClientOEC"
                    ControlSource ="=Sum([TotalVendorCost])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6036
                    LayoutCachedTop =12
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =288
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4500
                    Top =12
                    Width =1530
                    Height =276
                    FontSize =10
                    TabIndex =1
                    ForeColor =5855577
                    Name ="Text194"
                    ControlSource ="=\"← Totals →\""
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4500
                    LayoutCachedTop =12
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =288
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4008
                    Top =12
                    Width =492
                    Height =276
                    FontSize =9
                    FontWeight =700
                    BorderColor =9211020
                    ForeColor =0
                    Name ="Text193"
                    ControlSource ="=Sum([UnitCount])"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4008
                    LayoutCachedTop =12
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =288
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10572
                    Top =12
                    Width =1080
                    Height =276
                    FontSize =9
                    TabIndex =7
                    BorderColor =9211020
                    ForeColor =0
                    Name ="ClientNetSyndFee"
                    ControlSource ="=Sum([NetSyndFee])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10572
                    LayoutCachedTop =12
                    LayoutCachedWidth =11652
                    LayoutCachedHeight =288
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =11652
                    Top =12
                    Width =600
                    Height =276
                    FontSize =8
                    TabIndex =8
                    ForeColor =8872517
                    Name ="Text293"
                    ControlSource ="=Format([ClientNetSyndFee]/[ClientOEC],\"#.##%;[Red]#.##%\")"
                    Format ="Percent"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =11652
                    LayoutCachedTop =12
                    LayoutCachedWidth =12252
                    LayoutCachedHeight =288
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin Rectangle
                    OldBorderStyle =0
                    Left =15
                    Width =19512
                    Height =29
                    BackColor =9605778
                    Name ="Box185"
                    LayoutCachedLeft =15
                    LayoutCachedWidth =19527
                    LayoutCachedHeight =29
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8916
                    Top =288
                    Width =1056
                    Height =12
                    FontSize =9
                    TabIndex =22
                    ForeColor =4144959
                    Name ="ClientTermPenaltyLO"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Sum([TermPenaltyProj]),Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8916
                    LayoutCachedTop =288
                    LayoutCachedWidth =9972
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10572
                    Top =288
                    Width =1080
                    Height =12
                    FontSize =9
                    TabIndex =24
                    ForeColor =4144959
                    Name ="ClientNetSyndFeeLO"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Sum([NetSyndProj]),Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =10572
                    LayoutCachedTop =288
                    LayoutCachedWidth =11652
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12252
                    Top =288
                    Width =900
                    Height =12
                    FontSize =9
                    TabIndex =26
                    ForeColor =4144959
                    Name ="ClientPerDiemGroupLO"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Sum([PDRentProj]),Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =12252
                    LayoutCachedTop =288
                    LayoutCachedWidth =13152
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7248
                    Top =288
                    Width =1080
                    Height =12
                    FontSize =9
                    TabIndex =20
                    ForeColor =4144959
                    Name ="ClientGLOFee"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Sum([SyndFeeProj]),Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =7248
                    LayoutCachedTop =288
                    LayoutCachedWidth =8328
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14232
                    Top =288
                    Width =768
                    Height =12
                    FontSize =9
                    TabIndex =28
                    ForeColor =4144959
                    Name ="ClientUTALo"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Sum([UTAProj]),Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =14232
                    LayoutCachedTop =288
                    LayoutCachedWidth =15000
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15492
                    Top =288
                    Width =1080
                    Height =12
                    FontSize =9
                    TabIndex =30
                    ForeColor =4144959
                    Name ="ClientNetFeeLO"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Sum([TotalFeeProj]),Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =15492
                    LayoutCachedTop =288
                    LayoutCachedWidth =16572
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =12
                    CurrencySymbol ="$"
                End
                Begin EmptyCell
                    Left =4008
                    Top =288
                    Width =492
                    Height =12
                    Name ="EmptyCell758"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4008
                    LayoutCachedTop =288
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    GroupTable =12
                End
                Begin EmptyCell
                    Left =4500
                    Top =288
                    Width =1530
                    Height =12
                    Name ="EmptyCell759"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4500
                    LayoutCachedTop =288
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =12
                End
                Begin EmptyCell
                    Left =13692
                    Top =288
                    Width =540
                    Height =12
                    Name ="EmptyCell769"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13692
                    LayoutCachedTop =288
                    LayoutCachedWidth =14232
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =2
                    GroupTable =12
                End
                Begin EmptyCell
                    Left =17148
                    Top =288
                    Width =840
                    Height =12
                    Name ="EmptyCell774"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17148
                    LayoutCachedTop =288
                    LayoutCachedWidth =17988
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =2
                    GroupTable =12
                End
                Begin EmptyCell
                    Left =17988
                    Top =288
                    Width =1080
                    Height =12
                    Name ="EmptyCell775"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17988
                    LayoutCachedTop =288
                    LayoutCachedWidth =19068
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =2
                    GroupTable =12
                End
                Begin EmptyCell
                    Left =19068
                    Top =288
                    Width =576
                    Height =12
                    Name ="EmptyCell776"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19068
                    LayoutCachedTop =288
                    LayoutCachedWidth =19644
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =2
                    GroupTable =12
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =8328
                    Top =288
                    Width =588
                    Height =12
                    FontSize =8
                    TabIndex =21
                    ForeColor =8872517
                    Name ="Text782"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Format([ClientGLOFee]/[ClientOEC],\"#.##%;[Red]#"
                        ".##%\"),Null)"
                    Format ="Percent"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8328
                    LayoutCachedTop =288
                    LayoutCachedWidth =8916
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =9972
                    Top =288
                    Width =600
                    Height =12
                    FontSize =8
                    TabIndex =23
                    ForeColor =8872517
                    Name ="Text784"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Format([ClientTermPenaltyLO]/[ClientOEC],\"#.##%"
                        ";[Red]#.##%\"),Null)"
                    Format ="Percent"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =9972
                    LayoutCachedTop =288
                    LayoutCachedWidth =10572
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =11652
                    Top =288
                    Width =600
                    Height =12
                    FontSize =8
                    TabIndex =25
                    ForeColor =8872517
                    Name ="Text785"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Format([ClientNetSyndFeeLO]/[ClientOEC],\"#.##%;"
                        "[Red]#.##%\"),Null)"
                    Format ="Percent"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =11652
                    LayoutCachedTop =288
                    LayoutCachedWidth =12252
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =13152
                    Top =288
                    Width =540
                    Height =12
                    FontSize =8
                    TabIndex =27
                    ForeColor =8872517
                    Name ="Text786"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Format([ClientPerDiemGroupLO]/[ClientOEC],\"#.##"
                        "%;[Red]#.##%\"),Null)"
                    Format ="Percent"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =13152
                    LayoutCachedTop =288
                    LayoutCachedWidth =13692
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =15000
                    Top =288
                    Width =495
                    Height =12
                    FontSize =8
                    TabIndex =29
                    ForeColor =8872517
                    Name ="Text787"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Format([ClientUTALo]/[ClientOEC],\"#.##%;[Red]#."
                        "##%\"),Null)"
                    Format ="Percent"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =15000
                    LayoutCachedTop =288
                    LayoutCachedWidth =15495
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =16572
                    Top =288
                    Width =576
                    Height =12
                    FontSize =8
                    TabIndex =31
                    ForeColor =8872517
                    Name ="Text792"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",Format([ClientNetFeeLO]/[ClientOEC],\"#.##%;[Red"
                        "]#.##%\"),Null)"
                    Format ="Percent"
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =16572
                    LayoutCachedTop =288
                    LayoutCachedWidth =17148
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =12
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6036
                    Top =288
                    Width =1212
                    Height =12
                    FontSize =9
                    TabIndex =19
                    ForeColor =4144959
                    Name ="Text794"
                    ControlSource ="=IIf([Assigned_Indicator]=\"F\",\"LO Totals →\",Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =12
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =6036
                    LayoutCachedTop =288
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =12
                    CurrencySymbol ="$"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =15
            Name ="GroupFooter1"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin PageBreak
                    Top =15
                    Name ="PageBreak748"
                End
            End
        End
        Begin PageFooter
            Height =255
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14250
                    Width =5130
                    Height =225
                    FontSize =8
                    BorderColor =0
                    ForeColor =0
                    Name ="Text30"
                    ControlSource ="=Date() & \"          Page \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =0

                    LayoutCachedLeft =14250
                    LayoutCachedWidth =19380
                    LayoutCachedHeight =225
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5775
                    Width =7215
                    Height =255
                    FontSize =10
                    TabIndex =1
                    BackColor =12566463
                    ForeColor =0
                    Name ="Text749"
                    ControlSource ="=\"<B>\" & [FundingStatus] & \"</b>\""
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =5775
                    LayoutCachedWidth =12990
                    LayoutCachedHeight =255
                    ColumnEnd =2
                    BackShade =75.0
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Width =7545
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    BorderColor =0
                    ForeColor =0
                    Name ="Text70"
                    ControlSource ="=\"© \" & Format(Date(),\"yyyy\") & \" Fleet Advanatge | Internal Use Only\""
                    GridlineColor =0

                    LayoutCachedWidth =7545
                    LayoutCachedHeight =225
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "TFRCompare.cls"
