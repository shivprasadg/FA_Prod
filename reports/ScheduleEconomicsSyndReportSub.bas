Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15294
    DatasheetFontHeight =11
    ItemSuffix =3261
    Left =4560
    Top =4440
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x06fc913b3050e540
    End
    RecordSource ="SELECT vw_EconoProfileSchedule.* FROM vw_EconoProfileSchedule;"
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
        Begin Image
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
            GroupHeader = NotDefault
            ControlSource ="Client"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Schedule"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="GroupID"
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
            Height =0
            Name ="GroupHeader0"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =960
            BreakLevel =1
            Name ="GroupHeader2"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    Left =30
                    Top =405
                    Width =15264
                    Height =555
                    Name ="Box532"
                    LayoutCachedLeft =30
                    LayoutCachedTop =405
                    LayoutCachedWidth =15294
                    LayoutCachedHeight =960
                    BackThemeColorIndex =0
                    BackTint =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =90
                    Top =75
                    Width =14970
                    Height =330
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="GroupHeader"
                    ControlSource ="=\"Summary For Schedule \" & [Schedule] & \" - \" & Sum([Unitcount]) & \" Units\""
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =90
                    LayoutCachedTop =75
                    LayoutCachedWidth =15060
                    LayoutCachedHeight =405
                    ThemeFontIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =672
                    Top =420
                    Width =1308
                    Height =480
                    FontSize =9
                    Name ="Label7"
                    Caption ="Cost \015\012(grouped)"
                    FontName ="Segoe UI"
                    GroupTable =79
                    LayoutCachedLeft =672
                    LayoutCachedTop =420
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =900
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =79
                End
                Begin Label
                    TextAlign =2
                    Left =3540
                    Top =420
                    Width =1308
                    Height =480
                    FontSize =9
                    Name ="Label8"
                    Caption ="Synd Sale Price\015\012(grouped)"
                    FontName ="Segoe UI"
                    GroupTable =79
                    LayoutCachedLeft =3540
                    LayoutCachedTop =420
                    LayoutCachedWidth =4848
                    LayoutCachedHeight =900
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =79
                End
                Begin Label
                    TextAlign =2
                    Left =36
                    Top =420
                    Width =576
                    Height =480
                    FontSize =9
                    Name ="Label9"
                    Caption ="# Units"
                    FontName ="Segoe UI"
                    GroupTable =79
                    LayoutCachedLeft =36
                    LayoutCachedTop =420
                    LayoutCachedWidth =612
                    LayoutCachedHeight =900
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =79
                End
                Begin Label
                    TextAlign =2
                    Left =8172
                    Top =420
                    Width =1080
                    Height =480
                    FontSize =9
                    Name ="Label16"
                    Caption ="Payment\015\012(grouped)"
                    FontName ="Segoe UI"
                    GroupTable =79
                    LayoutCachedLeft =8172
                    LayoutCachedTop =420
                    LayoutCachedWidth =9252
                    LayoutCachedHeight =900
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =79
                End
                Begin Label
                    TextAlign =2
                    Left =9312
                    Top =420
                    Width =1188
                    Height =480
                    FontSize =9
                    Name ="Label19"
                    Caption ="RV $\015\012(grouped)"
                    FontName ="Segoe UI"
                    GroupTable =79
                    LayoutCachedLeft =9312
                    LayoutCachedTop =420
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =900
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =79
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13380
                    Width =765
                    FontSize =9
                    Name ="SchID"
                    ControlSource ="SchID"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13380
                    LayoutCachedWidth =14145
                    LayoutCachedHeight =240
                    ThemeFontIndex =-1
                End
                Begin Label
                    TextAlign =2
                    Left =12804
                    Top =420
                    Width =1296
                    Height =480
                    FontSize =9
                    Name ="Label26"
                    Caption ="Fee \015\012(grouped)"
                    FontName ="Segoe UI"
                    GroupTable =79
                    LayoutCachedLeft =12804
                    LayoutCachedTop =420
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =900
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =79
                End
                Begin Label
                    TextAlign =2
                    Left =11412
                    Top =420
                    Width =1332
                    Height =480
                    FontSize =9
                    Name ="Label878"
                    FontName ="Segoe UI"
                    GroupTable =79
                    LayoutCachedLeft =11412
                    LayoutCachedTop =420
                    LayoutCachedWidth =12744
                    LayoutCachedHeight =900
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =79
                End
                Begin Label
                    TextAlign =2
                    Left =7200
                    Top =420
                    Width =912
                    Height =480
                    FontSize =9
                    Name ="Label14"
                    Caption ="LRF"
                    FontName ="Segoe UI"
                    GroupTable =79
                    LayoutCachedLeft =7200
                    LayoutCachedTop =420
                    LayoutCachedWidth =8112
                    LayoutCachedHeight =900
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =79
                End
                Begin EmptyCell
                    Left =4908
                    Top =420
                    Width =2232
                    Height =480
                    Name ="EmptyCell2553"
                    GroupTable =79
                    LayoutCachedLeft =4908
                    LayoutCachedTop =420
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =900
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    GroupTable =79
                End
                Begin Label
                    TextAlign =2
                    Left =10560
                    Top =420
                    Width =792
                    Height =480
                    FontSize =9
                    Name ="Label18"
                    Caption ="RV %"
                    FontName ="Segoe UI"
                    GroupTable =79
                    LayoutCachedLeft =10560
                    LayoutCachedTop =420
                    LayoutCachedWidth =11352
                    LayoutCachedHeight =900
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =79
                End
                Begin Label
                    TextAlign =3
                    Left =14160
                    Top =420
                    Width =768
                    Height =480
                    FontSize =9
                    Name ="Label3239"
                    Caption ="Fee %"
                    FontName ="Segoe UI"
                    GroupTable =79
                    LayoutCachedLeft =14160
                    LayoutCachedTop =420
                    LayoutCachedWidth =14928
                    LayoutCachedHeight =900
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =79
                End
                Begin Label
                    TextAlign =2
                    Left =2040
                    Top =420
                    Width =1440
                    Height =480
                    FontSize =9
                    Name ="Label3251"
                    Caption ="Vendor Cost \015\012(grouped)"
                    FontName ="Segoe UI"
                    GroupTable =79
                    LayoutCachedLeft =2040
                    LayoutCachedTop =420
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =900
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =79
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            BreakLevel =2
            Name ="GroupHeader1"
            AlternateBackColor =16777215
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =0
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =354
            BreakLevel =2
            Name ="GroupFooter0"
            AlternateBackColor =16777215
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =672
                    Top =36
                    Width =1308
                    Height =288
                    FontSize =9
                    TabIndex =1
                    BorderColor =10040879
                    Name ="GroupCost"
                    ControlSource ="OECGroup"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =81

                    LayoutCachedLeft =672
                    LayoutCachedTop =36
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =324
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    GroupTable =81
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =36
                    Width =576
                    Height =288
                    FontSize =9
                    FontWeight =700
                    BorderColor =10040879
                    Name ="UnitCountLine"
                    ControlSource ="UnitCount"
                    FontName ="Segoe UI"
                    GroupTable =81

                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =612
                    LayoutCachedHeight =324
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    GroupTable =81
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9180
                    Top =36
                    Width =1188
                    Height =288
                    FontSize =9
                    TabIndex =7
                    BorderColor =10040879
                    Name ="GroupResidual"
                    ControlSource ="RVGroup"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =81

                    LayoutCachedLeft =9180
                    LayoutCachedTop =36
                    LayoutCachedWidth =10368
                    LayoutCachedHeight =324
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    GroupTable =81
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3408
                    Top =36
                    Width =1308
                    Height =288
                    FontSize =9
                    TabIndex =3
                    BorderColor =10040879
                    Name ="SyndSalePrice"
                    ControlSource ="SalePrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =81

                    LayoutCachedLeft =3408
                    LayoutCachedTop =36
                    LayoutCachedWidth =4716
                    LayoutCachedHeight =324
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    GroupTable =81
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8040
                    Top =36
                    Width =1080
                    Height =288
                    FontSize =9
                    TabIndex =6
                    BorderColor =10040879
                    Name ="LeasePaymentForGroup"
                    ControlSource ="PaymentGroup"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =81

                    LayoutCachedLeft =8040
                    LayoutCachedTop =36
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =324
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    GroupTable =81
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13944
                    Top =36
                    Width =888
                    Height =288
                    FontSize =9
                    TabIndex =10
                    ForeColor =8355711
                    Name ="FeePct"
                    ControlSource ="=[FeeUnitLine]/[OECGroup]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =81

                    LayoutCachedLeft =13944
                    LayoutCachedTop =36
                    LayoutCachedWidth =14832
                    LayoutCachedHeight =324
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    ForeTint =50.0
                    GroupTable =81
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12552
                    Top =36
                    Width =1332
                    Height =288
                    FontSize =9
                    TabIndex =9
                    BorderColor =10040879
                    Name ="FeeUnitLine"
                    ControlSource ="SyndFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =81

                    LayoutCachedLeft =12552
                    LayoutCachedTop =36
                    LayoutCachedWidth =13884
                    LayoutCachedHeight =324
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    GroupTable =81
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =4
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7068
                    Top =36
                    Width =912
                    Height =288
                    FontSize =9
                    TabIndex =5
                    Name ="Text2549"
                    ControlSource ="=[PaymentGroup]/[OECGroup]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =81

                    LayoutCachedLeft =7068
                    LayoutCachedTop =36
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =324
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =81
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10428
                    Top =36
                    Width =792
                    Height =288
                    FontSize =9
                    TabIndex =8
                    Name ="Text2582"
                    ControlSource ="=([GroupResidual])/[GroupCost]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =81

                    LayoutCachedLeft =10428
                    LayoutCachedTop =36
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =324
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =81
                End
                Begin EmptyCell
                    Left =11280
                    Top =36
                    Width =1212
                    Height =288
                    Name ="EmptyCell2584"
                    GroupTable =81
                    LayoutCachedLeft =11280
                    LayoutCachedTop =36
                    LayoutCachedWidth =12492
                    LayoutCachedHeight =324
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    GroupTable =81
                End
                Begin Label
                    TextAlign =1
                    Left =11160
                    Top =45
                    Width =1320
                    Height =240
                    FontSize =9
                    Name ="Label1487"
                    Caption ="──────────────────────"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =11160
                    LayoutCachedTop =45
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =285
                    ColumnEnd =5
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4776
                    Top =36
                    Width =2232
                    Height =288
                    FontSize =9
                    TabIndex =4
                    ForeColor =8355711
                    Name ="Label1153"
                    ControlSource ="=\"────────────────── Group \" & [GroupNum] & \" Subtotals ──────────────────\""
                    FontName ="Segoe UI"
                    GroupTable =81

                    LayoutCachedLeft =4776
                    LayoutCachedTop =36
                    LayoutCachedWidth =7008
                    LayoutCachedHeight =324
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =50.0
                    GroupTable =81
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2040
                    Top =36
                    Width =1308
                    Height =288
                    FontSize =9
                    TabIndex =2
                    BorderColor =10040879
                    Name ="Text3254"
                    ControlSource ="OECVendorGroup"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =81

                    LayoutCachedLeft =2040
                    LayoutCachedTop =36
                    LayoutCachedWidth =3348
                    LayoutCachedHeight =324
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    GroupTable =81
                    CurrencySymbol ="$"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =4524
            BreakLevel =1
            Name ="GroupFooter3"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =1
                    Left =120
                    Top =2340
                    Width =3135
                    Height =210
                    FontSize =8
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label1156"
                    Caption =" PREPARED BY"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =2340
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =2550
                    RowStart =1
                    RowEnd =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =3240
                    Top =2340
                    Width =1245
                    Height =210
                    FontSize =8
                    FontWeight =700
                    BorderColor =1643706
                    ForeColor =1643706
                    Name ="Label1157"
                    Caption ="DATE"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3240
                    LayoutCachedTop =2340
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =2550
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =120
                    Top =2910
                    Width =3135
                    Height =240
                    FontSize =8
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label1158"
                    Caption =" VERIFIED BY"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =2910
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =3150
                    RowStart =3
                    RowEnd =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =3240
                    Top =2910
                    Width =1245
                    Height =240
                    FontSize =8
                    FontWeight =700
                    BorderColor =1643706
                    ForeColor =1643706
                    Name ="Label1159"
                    Caption ="DATE"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3240
                    LayoutCachedTop =2910
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =3150
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =2685
                    Top =3150
                    Width =1920
                    Height =360
                    FontSize =8
                    FontWeight =700
                    BorderColor =1643706
                    ForeColor =1643706
                    Name ="Label1172"
                    Caption =" "
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2685
                    LayoutCachedTop =3150
                    LayoutCachedWidth =4605
                    LayoutCachedHeight =3510
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =60
                    Top =2550
                    Width =2625
                    Height =360
                    FontSize =8
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label1173"
                    Caption =" "
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =60
                    LayoutCachedTop =2550
                    LayoutCachedWidth =2685
                    LayoutCachedHeight =2910
                    RowStart =2
                    RowEnd =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =2685
                    Top =2550
                    Width =1920
                    Height =360
                    FontSize =8
                    FontWeight =700
                    BorderColor =1643706
                    ForeColor =1643706
                    Name ="Label1175"
                    Caption =" "
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2685
                    LayoutCachedTop =2550
                    LayoutCachedWidth =4605
                    LayoutCachedHeight =2910
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =120
                    Top =3150
                    Width =3135
                    Height =360
                    FontSize =8
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label1177"
                    Caption =" "
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =3150
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =3510
                    RowStart =4
                    RowEnd =4
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =60
                    Top =1980
                    Width =4545
                    Height =360
                    FontWeight =700
                    BackColor =1643706
                    BorderColor =1643706
                    Name ="Label1188"
                    Caption =" Syndication Approval"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =60
                    LayoutCachedTop =1980
                    LayoutCachedWidth =4605
                    LayoutCachedHeight =2340
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Rectangle
                    BackStyle =0
                    Left =60
                    Top =1980
                    Width =4545
                    Height =915
                    BorderColor =1643706
                    Name ="Box1192"
                    LayoutCachedLeft =60
                    LayoutCachedTop =1980
                    LayoutCachedWidth =4605
                    LayoutCachedHeight =2895
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    BackStyle =0
                    Left =60
                    Top =2894
                    Width =4545
                    Height =1260
                    BorderColor =1643706
                    Name ="Box1194"
                    LayoutCachedLeft =60
                    LayoutCachedTop =2894
                    LayoutCachedWidth =4605
                    LayoutCachedHeight =4154
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =2565
                    Width =3135
                    Height =315
                    FontSize =9
                    TabIndex =27
                    ForeColor =0
                    Name ="FullUserName"
                    ControlSource ="CurrentUser"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =120
                    LayoutCachedTop =2565
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =2880
                    ThemeFontIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3240
                    Top =2565
                    Width =1335
                    Height =330
                    FontSize =9
                    TabIndex =28
                    ForeColor =0
                    Name ="DateRan"
                    ControlSource ="=Date()"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =3240
                    LayoutCachedTop =2565
                    LayoutCachedWidth =4575
                    LayoutCachedHeight =2895
                    ThemeFontIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =684
                    Top =60
                    Width =1308
                    Height =285
                    FontSize =9
                    TabIndex =1
                    BorderColor =10040879
                    Name ="ScheduleOECCost"
                    ControlSource ="=Sum([OECGroup])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =71

                    LayoutCachedLeft =684
                    LayoutCachedTop =60
                    LayoutCachedWidth =1992
                    LayoutCachedHeight =345
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    GroupTable =71
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =60
                    Width =576
                    Height =285
                    FontSize =9
                    FontWeight =700
                    BorderColor =10040879
                    Name ="Text1593"
                    ControlSource ="=Sum([Unitcount])"
                    FontName ="Segoe UI"
                    GroupTable =71

                    LayoutCachedLeft =36
                    LayoutCachedTop =60
                    LayoutCachedWidth =612
                    LayoutCachedHeight =345
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    GroupTable =71
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9240
                    Top =60
                    Width =1188
                    Height =285
                    FontSize =9
                    TabIndex =5
                    BorderColor =10040879
                    Name ="ResGroup"
                    ControlSource ="=Sum([RVGroup])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =71

                    LayoutCachedLeft =9240
                    LayoutCachedTop =60
                    LayoutCachedWidth =10428
                    LayoutCachedHeight =345
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    GroupTable =71
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12600
                    Top =60
                    Width =1332
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =7
                    BorderColor =10040879
                    Name ="FeeGroupLine"
                    ControlSource ="=Sum([SyndFee])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =71

                    LayoutCachedLeft =12600
                    LayoutCachedTop =60
                    LayoutCachedWidth =13932
                    LayoutCachedHeight =345
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    GroupTable =71
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =11340
                            Top =60
                            Width =1188
                            Height =285
                            FontSize =9
                            FontWeight =700
                            ForeColor =10040879
                            Name ="Label3257"
                            Caption ="Synd Totals"
                            FontName ="Segoe UI"
                            GroupTable =71
                            LayoutCachedLeft =11340
                            LayoutCachedTop =60
                            LayoutCachedWidth =12528
                            LayoutCachedHeight =345
                            ColumnStart =8
                            ColumnEnd =8
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =71
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13992
                    Top =60
                    Width =888
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =8
                    ForeColor =8355711
                    Name ="Text1597"
                    ControlSource ="=[FeeGroupLine]/[ScheduleCost]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =71

                    LayoutCachedLeft =13992
                    LayoutCachedTop =60
                    LayoutCachedWidth =14880
                    LayoutCachedHeight =345
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =50.0
                    GroupTable =71
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2064
                    Top =60
                    Width =1308
                    Height =285
                    FontSize =9
                    TabIndex =2
                    BorderColor =10040879
                    Name ="ScheduleCost"
                    ControlSource ="=Sum([OECVendorGroup])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =71

                    LayoutCachedLeft =2064
                    LayoutCachedTop =60
                    LayoutCachedWidth =3372
                    LayoutCachedHeight =345
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    GroupTable =71
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8088
                    Top =60
                    Width =1080
                    Height =285
                    FontSize =9
                    TabIndex =4
                    BorderColor =10040879
                    Name ="Text1599"
                    ControlSource ="=Sum([PaymentGroup])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =71

                    LayoutCachedLeft =8088
                    LayoutCachedTop =60
                    LayoutCachedWidth =9168
                    LayoutCachedHeight =345
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    GroupTable =71
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =4
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7104
                    Top =60
                    Width =912
                    Height =285
                    FontSize =9
                    TabIndex =3
                    Name ="Text1603"
                    ControlSource ="=Sum([PaymentGroup])/[ScheduleCost]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =71

                    LayoutCachedLeft =7104
                    LayoutCachedTop =60
                    LayoutCachedWidth =8016
                    LayoutCachedHeight =345
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =71
                    Begin
                        Begin Label
                            TextAlign =2
                            Left =3432
                            Top =60
                            Width =3612
                            Height =285
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label1600"
                            Caption ="────────────────────── Schedule Totals ──────────────────────"
                            FontName ="Segoe UI"
                            GroupTable =71
                            LayoutCachedLeft =3432
                            LayoutCachedTop =60
                            LayoutCachedWidth =7044
                            LayoutCachedHeight =345
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =71
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12525
                    Top =1080
                    Width =1335
                    FontSize =9
                    TabIndex =14
                    ForeColor =1643706
                    Name ="SyndPenaltyLine"
                    ControlSource ="=Sum([TermPenaltyActual])"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="amount/expenses normally reducing the syndication fee - offent money used to sup"
                        "port Remarketing"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12525
                    LayoutCachedTop =1080
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =1320
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =9944
                            Top =1080
                            Width =2520
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label1684"
                            Caption ="- Termination Penalty"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =9944
                            LayoutCachedTop =1080
                            LayoutCachedWidth =12464
                            LayoutCachedHeight =1320
                            RowStart =3
                            RowEnd =3
                            ColumnStart =11
                            ColumnEnd =11
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12525
                    Top =1440
                    Width =1335
                    FontSize =9
                    FontWeight =700
                    TabIndex =16
                    Name ="BankFeeLine"
                    ControlSource ="=(Sum([SyndFee])+Sum([StripRentActual]))+Sum([TermPenaltyActual])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12525
                    LayoutCachedTop =1440
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =1680
                    RowStart =6
                    RowEnd =6
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =11279
                            Top =1440
                            Width =1185
                            Height =285
                            FontSize =9
                            FontWeight =700
                            ForeColor =10040879
                            Name ="Label1686"
                            Caption ="Net Synd Fee"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =11279
                            LayoutCachedTop =1440
                            LayoutCachedWidth =12464
                            LayoutCachedHeight =1725
                            RowStart =6
                            RowEnd =6
                            ColumnStart =11
                            ColumnEnd =11
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12525
                    Top =1725
                    Width =1335
                    FontSize =9
                    TabIndex =19
                    Name ="UATLine"
                    ControlSource ="=Sum([UTAGroup])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12525
                    LayoutCachedTop =1725
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =1965
                    RowStart =4
                    RowEnd =4
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13935
                    Top =1725
                    Width =885
                    FontSize =9
                    TabIndex =20
                    ForeColor =8355711
                    Name ="Text1688"
                    ControlSource ="=IIf([UATLine]>0,[UATLine]/[ScheduleCost],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13935
                    LayoutCachedTop =1725
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =1965
                    RowStart =4
                    RowEnd =4
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12525
                    Top =2025
                    Width =1335
                    FontSize =9
                    TabIndex =22
                    Name ="PerdiemLine"
                    ControlSource ="=Sum([PerdiemActive])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12525
                    LayoutCachedTop =2025
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =2265
                    RowStart =5
                    RowEnd =5
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13935
                    Top =2025
                    Width =885
                    FontSize =9
                    TabIndex =23
                    ForeColor =8355711
                    Name ="Text1691"
                    ControlSource ="=IIf([PerdiemLine]>0,[PerdiemLine]/[ScheduleCost],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13935
                    LayoutCachedTop =2025
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =2265
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12525
                    Top =2970
                    Width =1335
                    FontSize =9
                    FontWeight =700
                    TabIndex =32
                    Name ="TotalFeeLine"
                    ControlSource ="=[BankFeeLine]+[UATLine]+[PerdiemLine]+[ServFeePV]+[InterimLine]"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12525
                    LayoutCachedTop =2970
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =3210
                    RowStart =8
                    RowEnd =8
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =10785
                            Top =2970
                            Width =1680
                            Height =285
                            FontSize =9
                            FontWeight =700
                            ForeColor =10040879
                            Name ="Label1693"
                            Caption ="Schedule Sub Total"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =10785
                            LayoutCachedTop =2970
                            LayoutCachedWidth =12465
                            LayoutCachedHeight =3255
                            RowStart =8
                            RowEnd =8
                            ColumnStart =11
                            ColumnEnd =11
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13935
                    Top =2970
                    Width =885
                    FontSize =9
                    FontWeight =700
                    TabIndex =33
                    ForeColor =8355711
                    Name ="Text1694"
                    ControlSource ="=[TotalFeeLine]/[ScheduleCost]"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13935
                    LayoutCachedTop =2970
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =3210
                    RowStart =8
                    RowEnd =8
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin Rectangle
                    Left =10140
                    Top =2910
                    Width =4680
                    Height =15
                    Name ="Box1695"
                    LayoutCachedLeft =10140
                    LayoutCachedTop =2910
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =2925
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6959
                    Top =1725
                    Width =5505
                    FontSize =9
                    TabIndex =18
                    ForeColor =10040879
                    Name ="Text1696"
                    ControlSource ="[MsgUTA]"
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedLeft =6959
                    LayoutCachedTop =1725
                    LayoutCachedWidth =12464
                    LayoutCachedHeight =1965
                    RowStart =4
                    RowEnd =4
                    ColumnStart =15
                    ColumnEnd =15
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12525
                    Top =780
                    Width =1335
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =12
                    Name ="GrossFeeLine"
                    ControlSource ="=Sum([SyndFee])+Sum([StripRentActual])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12525
                    LayoutCachedTop =780
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =1065
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin Rectangle
                    Left =12593
                    Top =1365
                    Width =1267
                    Height =15
                    Name ="Box1698"
                    LayoutCachedLeft =12593
                    LayoutCachedTop =1365
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =1380
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13935
                    Top =1080
                    Width =885
                    FontSize =9
                    TabIndex =15
                    ForeColor =8355711
                    Name ="Text1700"
                    ControlSource ="=IIf([SyndPenaltyLine]<>0,[SyndPenaltyLine]/[ScheduleCost],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13935
                    LayoutCachedTop =1080
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =1320
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13935
                    Top =780
                    Width =885
                    FontSize =9
                    FontWeight =700
                    TabIndex =13
                    ForeColor =8355711
                    Name ="Text1701"
                    ControlSource ="=[GrossFeeLine]/[ScheduleCost]"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13935
                    LayoutCachedTop =780
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =1020
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13935
                    Top =1440
                    Width =885
                    FontSize =9
                    FontWeight =700
                    TabIndex =17
                    ForeColor =8355711
                    Name ="Text1702"
                    ControlSource ="=[BankFeeLine]/[ScheduleCost]"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13935
                    LayoutCachedTop =1440
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =1680
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10488
                    Top =60
                    Width =792
                    Height =285
                    FontSize =9
                    TabIndex =6
                    Name ="ResGroupPct"
                    ControlSource ="=Sum([RVGroup])/[ScheduleCost]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =71

                    LayoutCachedLeft =10488
                    LayoutCachedTop =60
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =345
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =71
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12525
                    Top =420
                    Width =1335
                    FontSize =9
                    TabIndex =10
                    Name ="StripsLine"
                    ControlSource ="=Sum([StripRentActual])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12525
                    LayoutCachedTop =420
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =660
                    RowStart =5
                    RowEnd =5
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13935
                    Top =420
                    Width =885
                    FontSize =9
                    TabIndex =11
                    ForeColor =8355711
                    Name ="Text1908"
                    ControlSource ="=IIf([StripsLine]>0,[StripsLine]/[ScheduleCost],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13935
                    LayoutCachedTop =420
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =660
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6959
                    Top =420
                    Width =5505
                    Height =300
                    FontSize =9
                    TabIndex =9
                    ForeColor =10040879
                    Name ="Text1909"
                    ControlSource ="=[MsgStrips]"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =6959
                    LayoutCachedTop =420
                    LayoutCachedWidth =12464
                    LayoutCachedHeight =720
                    RowStart =5
                    RowEnd =5
                    ColumnStart =11
                    ColumnEnd =11
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =3
                    Left =10784
                    Top =780
                    Width =1680
                    Height =285
                    FontSize =9
                    FontWeight =700
                    ForeColor =10040879
                    Name ="Text1699"
                    Caption ="Synd Fee + Strip(s)"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =10784
                    LayoutCachedTop =780
                    LayoutCachedWidth =12464
                    LayoutCachedHeight =1065
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =11
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Rectangle
                    Left =12593
                    Top =720
                    Width =1267
                    Height =15
                    Name ="Box1910"
                    LayoutCachedLeft =12593
                    LayoutCachedTop =720
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =735
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12525
                    Top =2625
                    Width =1335
                    FontSize =9
                    TabIndex =30
                    Name ="ServFeePV2"
                    ControlSource ="=Sum([ServFeePV])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12525
                    LayoutCachedTop =2625
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =2865
                    RowStart =5
                    RowEnd =5
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13935
                    Top =2625
                    Width =885
                    FontSize =9
                    TabIndex =31
                    ForeColor =8355711
                    Name ="Text2061"
                    ControlSource ="=IIf([ServFeePV2]>0,[ServFeePV2]/[ScheduleCost],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13935
                    LayoutCachedTop =2625
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =2865
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12525
                    Top =2325
                    Width =1335
                    FontSize =9
                    TabIndex =25
                    ForeColor =1643706
                    Name ="InterimLine"
                    ControlSource ="=Sum([InterimExpenseGroup])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12525
                    LayoutCachedTop =2325
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =2565
                    RowStart =5
                    RowEnd =5
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13935
                    Top =2325
                    Width =885
                    FontSize =9
                    TabIndex =26
                    ForeColor =1643706
                    Name ="Text3219"
                    ControlSource ="=IIf([InterimLine]<>0,[InterimLine]/[ScheduleCost],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13935
                    LayoutCachedTop =2325
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =2565
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =120
                    Top =3525
                    Width =3135
                    Height =240
                    FontSize =8
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label3220"
                    Caption =" APPROVED  BY"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =3525
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =3765
                    RowStart =3
                    RowEnd =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =3240
                    Top =3525
                    Width =1245
                    Height =240
                    FontSize =8
                    FontWeight =700
                    BorderColor =1643706
                    ForeColor =1643706
                    Name ="Label3221"
                    Caption ="DATE"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3240
                    LayoutCachedTop =3525
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =3765
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =2685
                    Top =3765
                    Width =1920
                    Height =360
                    FontSize =8
                    FontWeight =700
                    BorderColor =1643706
                    ForeColor =1643706
                    Name ="Label3222"
                    Caption =" "
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2685
                    LayoutCachedTop =3765
                    LayoutCachedWidth =4605
                    LayoutCachedHeight =4125
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =120
                    Top =3765
                    Width =3135
                    Height =360
                    FontSize =8
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label3223"
                    Caption =" "
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =3765
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =4125
                    RowStart =4
                    RowEnd =4
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Rectangle
                    BackStyle =0
                    Left =60
                    Top =3494
                    Width =4545
                    Height =660
                    BorderColor =1643706
                    Name ="Box3224"
                    LayoutCachedLeft =60
                    LayoutCachedTop =3494
                    LayoutCachedWidth =4605
                    LayoutCachedHeight =4154
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    Left =10140
                    Top =4080
                    Width =4680
                    Height =15
                    Name ="Box880"
                    LayoutCachedLeft =10140
                    LayoutCachedTop =4080
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =4095
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12525
                    Top =3510
                    Width =1335
                    FontSize =9
                    TabIndex =36
                    Name ="HBResidualLine"
                    ControlSource ="=Sum([HoldbackResidual])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12525
                    LayoutCachedTop =3510
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =3750
                    RowStart =5
                    RowEnd =5
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =10230
                            Top =3510
                            Width =2235
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2067"
                            Caption ="- Holdback Residual"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =10230
                            LayoutCachedTop =3510
                            LayoutCachedWidth =12465
                            LayoutCachedHeight =3750
                            RowStart =5
                            RowEnd =5
                            ColumnStart =11
                            ColumnEnd =11
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13935
                    Top =3510
                    Width =885
                    FontSize =9
                    TabIndex =37
                    ForeColor =8355711
                    Name ="Text2068"
                    ControlSource ="=IIf([HBResidualLine]>0,[HBResidualLine]/[ScheduleCost],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13935
                    LayoutCachedTop =3510
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =3750
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12525
                    Top =3810
                    Width =1335
                    FontSize =9
                    TabIndex =38
                    Name ="HBConversionLine"
                    ControlSource ="=Sum([HoldbackConversion])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12525
                    LayoutCachedTop =3810
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =4050
                    RowStart =5
                    RowEnd =5
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =10230
                            Top =3810
                            Width =2235
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2072"
                            Caption ="- Holdback Conversion"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =10230
                            LayoutCachedTop =3810
                            LayoutCachedWidth =12465
                            LayoutCachedHeight =4050
                            RowStart =5
                            RowEnd =5
                            ColumnStart =11
                            ColumnEnd =11
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13935
                    Top =3810
                    Width =885
                    FontSize =9
                    TabIndex =39
                    ForeColor =8355711
                    Name ="Text2073"
                    ControlSource ="=IIf([HBConversionLine]>0,[HBConversionLine]/[ScheduleCost],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13935
                    LayoutCachedTop =3810
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =4050
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12540
                    Top =3240
                    Width =1335
                    Height =285
                    FontSize =9
                    TabIndex =34
                    Name ="WitheldLine"
                    ControlSource ="=Sum([WithheldAmount])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12540
                    LayoutCachedTop =3240
                    LayoutCachedWidth =13875
                    LayoutCachedHeight =3525
                    RowStart =5
                    RowEnd =5
                    ColumnStart =12
                    ColumnEnd =12
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =10965
                            Top =3240
                            Width =1515
                            Height =240
                            FontSize =9
                            ForeColor =10040879
                            Name ="Label2081"
                            Caption ="- Advanced Rent"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =10965
                            LayoutCachedTop =3240
                            LayoutCachedWidth =12480
                            LayoutCachedHeight =3480
                            RowStart =5
                            RowEnd =5
                            ColumnStart =11
                            ColumnEnd =11
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13935
                    Top =3240
                    Width =885
                    FontSize =9
                    TabIndex =35
                    ForeColor =8355711
                    Name ="Text2082"
                    ControlSource ="=IIf([WitheldLine]>0,[WitheldLine]/[ScheduleCost],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13935
                    LayoutCachedTop =3240
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =3480
                    RowStart =5
                    RowEnd =5
                    ColumnStart =13
                    ColumnEnd =13
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12540
                    Top =4260
                    Width =1335
                    FontSize =9
                    FontWeight =700
                    TabIndex =46
                    Name ="FinalFeeLine"
                    ControlSource ="=[TotalFeeLine]-([HoldbackConversion]+[HoldbackResidual]+[WitheldLine])+Nz([Synd"
                        "MiscValue_1],0)+Nz([SyndMiscValue_2],0)"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =83

                    LayoutCachedLeft =12540
                    LayoutCachedTop =4260
                    LayoutCachedWidth =13875
                    LayoutCachedHeight =4500
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    ThemeFontIndex =-1
                    GroupTable =83
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =9000
                            Top =4260
                            Width =3480
                            Height =240
                            FontSize =9
                            FontWeight =700
                            ForeColor =10040879
                            Name ="Label3247"
                            Caption ="Group Total Fee"
                            FontName ="Segoe UI"
                            GroupTable =83
                            LayoutCachedLeft =9000
                            LayoutCachedTop =4260
                            LayoutCachedWidth =12480
                            LayoutCachedHeight =4500
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =4
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =83
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13932
                    Top =4260
                    Width =888
                    FontSize =9
                    FontWeight =700
                    TabIndex =47
                    ForeColor =8355711
                    Name ="Text3248"
                    ControlSource ="=[FinalFeeLine]/[ScheduleCost]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =83

                    LayoutCachedLeft =13932
                    LayoutCachedTop =4260
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =4500
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    ThemeFontIndex =-1
                    ForeTint =50.0
                    GroupTable =83
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12540
                    Top =4140
                    Width =1335
                    Height =0
                    FontSize =9
                    TabIndex =41
                    Name ="SyndMiscValue_1"
                    ControlSource ="=IIf(Nz([SyndMiscValue1],0)<>0,Sum([SyndMiscValue1]),Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =83

                    LayoutCachedLeft =12540
                    LayoutCachedTop =4140
                    LayoutCachedWidth =13875
                    LayoutCachedHeight =4140
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    ThemeFontIndex =-1
                    GroupTable =83
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13932
                    Top =4140
                    Width =888
                    Height =0
                    FontSize =9
                    TabIndex =42
                    ForeColor =8355711
                    Name ="SyndMiscValue_1Pct"
                    ControlSource ="=IIf(Nz([SyndMiscValue_1],0)<>0,[SyndMiscValue_1]/[ScheduleCost],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =83

                    LayoutCachedLeft =13932
                    LayoutCachedTop =4140
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =4140
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    ThemeFontIndex =-1
                    ForeTint =50.0
                    GroupTable =83
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9000
                    Top =4140
                    Width =3480
                    Height =0
                    FontSize =9
                    TabIndex =40
                    ForeColor =10040879
                    Name ="SyndMiscLabel_1"
                    ControlSource ="=IIf(Nz([SyndMiscValue1],0)<>0,[SyndMiscLabel1],Null)"
                    FontName ="Segoe UI"
                    GroupTable =83

                    LayoutCachedLeft =9000
                    LayoutCachedTop =4140
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =4140
                    LayoutGroup =4
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =83
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12540
                    Top =4200
                    Width =1335
                    Height =0
                    FontSize =9
                    TabIndex =44
                    Name ="SyndMiscValue_2"
                    ControlSource ="=IIf(Nz([SyndMiscValue2],0)<>0,Sum([SyndMiscValue2]),Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =83

                    LayoutCachedLeft =12540
                    LayoutCachedTop =4200
                    LayoutCachedWidth =13875
                    LayoutCachedHeight =4200
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    ThemeFontIndex =-1
                    GroupTable =83
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13932
                    Top =4200
                    Width =888
                    Height =0
                    FontSize =9
                    TabIndex =45
                    ForeColor =8355711
                    Name ="SyndMiscValue_2Pct"
                    ControlSource ="=IIf(Nz([SyndMiscValue_2],0)<>0,[SyndMiscValue_2]/[ScheduleCost],Null)"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =83

                    LayoutCachedLeft =13932
                    LayoutCachedTop =4200
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =4200
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    ThemeFontIndex =-1
                    ForeTint =50.0
                    GroupTable =83
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9000
                    Top =4200
                    Width =3480
                    Height =0
                    FontSize =9
                    TabIndex =43
                    ForeColor =10040879
                    Name ="SyndMiscLabel_2"
                    ControlSource ="=IIf(Nz([SyndMiscValue2],0)<>0,[SyndMiscLabel2],Null)"
                    FontName ="Segoe UI"
                    GroupTable =83

                    LayoutCachedLeft =9000
                    LayoutCachedTop =4200
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =4200
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =4
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =83
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4680
                    Top =2325
                    Width =7785
                    Height =330
                    FontSize =9
                    TabIndex =24
                    ForeColor =10040879
                    Name ="Label3218"
                    ControlSource ="[MsgInterim]"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =4680
                    LayoutCachedTop =2325
                    LayoutCachedWidth =12465
                    LayoutCachedHeight =2655
                    RowStart =5
                    RowEnd =5
                    ColumnStart =11
                    ColumnEnd =11
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4680
                    Top =2625
                    Width =7785
                    FontSize =9
                    TabIndex =29
                    ForeColor =10040879
                    Name ="MsgFSL"
                    ControlSource ="MsgFSL"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =4680
                    LayoutCachedTop =2625
                    LayoutCachedWidth =12465
                    LayoutCachedHeight =2865
                    RowStart =5
                    RowEnd =5
                    ColumnStart =11
                    ColumnEnd =11
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4680
                    Top =2025
                    Width =7785
                    FontSize =9
                    TabIndex =21
                    ForeColor =10040879
                    Name ="Text3249"
                    ControlSource ="=[MsgPerdiem]"
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedLeft =4680
                    LayoutCachedTop =2025
                    LayoutCachedWidth =12465
                    LayoutCachedHeight =2265
                    RowStart =5
                    RowEnd =5
                    ColumnStart =11
                    ColumnEnd =11
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "ScheduleEconomicsSyndReportSub.cls"
