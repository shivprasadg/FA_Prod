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
    Width =15293
    DatasheetFontHeight =11
    ItemSuffix =123
    Left =5820
    Top =2700
    RecSrcDt = Begin
        0xe6367e3facdce440
    End
    RecordSource ="DrawDown_TransactionSummary"
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
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
        Begin Subform
            BorderLineStyle =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            ControlSource ="UnitGroup"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="UnitCost"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1260
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1140
                    Width =5730
                    Height =315
                    ColumnWidth =3765
                    ColumnOrder =0
                    Name ="BankName"
                    ControlSource ="BankName"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1140
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =315
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            TextAlign =3
                            Width =1140
                            Height =315
                            Name ="Label1"
                            Caption ="Bank "
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1140
                    Top =315
                    Width =5730
                    Height =315
                    ColumnWidth =3330
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Guaranty"
                    ControlSource ="Guaranty"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1140
                    LayoutCachedTop =315
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =630
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            TextAlign =3
                            Top =315
                            Width =1140
                            Height =315
                            Name ="Label2"
                            Caption ="Guaranty "
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedTop =315
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =630
                            RowStart =1
                            RowEnd =1
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1140
                    Top =630
                    Width =5730
                    Height =315
                    ColumnWidth =3765
                    ColumnOrder =2
                    TabIndex =2
                    Name ="CompanyName"
                    ControlSource ="CompanyName"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1140
                    LayoutCachedTop =630
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =945
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            TextAlign =3
                            Top =630
                            Width =1140
                            Height =315
                            Name ="Label3"
                            Caption ="Lessee "
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedTop =630
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =945
                            RowStart =2
                            RowEnd =2
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1140
                    Top =945
                    Width =5730
                    Height =315
                    ColumnOrder =3
                    TabIndex =3
                    Name ="Text11"
                    ControlSource ="=Date()"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1140
                    LayoutCachedTop =945
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =1260
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            TextAlign =3
                            Top =945
                            Width =1140
                            Height =315
                            Name ="Label12"
                            Caption ="Date "
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedTop =945
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =1260
                            RowStart =3
                            RowEnd =3
                        End
                    End
                End
            End
        End
        Begin PageHeader
            Height =600
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    Top =285
                    Width =1620
                    Height =300
                    FontSize =9
                    ForeColor =0
                    Name ="Label18"
                    Caption ="Schedule"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =285
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =585
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    TextAlign =2
                    Left =1620
                    Top =285
                    Width =750
                    Height =300
                    FontSize =9
                    ForeColor =0
                    Name ="Label21"
                    Caption ="Group"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1620
                    LayoutCachedTop =285
                    LayoutCachedWidth =2370
                    LayoutCachedHeight =585
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    TextAlign =2
                    Left =9210
                    Top =285
                    Width =714
                    Height =300
                    FontSize =9
                    ForeColor =0
                    Name ="Label28"
                    Caption ="QTY"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9210
                    LayoutCachedTop =285
                    LayoutCachedWidth =9924
                    LayoutCachedHeight =585
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    TextAlign =1
                    Left =2370
                    Top =285
                    Width =5472
                    Height =300
                    FontSize =9
                    ForeColor =0
                    Name ="Label35"
                    Caption =" Unit Desc"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2370
                    LayoutCachedTop =285
                    LayoutCachedWidth =7842
                    LayoutCachedHeight =585
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    TextAlign =2
                    Left =7845
                    Top =285
                    Width =1365
                    Height =300
                    FontSize =9
                    ForeColor =0
                    Name ="Label42"
                    Caption ="Unit Cost"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7845
                    LayoutCachedTop =285
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =585
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    TextAlign =2
                    Left =9930
                    Top =285
                    Width =1365
                    Height =300
                    FontSize =9
                    ForeColor =0
                    Name ="Label49"
                    Caption ="Total Cost"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9930
                    LayoutCachedTop =285
                    LayoutCachedWidth =11295
                    LayoutCachedHeight =585
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    TextAlign =2
                    Left =11295
                    Top =285
                    Width =900
                    Height =300
                    FontSize =9
                    ForeColor =0
                    Name ="Label56"
                    Caption ="Year"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11295
                    LayoutCachedTop =285
                    LayoutCachedWidth =12195
                    LayoutCachedHeight =585
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    TextAlign =2
                    Left =12195
                    Top =285
                    Width =705
                    Height =300
                    FontSize =9
                    ForeColor =0
                    Name ="Label63"
                    Caption ="Mths"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12195
                    LayoutCachedTop =285
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =585
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    TextAlign =2
                    Left =12900
                    Top =285
                    Width =1020
                    Height =300
                    FontSize =9
                    ForeColor =0
                    Name ="Label70"
                    Caption ="Delivery"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12900
                    LayoutCachedTop =285
                    LayoutCachedWidth =13920
                    LayoutCachedHeight =585
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    TextAlign =2
                    Left =13920
                    Top =285
                    Width =1365
                    Height =300
                    FontSize =9
                    ForeColor =0
                    Name ="Label77"
                    Caption ="Per Unit"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13920
                    LayoutCachedTop =285
                    LayoutCachedWidth =15285
                    LayoutCachedHeight =585
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    TextAlign =2
                    Left =11295
                    Width =900
                    Height =285
                    FontSize =9
                    ForeColor =0
                    Name ="Label84"
                    Caption ="Allowance"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11295
                    LayoutCachedWidth =12195
                    LayoutCachedHeight =285
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin EmptyCell
                    Width =1620
                    Height =285
                    Name ="EmptyCell86"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =285
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =1620
                    Width =750
                    Height =285
                    Name ="EmptyCell87"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1620
                    LayoutCachedWidth =2370
                    LayoutCachedHeight =285
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =2370
                    Width =5472
                    Height =285
                    Name ="EmptyCell89"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2370
                    LayoutCachedWidth =7842
                    LayoutCachedHeight =285
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =7845
                    Width =1365
                    Height =285
                    Name ="EmptyCell90"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7845
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =285
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =9930
                    Width =1365
                    Height =285
                    Name ="EmptyCell91"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9930
                    LayoutCachedWidth =11295
                    LayoutCachedHeight =285
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin Label
                    TextAlign =2
                    Left =12900
                    Width =1020
                    Height =285
                    FontSize =9
                    ForeColor =0
                    Name ="Label96"
                    Caption ="Est."
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12900
                    LayoutCachedWidth =13920
                    LayoutCachedHeight =285
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    TextAlign =2
                    Left =12195
                    Width =705
                    Height =285
                    FontSize =9
                    ForeColor =0
                    Name ="Label98"
                    Caption ="Term"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12195
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =285
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =9210
                    Width =714
                    Height =285
                    Name ="EmptyCell102"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9210
                    LayoutCachedWidth =9924
                    LayoutCachedHeight =285
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =4
                End
                Begin Label
                    TextAlign =2
                    Left =13920
                    Width =1365
                    Height =285
                    FontSize =9
                    ForeColor =0
                    Name ="Label110"
                    Caption ="PMT"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13920
                    LayoutCachedWidth =15285
                    LayoutCachedHeight =285
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Line
                    BorderWidth =1
                    Top =585
                    Width =15264
                    Name ="Line120"
                    LayoutCachedTop =585
                    LayoutCachedWidth =15264
                    LayoutCachedHeight =585
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =300
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Width =1620
                    Height =300
                    FontSize =9
                    BorderColor =14277081
                    Name ="Schedule"
                    ControlSource ="Schedule"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =1620
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    GroupTable =4
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1620
                    Width =750
                    Height =300
                    FontSize =9
                    TabIndex =1
                    BorderColor =14277081
                    Name ="UnitGroup"
                    ControlSource ="UnitGroup"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1620
                    LayoutCachedWidth =2370
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    GroupTable =4
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9210
                    Width =714
                    Height =300
                    FontSize =9
                    TabIndex =4
                    BorderColor =14277081
                    Name ="QTY"
                    ControlSource ="QTY"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9210
                    LayoutCachedWidth =9924
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    GroupTable =4
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2370
                    Width =5472
                    Height =300
                    ColumnWidth =5970
                    FontSize =9
                    TabIndex =2
                    BorderColor =14277081
                    Name ="UnitDesc"
                    ControlSource ="UnitDesc"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2370
                    LayoutCachedWidth =7842
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    GroupTable =4
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7845
                    Width =1365
                    Height =300
                    FontSize =9
                    TabIndex =3
                    BorderColor =14277081
                    Name ="UnitCost"
                    ControlSource ="UnitCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7845
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    GroupTable =4
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9930
                    Width =1365
                    Height =300
                    FontSize =9
                    TabIndex =5
                    BorderColor =14277081
                    Name ="TotalCost"
                    ControlSource ="TotalCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9930
                    LayoutCachedWidth =11295
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    GroupTable =4
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11295
                    Width =900
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =14277081
                    Name ="MPYAllowed"
                    ControlSource ="MPYAllowed"
                    Format ="Standard"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11295
                    LayoutCachedWidth =12195
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    GroupTable =4
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12195
                    Width =705
                    Height =300
                    FontSize =9
                    TabIndex =7
                    BorderColor =14277081
                    Name ="LeaseTermPrimary"
                    ControlSource ="LeaseTermPrimary"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12195
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    GroupTable =4
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12900
                    Width =1020
                    Height =300
                    FontSize =9
                    TabIndex =8
                    BorderColor =14277081
                    Name ="EstDelivery"
                    ControlSource ="EstDelivery"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =12900
                    LayoutCachedWidth =13920
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    GroupTable =4
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13920
                    Width =1365
                    Height =300
                    FontSize =9
                    TabIndex =9
                    BorderColor =14277081
                    Name ="PMT"
                    ControlSource ="=IIf([UNITSRENT]=0,Null,[UNITSRENT])"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="MONTLHY STRIP RENT"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13920
                    LayoutCachedWidth =15285
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    GroupTable =4
                    CurrencySymbol ="$"
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
        Begin FormFooter
            KeepTogether = NotDefault
            Height =420
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9915
                    Top =60
                    Width =1365
                    Height =300
                    FontSize =9
                    Name ="Text105"
                    ControlSource ="=Sum([TotalCost])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =9915
                    LayoutCachedTop =60
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13905
                    Top =60
                    Width =1368
                    Height =300
                    FontSize =9
                    TabIndex =1
                    Name ="Text106"
                    ControlSource ="=Sum([UNITSRENT])"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="MONTLHY STRIP RENT"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =13905
                    LayoutCachedTop =60
                    LayoutCachedWidth =15273
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    CurrencySymbol ="$"
                End
                Begin Line
                    BorderWidth =1
                    Width =15264
                    Name ="Line121"
                    LayoutCachedWidth =15264
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9210
                    Top =60
                    Width =714
                    Height =300
                    FontSize =9
                    TabIndex =2
                    Name ="Text122"
                    ControlSource ="=Sum([QTY])"
                    Format ="Standard"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =9210
                    LayoutCachedTop =60
                    LayoutCachedWidth =9924
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                End
            End
        End
    End
End
