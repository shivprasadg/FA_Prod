Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7568
    DatasheetFontHeight =11
    ItemSuffix =678
    Left =2250
    Top =2445
    Right =15030
    Bottom =10590
    RecSrcDt = Begin
        0xddb71fe671cee440
    End
    RecordSource ="BankRatesShare"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =300
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =5400
                    Width =1500
                    Height =300
                    FontSize =12
                    FontWeight =700
                    ForeColor =0
                    Name ="Label81"
                    Caption ="Bank Rates"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =5400
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =300
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3660
                    Width =660
                    FontSize =8
                    Name ="BankID"
                    ControlSource ="BankID"

                    LayoutCachedLeft =3660
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =240
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =50.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3120
                            Width =540
                            Height =240
                            FontSize =8
                            Name ="Label1"
                            Caption ="BankID"
                            LayoutCachedLeft =3120
                            LayoutCachedWidth =3660
                            LayoutCachedHeight =240
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =50.0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Width =2400
                    Height =300
                    FontSize =12
                    FontWeight =700
                    Name ="Label148"
                    Caption ="Buyout Rules (BTF)"
                    FontName ="Segoe UI"
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =300
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =75.0
                End
            End
        End
        Begin Section
            Height =5700
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1110
                    Top =375
                    Width =1080
                    Height =285
                    FontSize =10
                    Name ="ShareBank"
                    ControlSource ="ShareBank"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1110
                    LayoutCachedTop =375
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =240
                            Top =375
                            Width =864
                            Height =285
                            FontSize =9
                            ForeColor =16777215
                            Name ="Label91"
                            Caption ="Bank"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =240
                            LayoutCachedTop =375
                            LayoutCachedWidth =1104
                            LayoutCachedHeight =660
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =4
                            BackShade =75.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1110
                    Top =690
                    Width =1080
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="ShareFleet"
                    ControlSource ="ShareFleet"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1110
                    LayoutCachedTop =690
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =240
                            Top =690
                            Width =864
                            Height =285
                            FontSize =9
                            BackColor =4754549
                            ForeColor =16777215
                            Name ="Label88"
                            Caption ="FA"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =240
                            LayoutCachedTop =690
                            LayoutCachedWidth =1104
                            LayoutCachedHeight =975
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5955
                    Top =690
                    Width =1365
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="InterimRateLease"
                    ControlSource ="InterimRateLease"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =5955
                    LayoutCachedTop =690
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =4860
                            Top =690
                            Width =1095
                            Height =285
                            FontSize =10
                            FontWeight =700
                            ForeColor =0
                            Name ="Label104"
                            Caption ="Rate %"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =4860
                            LayoutCachedTop =690
                            LayoutCachedWidth =5955
                            LayoutCachedHeight =975
                            RowStart =2
                            RowEnd =2
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5955
                    Top =1665
                    Width =1365
                    Height =285
                    FontSize =10
                    TabIndex =9
                    Name ="InterimRateRemarketing"
                    ControlSource ="InterimRateRemarketing"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =5955
                    LayoutCachedTop =1665
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =1950
                    RowStart =5
                    RowEnd =5
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =4860
                            Top =1665
                            Width =1095
                            Height =285
                            FontSize =10
                            FontWeight =700
                            ForeColor =0
                            Name ="Label654"
                            Caption ="Rate %"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =4860
                            LayoutCachedTop =1665
                            LayoutCachedWidth =5955
                            LayoutCachedHeight =1950
                            RowStart =5
                            RowEnd =5
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =4020
                    Width =7440
                    Height =1590
                    FontSize =10
                    TabIndex =16
                    BackColor =13434879
                    Name ="Notes"
                    ControlSource ="Notes"
                    FontName ="Segoe UI"
                    GroupTable =2
                    TextFormat =1

                    LayoutCachedLeft =60
                    LayoutCachedTop =4020
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =5610
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =1110
                    Top =90
                    Width =1080
                    Height =285
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label7"
                    Caption ="Bank Share"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1110
                    LayoutCachedTop =90
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =375
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =5955
                    Top =375
                    Width =1365
                    Height =285
                    FontSize =9
                    ForeColor =16777215
                    Name ="Label8"
                    Caption =" Leasing"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5955
                    LayoutCachedTop =375
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =5955
                    Top =1335
                    Width =1365
                    Height =270
                    FontSize =9
                    ForeColor =16777215
                    Name ="Label9"
                    Caption ="RMKT"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =5955
                    LayoutCachedTop =1335
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =1605
                    RowStart =4
                    RowEnd =4
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =9
                    BackShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5955
                    Top =90
                    Width =1365
                    Height =285
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label83"
                    Caption ="Interim"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5955
                    LayoutCachedTop =90
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =375
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5955
                    Top =1035
                    Width =1365
                    FontSize =9
                    TabIndex =5
                    Name ="InterimCreditLimit"
                    ControlSource ="InterimCreditLimit"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =5955
                    LayoutCachedTop =1035
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =1275
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =4860
                            Top =1035
                            Width =1095
                            Height =240
                            FontSize =10
                            FontWeight =700
                            ForeColor =0
                            Name ="Label112"
                            Caption ="Credit Line"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =4860
                            LayoutCachedTop =1035
                            LayoutCachedWidth =5955
                            LayoutCachedHeight =1275
                            RowStart =3
                            RowEnd =3
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5955
                    Top =2010
                    Width =1365
                    Height =270
                    FontSize =9
                    TabIndex =11
                    Name ="InterimCreditLimit2"
                    ControlSource ="InterimCreditLimitRemarketing"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =5955
                    LayoutCachedTop =2010
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =2280
                    RowStart =6
                    RowEnd =6
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =4860
                            Top =2010
                            Width =1095
                            Height =270
                            FontSize =10
                            FontWeight =700
                            ForeColor =0
                            Name ="Label655"
                            Caption ="Credit Line"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =4860
                            LayoutCachedTop =2010
                            LayoutCachedWidth =5955
                            LayoutCachedHeight =2280
                            RowStart =6
                            RowEnd =6
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1110
                    Top =2340
                    Width =1080
                    Height =270
                    TabIndex =12
                    Name ="Text149"
                    ControlSource ="CalcType"
                    RowSourceType ="Value List"
                    RowSource ="\"IRR-PT\";\"MISF-PT\""
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1110
                    LayoutCachedTop =2340
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =2610
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =240
                            Top =2340
                            Width =864
                            Height =270
                            FontSize =9
                            Name ="Label150"
                            Caption ="Calc Type"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =240
                            LayoutCachedTop =2340
                            LayoutCachedWidth =1104
                            LayoutCachedHeight =2610
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =4
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =3
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3000
                    Top =375
                    Width =1245
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="ExchangeROFR"
                    ControlSource ="ExchangeROFR"
                    Format ="General Number"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3000
                    LayoutCachedTop =375
                    LayoutCachedWidth =4245
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =2415
                            Top =375
                            Width =585
                            Height =285
                            FontSize =9
                            ForeColor =16777215
                            Name ="Label152"
                            Caption ="Yes"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =2415
                            LayoutCachedTop =375
                            LayoutCachedWidth =3000
                            LayoutCachedHeight =660
                            RowStart =1
                            RowEnd =1
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =8
                            BackShade =75.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =3
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3000
                    Top =690
                    Width =1245
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="ExchangeNoROFR"
                    ControlSource ="ExchangeNoROFR"
                    Format ="General Number"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =3000
                    LayoutCachedTop =690
                    LayoutCachedWidth =4245
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =2415
                            Top =690
                            Width =585
                            Height =285
                            FontSize =9
                            ForeColor =16777215
                            Name ="Label154"
                            Caption ="No"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2415
                            LayoutCachedTop =690
                            LayoutCachedWidth =3000
                            LayoutCachedHeight =975
                            RowStart =2
                            RowEnd =2
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =8
                            BackShade =75.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1260
                    Top =3150
                    Width =6240
                    Height =270
                    FontSize =10
                    TabIndex =14
                    Name ="Text155"
                    ControlSource ="RMKTExclusive"
                    GroupTable =2

                    LayoutCachedLeft =1260
                    LayoutCachedTop =3150
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =3420
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =3150
                            Width =1140
                            Height =270
                            FontSize =10
                            FontWeight =700
                            Name ="Label156"
                            Caption ="Exclusive"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =60
                            LayoutCachedTop =3150
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =3420
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =4
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =4
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1110
                    Top =2010
                    Width =1080
                    Height =270
                    FontSize =10
                    TabIndex =10
                    Name ="ImpRateAdj"
                    ControlSource ="ImpRateAdj"
                    Format ="General Number"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1110
                    LayoutCachedTop =2010
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =2280
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =240
                            Top =2010
                            Width =864
                            Height =270
                            FontSize =9
                            Name ="Label158"
                            Caption ="Rate Adj"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =240
                            LayoutCachedTop =2010
                            LayoutCachedWidth =1104
                            LayoutCachedHeight =2280
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =4
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1260
                    Top =2820
                    Width =6240
                    Height =270
                    FontSize =10
                    TabIndex =13
                    Name ="InventoryRule"
                    ControlSource ="InventoryRule"
                    GroupTable =2

                    LayoutCachedLeft =1260
                    LayoutCachedTop =2820
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =3090
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =2820
                            Width =1140
                            Height =270
                            FontSize =10
                            FontWeight =700
                            Name ="Label160"
                            Caption ="Inventory"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =60
                            LayoutCachedTop =2820
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =3090
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =4
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1260
                    Top =3480
                    Width =6240
                    Height =270
                    FontSize =10
                    TabIndex =15
                    Name ="PVFormula"
                    ControlSource ="PVFormula"
                    GroupTable =2

                    LayoutCachedLeft =1260
                    LayoutCachedTop =3480
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =3750
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =3480
                            Width =1140
                            Height =270
                            FontSize =10
                            FontWeight =700
                            Name ="Label162"
                            Caption ="PV Formula"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =60
                            LayoutCachedTop =3480
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =3750
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =4
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3000
                    Top =1665
                    Width =1245
                    Height =285
                    FontSize =10
                    TabIndex =8
                    Name ="RepairTreshold"
                    ControlSource ="RepairTreshold"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =3000
                    LayoutCachedTop =1665
                    LayoutCachedWidth =4245
                    LayoutCachedHeight =1950
                    RowStart =5
                    RowEnd =5
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =2415
                            Top =1665
                            Width =585
                            Height =285
                            FontSize =9
                            Name ="Label192"
                            Caption ="Repair"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2415
                            LayoutCachedTop =1665
                            LayoutCachedWidth =3000
                            LayoutCachedHeight =1950
                            RowStart =5
                            RowEnd =5
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =4
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =720
                    Left =1110
                    Top =1665
                    Width =1080
                    Height =285
                    FontSize =10
                    TabIndex =7
                    ForeColor =4210752
                    Name ="ShortFall"
                    ControlSource ="ShortFall"
                    RowSourceType ="Value List"
                    RowSource ="1;\"Yes\";0;\"No\""
                    ColumnWidths ="0;576"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1110
                    LayoutCachedTop =1665
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =1950
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =75.0
                    ForeShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =240
                            Top =1665
                            Width =864
                            Height =285
                            FontSize =9
                            Name ="Combo279_Label"
                            Caption ="Short Fall"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =240
                            LayoutCachedTop =1665
                            LayoutCachedWidth =1104
                            LayoutCachedHeight =1950
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =4
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =720
                    Left =1110
                    Top =1335
                    Width =1080
                    Height =270
                    FontSize =10
                    TabIndex =6
                    ForeColor =4210752
                    Name ="OverMiles"
                    ControlSource ="OverMiles"
                    RowSourceType ="Value List"
                    RowSource ="1;\"Yes\";0;\"No\""
                    ColumnWidths ="0;576"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1110
                    LayoutCachedTop =1335
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =1605
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =75.0
                    ForeShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =240
                            Top =1335
                            Width =864
                            Height =270
                            FontSize =9
                            Name ="Label294"
                            Caption ="Over Miles"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =240
                            LayoutCachedTop =1335
                            LayoutCachedWidth =1104
                            LayoutCachedHeight =1605
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =4
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin EmptyCell
                    Left =240
                    Top =90
                    Width =864
                    Height =285
                    Name ="EmptyCell330"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =240
                    LayoutCachedTop =90
                    LayoutCachedWidth =1104
                    LayoutCachedHeight =375
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =2415
                    Top =90
                    Width =1830
                    Height =285
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label354"
                    Caption ="Early Term Fee ROFR"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2415
                    LayoutCachedTop =90
                    LayoutCachedWidth =4245
                    LayoutCachedHeight =375
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2190
                    Top =90
                    Width =225
                    Height =285
                    Name ="EmptyCell501"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2190
                    LayoutCachedTop =90
                    LayoutCachedWidth =2415
                    LayoutCachedHeight =375
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2190
                    Top =375
                    Width =225
                    Height =285
                    Name ="EmptyCell502"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2190
                    LayoutCachedTop =375
                    LayoutCachedWidth =2415
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2190
                    Top =690
                    Width =225
                    Height =285
                    Name ="EmptyCell503"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2190
                    LayoutCachedTop =690
                    LayoutCachedWidth =2415
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2190
                    Top =1335
                    Width =225
                    Height =270
                    Name ="EmptyCell504"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2190
                    LayoutCachedTop =1335
                    LayoutCachedWidth =2415
                    LayoutCachedHeight =1605
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2190
                    Top =1665
                    Width =225
                    Height =285
                    Name ="EmptyCell505"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2190
                    LayoutCachedTop =1665
                    LayoutCachedWidth =2415
                    LayoutCachedHeight =1950
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2190
                    Top =2010
                    Width =225
                    Height =270
                    Name ="EmptyCell506"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2190
                    LayoutCachedTop =2010
                    LayoutCachedWidth =2415
                    LayoutCachedHeight =2280
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2190
                    Top =2340
                    Width =225
                    Height =270
                    Name ="EmptyCell544"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2190
                    LayoutCachedTop =2340
                    LayoutCachedWidth =2415
                    LayoutCachedHeight =2610
                    RowStart =7
                    RowEnd =7
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2415
                    Top =2340
                    Width =585
                    Height =270
                    Name ="EmptyCell545"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2415
                    LayoutCachedTop =2340
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =2610
                    RowStart =7
                    RowEnd =7
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3000
                    Top =2340
                    Width =1245
                    Height =270
                    Name ="EmptyCell546"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =3000
                    LayoutCachedTop =2340
                    LayoutCachedWidth =4245
                    LayoutCachedHeight =2610
                    RowStart =7
                    RowEnd =7
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2415
                    Top =2010
                    Width =585
                    Height =270
                    Name ="EmptyCell547"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2415
                    LayoutCachedTop =2010
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =2280
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3000
                    Top =2010
                    Width =1245
                    Height =270
                    Name ="EmptyCell548"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =3000
                    LayoutCachedTop =2010
                    LayoutCachedWidth =4245
                    LayoutCachedHeight =2280
                    RowStart =6
                    RowEnd =6
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3000
                    Top =1335
                    Width =1245
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label563"
                    Caption ="Treshold"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =3000
                    LayoutCachedTop =1335
                    LayoutCachedWidth =4245
                    LayoutCachedHeight =1605
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =8
                    BackShade =75.0
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =240
                    Top =1035
                    Width =864
                    Name ="EmptyCell568"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =240
                    LayoutCachedTop =1035
                    LayoutCachedWidth =1104
                    LayoutCachedHeight =1275
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =1110
                    Top =1035
                    Width =1080
                    Name ="EmptyCell569"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =1110
                    LayoutCachedTop =1035
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =1275
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2190
                    Top =1035
                    Width =225
                    Name ="EmptyCell571"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2190
                    LayoutCachedTop =1035
                    LayoutCachedWidth =2415
                    LayoutCachedHeight =1275
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2415
                    Top =1035
                    Width =585
                    Name ="EmptyCell572"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2415
                    LayoutCachedTop =1035
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =1275
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4245
                    Top =90
                    Width =615
                    Height =285
                    Name ="EmptyCell577"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4245
                    LayoutCachedTop =90
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =375
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4245
                    Top =375
                    Width =615
                    Height =285
                    Name ="EmptyCell578"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4245
                    LayoutCachedTop =375
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4245
                    Top =690
                    Width =615
                    Height =285
                    Name ="EmptyCell579"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4245
                    LayoutCachedTop =690
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4245
                    Top =1035
                    Width =615
                    Name ="EmptyCell580"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4245
                    LayoutCachedTop =1035
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =1275
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4245
                    Top =1335
                    Width =615
                    Height =270
                    Name ="EmptyCell581"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4245
                    LayoutCachedTop =1335
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =1605
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4245
                    Top =1665
                    Width =615
                    Height =285
                    Name ="EmptyCell582"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4245
                    LayoutCachedTop =1665
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =1950
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4245
                    Top =2010
                    Width =615
                    Height =270
                    Name ="EmptyCell583"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4245
                    LayoutCachedTop =2010
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =2280
                    RowStart =6
                    RowEnd =6
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4245
                    Top =2340
                    Width =615
                    Height =270
                    Name ="EmptyCell584"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4245
                    LayoutCachedTop =2340
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =2610
                    RowStart =7
                    RowEnd =7
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4860
                    Top =90
                    Width =1095
                    Height =285
                    Name ="EmptyCell585"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4860
                    LayoutCachedTop =90
                    LayoutCachedWidth =5955
                    LayoutCachedHeight =375
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4860
                    Top =2340
                    Width =1095
                    Height =270
                    Name ="EmptyCell592"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4860
                    LayoutCachedTop =2340
                    LayoutCachedWidth =5955
                    LayoutCachedHeight =2610
                    RowStart =7
                    RowEnd =7
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5955
                    Top =2340
                    Width =1365
                    Height =270
                    Name ="EmptyCell600"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =5955
                    LayoutCachedTop =2340
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =2610
                    RowStart =7
                    RowEnd =7
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4860
                    Top =1335
                    Width =1095
                    Height =270
                    Name ="EmptyCell663"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4860
                    LayoutCachedTop =1335
                    LayoutCachedWidth =5955
                    LayoutCachedHeight =1605
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4860
                    Top =375
                    Width =1095
                    Height =285
                    Name ="EmptyCell668"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4860
                    LayoutCachedTop =375
                    LayoutCachedWidth =5955
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2415
                    Top =1335
                    Width =585
                    Height =270
                    Name ="EmptyCell674"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2415
                    LayoutCachedTop =1335
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =1605
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3000
                    Top =1035
                    Width =1245
                    Name ="EmptyCell675"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =3000
                    LayoutCachedTop =1035
                    LayoutCachedWidth =4245
                    LayoutCachedHeight =1275
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =60
                    Top =3810
                    Width =1140
                    Height =150
                    Name ="EmptyCell676"
                    GroupTable =2
                    LayoutCachedLeft =60
                    LayoutCachedTop =3810
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =3960
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =1260
                    Top =3810
                    Width =6240
                    Height =150
                    Name ="EmptyCell677"
                    GroupTable =2
                    LayoutCachedLeft =1260
                    LayoutCachedTop =3810
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =3960
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
