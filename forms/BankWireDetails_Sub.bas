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
    Width =7643
    DatasheetFontHeight =11
    ItemSuffix =168
    Left =2850
    Top =10440
    Right =11715
    Bottom =14265
    RecSrcDt = Begin
        0x19762c1e3013e640
    End
    RecordSource ="BankPAI"
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
                    Width =2400
                    Height =300
                    FontSize =12
                    FontWeight =700
                    Name ="Label81"
                    Caption ="Bank Wire Details"
                    FontName ="Segoe UI"
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =300
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =2820
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1080
                    Top =2160
                    Width =660
                    FontSize =9
                    TabIndex =5
                    Name ="BankID"
                    ControlSource ="BankID"

                    LayoutCachedLeft =1080
                    LayoutCachedTop =2160
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =2400
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =50.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =420
                            Top =2160
                            Width =660
                            Height =240
                            FontSize =9
                            Name ="Label1"
                            Caption ="BankID"
                            LayoutCachedLeft =420
                            LayoutCachedTop =2160
                            LayoutCachedWidth =1080
                            LayoutCachedHeight =2400
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =50.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1875
                    Top =60
                    Width =5580
                    Height =315
                    FontSize =12
                    FontWeight =700
                    Name ="ABANumber"
                    ControlSource ="ABANumber"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1875
                    LayoutCachedTop =60
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =375
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =60
                            Width =1695
                            Height =315
                            FontSize =10
                            FontWeight =700
                            Name ="Label104"
                            Caption ="ABA Number"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =60
                            LayoutCachedWidth =1875
                            LayoutCachedHeight =375
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1875
                    Top =435
                    Width =5580
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="InterimCreditLimit"
                    ControlSource ="AccountName"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1875
                    LayoutCachedTop =435
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =435
                            Width =1695
                            Height =315
                            FontSize =10
                            FontWeight =700
                            Name ="Label112"
                            Caption ="Account Name"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =435
                            LayoutCachedWidth =1875
                            LayoutCachedHeight =750
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1875
                    Top =810
                    Width =5580
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="AccountNumber"
                    ControlSource ="AccountNumber"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1875
                    LayoutCachedTop =810
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =1125
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =810
                            Width =1695
                            Height =315
                            FontSize =10
                            FontWeight =700
                            Name ="Label136"
                            Caption ="Account Number"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =810
                            LayoutCachedWidth =1875
                            LayoutCachedHeight =1125
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1875
                    Top =1185
                    Width =5580
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    Name ="BankName"
                    ControlSource ="BankName"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1875
                    LayoutCachedTop =1185
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =1500
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =1185
                            Width =1695
                            Height =315
                            FontSize =10
                            FontWeight =700
                            Name ="Label138"
                            Caption ="Bank Name"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =1185
                            LayoutCachedWidth =1875
                            LayoutCachedHeight =1500
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1875
                    Top =1560
                    Width =5580
                    Height =1125
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    Name ="BankAddress"
                    ControlSource ="BankAddress"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1875
                    LayoutCachedTop =1560
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =2685
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =180
                            Top =1560
                            Width =1695
                            Height =1125
                            FontSize =10
                            FontWeight =700
                            Name ="Label140"
                            Caption ="Bank Address"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =1560
                            LayoutCachedWidth =1875
                            LayoutCachedHeight =2685
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
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
