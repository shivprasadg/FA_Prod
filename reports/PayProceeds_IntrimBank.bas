Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11250
    DatasheetFontHeight =11
    ItemSuffix =40
    Left =3765
    Top =5625
    RecSrcDt = Begin
        0xf72edf9fd524e640
    End
    RecordSource ="SELECT * FROM PayProceeds_data_Letter WHERE (((PayProceeds_data_Letter.InterimAm"
        "t)>0));"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
            ForeTint =60.0
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
            ControlSource ="InterimBank"
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
            Height =3912
            Name ="GroupHeader0"
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =420
                    Top =1620
                    Width =8676
                    Height =360
                    TabIndex =3
                    LeftMargin =29
                    RightMargin =29
                    ForeColor =0
                    Name ="Text2127"
                    ControlSource ="=\"<b>Equipment:</b> \" & [Schs]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =420
                    LayoutCachedTop =1620
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =1980
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =420
                    Top =300
                    Width =3975
                    Height =930
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =8355711
                    ForeColor =0
                    Name ="Text2141"
                    ControlSource ="=[IntBankAddr]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =420
                    LayoutCachedTop =300
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =1230
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =420
                    Top =1992
                    Width =8676
                    Height =252
                    FontSize =9
                    TabIndex =4
                    LeftMargin =29
                    RightMargin =29
                    BackColor =14277081
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text2147"
                    ControlSource ="=\" <b> \" & [InterimBank] & \"</b> Wire Instructions\""
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =420
                    LayoutCachedTop =1992
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =2244
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =1
                    LayoutGroup =1
                    BackShade =85.0
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4404
                    Top =2256
                    Width =4692
                    Height =396
                    TabIndex =5
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text2165"
                    ControlSource ="iWireBank"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4404
                    LayoutCachedTop =2256
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =2652
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =3
                            Left =420
                            Top =2256
                            Width =3975
                            Height =396
                            LeftMargin =29
                            RightMargin =29
                            ForeColor =0
                            Name ="Label2166"
                            Caption ="Bank Name: "
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =420
                            LayoutCachedTop =2256
                            LayoutCachedWidth =4395
                            LayoutCachedHeight =2652
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            BorderShade =75.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4404
                    Top =2664
                    Width =4692
                    Height =348
                    TabIndex =6
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text2167"
                    ControlSource ="iWireAcctNum"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4404
                    LayoutCachedTop =2664
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =3012
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =3
                            Left =420
                            Top =2664
                            Width =3975
                            Height =348
                            LeftMargin =29
                            RightMargin =29
                            ForeColor =0
                            Name ="Label2168"
                            Caption ="Account #: "
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =420
                            LayoutCachedTop =2664
                            LayoutCachedWidth =4395
                            LayoutCachedHeight =3012
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            BorderShade =75.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4404
                    Top =3024
                    Width =4692
                    Height =348
                    TabIndex =7
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text2169"
                    ControlSource ="iWireABA"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4404
                    LayoutCachedTop =3024
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =3372
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =3
                            Left =420
                            Top =3024
                            Width =3975
                            Height =348
                            LeftMargin =29
                            RightMargin =29
                            ForeColor =0
                            Name ="Label2170"
                            Caption ="ABA Routing #: "
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =420
                            LayoutCachedTop =3024
                            LayoutCachedWidth =4395
                            LayoutCachedHeight =3372
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            BorderShade =75.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4404
                    Top =3384
                    Width =4692
                    Height =348
                    TabIndex =8
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text2171"
                    ControlSource ="iWireAcctName"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4404
                    LayoutCachedTop =3384
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =3732
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            TextAlign =3
                            Left =420
                            Top =3384
                            Width =3975
                            Height =348
                            LeftMargin =29
                            RightMargin =29
                            ForeColor =0
                            Name ="Label2172"
                            Caption ="Account Name: "
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =420
                            LayoutCachedTop =3384
                            LayoutCachedWidth =4395
                            LayoutCachedHeight =3732
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            BorderShade =75.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin Rectangle
                    Left =420
                    Top =3744
                    Width =10824
                    Height =168
                    BackColor =14277081
                    BorderColor =14277081
                    Name ="Box2197"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =420
                    LayoutCachedTop =3744
                    LayoutCachedWidth =11244
                    LayoutCachedHeight =3912
                    RowStart =8
                    RowEnd =8
                    ColumnEnd =2
                    LayoutGroup =1
                    BackShade =85.0
                    BorderShade =85.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =9108
                    Top =300
                    Width =2136
                    Height =930
                    Name ="EmptyCell2214"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9108
                    LayoutCachedTop =300
                    LayoutCachedWidth =11244
                    LayoutCachedHeight =1230
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =420
                    Top =1248
                    Width =8676
                    Height =360
                    TabIndex =1
                    LeftMargin =29
                    RightMargin =29
                    BackColor =14277081
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text11"
                    ControlSource ="=\"Wired Loan Pay-Off to<b> \" & [InterimBank] & \" (Lender):\""
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =420
                    LayoutCachedTop =1248
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =1608
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackShade =85.0
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4404
                    Top =300
                    Width =4692
                    Height =930
                    Name ="EmptyCell36"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4404
                    LayoutCachedTop =300
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =1230
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9108
                    Top =1248
                    Width =2136
                    Height =2484
                    FontWeight =700
                    TabIndex =2
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =0
                    ForeColor =0
                    Name ="Text1313"
                    ControlSource ="=\"<br/>\" & Format(Sum([InterimAmt]),\"$#,###.00\") & \"  \""
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =9108
                    LayoutCachedTop =1248
                    LayoutCachedWidth =11244
                    LayoutCachedHeight =3732
                    RowStart =1
                    RowEnd =7
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Width =2700
                    Height =299
                    FontWeight =700
                    TabIndex =9
                    LeftMargin =29
                    RightMargin =29
                    BackColor =14277081
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Label33"
                    ControlSource ="=\"   Interim Bank \""
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =2700
                    LayoutCachedHeight =299
                    RowEnd =8
                    BackShade =85.0
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =0
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' See "PayProceeds_IntrimBank.cls"
