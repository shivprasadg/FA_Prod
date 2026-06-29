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
    ItemSuffix =162
    Top =4740
    RecSrcDt = Begin
        0xa6bce09fd524e640
    End
    RecordSource ="SELECT * FROM PayProceeds_data_Letter WHERE (((PayProceeds_data_Letter.PayVendor"
        ")>\"0\"));"
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
            ControlSource ="Vendor"
        End
        Begin BreakLevel
            ControlSource ="PayVendor"
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
            Height =3180
            Name ="GroupHeader0"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    Left =420
                    Top =3012
                    Width =10824
                    Height =168
                    Name ="Box1730"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =420
                    LayoutCachedTop =3012
                    LayoutCachedWidth =11244
                    LayoutCachedHeight =3180
                    RowStart =7
                    RowEnd =7
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =2
                    BorderTint =60.0
                    BorderShade =100.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4764
                    Top =2652
                    Width =4332
                    Height =348
                    TabIndex =7
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text1600"
                    ControlSource ="VendorPayee"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4764
                    LayoutCachedTop =2652
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =3000
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
                            TextAlign =1
                            Left =420
                            Top =2652
                            Width =4335
                            Height =348
                            LeftMargin =29
                            RightMargin =29
                            ForeColor =0
                            Name ="Label1601"
                            Caption ="Account Name: "
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =420
                            LayoutCachedTop =2652
                            LayoutCachedWidth =4755
                            LayoutCachedHeight =3000
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
                    CanGrow = NotDefault
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4764
                    Top =2292
                    Width =4332
                    Height =348
                    TabIndex =6
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text1592"
                    ControlSource ="VendorABA"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4764
                    LayoutCachedTop =2292
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =2640
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
                            TextAlign =1
                            Left =420
                            Top =2292
                            Width =4335
                            Height =348
                            LeftMargin =29
                            RightMargin =29
                            ForeColor =0
                            Name ="Label1593"
                            Caption ="ABA Routing #: "
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =420
                            LayoutCachedTop =2292
                            LayoutCachedWidth =4755
                            LayoutCachedHeight =2640
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
                    CanGrow = NotDefault
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4764
                    Top =1932
                    Width =4332
                    Height =348
                    TabIndex =5
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Label1585"
                    ControlSource ="VendorBankAcct"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4764
                    LayoutCachedTop =1932
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =2280
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
                            TextAlign =1
                            Left =420
                            Top =1932
                            Width =4335
                            Height =348
                            LeftMargin =29
                            RightMargin =29
                            ForeColor =0
                            Name ="Label1584"
                            Caption ="Account #: "
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =420
                            LayoutCachedTop =1932
                            LayoutCachedWidth =4755
                            LayoutCachedHeight =2280
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
                    CanGrow = NotDefault
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4764
                    Top =1572
                    Width =4332
                    Height =348
                    TabIndex =4
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Label1583"
                    ControlSource ="VendorBank"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4764
                    LayoutCachedTop =1572
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =1920
                    RowStart =3
                    RowEnd =3
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
                            TextAlign =1
                            Left =420
                            Top =1572
                            Width =4335
                            Height =348
                            LeftMargin =29
                            RightMargin =29
                            ForeColor =0
                            Name ="Label1582"
                            Caption ="Bank Name: "
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =420
                            LayoutCachedTop =1572
                            LayoutCachedWidth =4755
                            LayoutCachedHeight =1920
                            RowStart =3
                            RowEnd =3
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
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =420
                    Top =1320
                    Width =8676
                    FontSize =9
                    TabIndex =3
                    LeftMargin =29
                    RightMargin =29
                    Name ="Label1684"
                    ControlSource ="=\" <b> \" & [VendorName] & \"</b> Wire Instructions\""
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =420
                    LayoutCachedTop =1320
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =1560
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =2
                    BorderTint =60.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =420
                    Top =300
                    Width =8676
                    Height =645
                    FontSize =12
                    LeftMargin =29
                    RightMargin =29
                    Name ="Text11"
                    ControlSource ="=\" <b>Amount to be wired to: \" & [VendorName] & \"</b>  (Vendor):\""
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =420
                    LayoutCachedTop =300
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =945
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =2
                    BorderTint =60.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9108
                    Top =300
                    Width =2136
                    Height =2700
                    FontWeight =700
                    TabIndex =1
                    LeftMargin =29
                    RightMargin =29
                    ForeColor =0
                    Name ="Text57"
                    ControlSource ="=\"<br/><br/>\" & Format(Sum([PayVendor]),\"$#,###.00\") & \"  \""
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =9108
                    LayoutCachedTop =300
                    LayoutCachedWidth =11244
                    LayoutCachedHeight =3000
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =420
                    Top =960
                    Width =8676
                    Height =348
                    TabIndex =2
                    LeftMargin =29
                    RightMargin =29
                    ForeColor =0
                    Name ="Text71"
                    ControlSource ="=\"<b>Equipment:</b> - (\" & Sum(Nz([vUnits],0)) & \") \" & [ItemDesc] & \"\""
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =420
                    LayoutCachedTop =960
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =1308
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =1
                    Width =2100
                    Height =300
                    FontWeight =700
                    LeftMargin =29
                    RightMargin =29
                    Name ="Label33"
                    Caption ="  Vendor"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =300
                    RowEnd =6
                    ThemeFontIndex =-1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =2
                    BorderTint =60.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =75.0
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
