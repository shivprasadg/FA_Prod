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
    Width =11228
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
                    Top =2910
                    Width =10800
                    Height =165
                    Name ="Box1730"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =420
                    LayoutCachedTop =2910
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =3075
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
                    Left =4755
                    Top =2565
                    Width =4335
                    Height =345
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

                    LayoutCachedLeft =4755
                    LayoutCachedTop =2565
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =2910
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
                            Top =2565
                            Width =4335
                            Height =345
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
                            LayoutCachedTop =2565
                            LayoutCachedWidth =4755
                            LayoutCachedHeight =2910
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
                    Left =4755
                    Top =2220
                    Width =4335
                    Height =345
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

                    LayoutCachedLeft =4755
                    LayoutCachedTop =2220
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =2565
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
                            Top =2220
                            Width =4335
                            Height =345
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
                            LayoutCachedTop =2220
                            LayoutCachedWidth =4755
                            LayoutCachedHeight =2565
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
                    Left =4755
                    Top =1875
                    Width =4335
                    Height =345
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

                    LayoutCachedLeft =4755
                    LayoutCachedTop =1875
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =2220
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
                            Top =1875
                            Width =4335
                            Height =345
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
                            LayoutCachedTop =1875
                            LayoutCachedWidth =4755
                            LayoutCachedHeight =2220
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
                    Left =4755
                    Top =1530
                    Width =4335
                    Height =345
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

                    LayoutCachedLeft =4755
                    LayoutCachedTop =1530
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =1875
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
                            Top =1530
                            Width =4335
                            Height =345
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
                            LayoutCachedTop =1530
                            LayoutCachedWidth =4755
                            LayoutCachedHeight =1875
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
                    Top =1290
                    Width =8670
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
                    LayoutCachedTop =1290
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =1530
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
                    Width =8670
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
                    LayoutCachedWidth =9090
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
                    Left =9090
                    Top =300
                    Width =2130
                    Height =2610
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

                    LayoutCachedLeft =9090
                    LayoutCachedTop =300
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =2910
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
                    Top =945
                    Width =8670
                    Height =345
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
                    LayoutCachedTop =945
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =1290
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
