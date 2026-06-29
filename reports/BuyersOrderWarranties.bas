Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11700
    DatasheetFontHeight =11
    ItemSuffix =1
    Top =600
    RecSrcDt = Begin
        0x2849d784557ee540
    End
    RecordSource ="Select * From vw_RMKT_BuyersOrdersWarranty"
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
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =285
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3948
                    Width =5280
                    Height =285
                    FontSize =10
                    BorderColor =0
                    ForeColor =0
                    Name ="DescriptionLong"
                    ControlSource ="wDesc"
                    StatusBarText ="unit make"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3948
                    LayoutCachedWidth =9228
                    LayoutCachedHeight =285
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10320
                    Width =1380
                    Height =285
                    FontSize =10
                    TabIndex =1
                    BorderColor =-2147483617
                    ForeColor =-2147483617
                    Name ="WarrantyPrice"
                    ControlSource ="WarrantyPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10320
                    LayoutCachedWidth =11700
                    LayoutCachedHeight =285
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9228
                    Width =1092
                    Height =285
                    FontSize =10
                    TabIndex =2
                    BorderColor =0
                    ForeColor =0
                    Name ="TargetSale"
                    Format ="Standard"
                    StatusBarText ="price Rmk Sm should get for unit"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9228
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =285
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1848
                    Width =2100
                    Height =285
                    FontSize =10
                    TabIndex =3
                    BorderColor =0
                    ForeColor =0
                    Name ="UnitVIN"
                    Format =">"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1848
                    LayoutCachedWidth =3948
                    LayoutCachedHeight =285
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Width =765
                    Height =285
                    FontSize =10
                    TabIndex =4
                    BorderColor =0
                    ForeColor =0
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    StatusBarText ="initials of Remarketing Salesperson"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =765
                    LayoutCachedHeight =285
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =768
                    Width =1080
                    Height =285
                    FontSize =10
                    TabIndex =5
                    BorderColor =0
                    ForeColor =0
                    Name ="Text45"
                    ControlSource ="BuyersOrderID"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =768
                    LayoutCachedWidth =1848
                    LayoutCachedHeight =285
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
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
