Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =163
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11505
    DatasheetFontHeight =11
    ItemSuffix =198
    Left =7095
    Top =5700
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x8a8ad5e5d6fee440
    End
    RecordSource ="PO_ChildParts"
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
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="PONumber"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="GroupHeader0"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4470
                    Width =1140
                    Height =0
                    LeftMargin =72
                    BorderColor =9211020
                    Name ="PONumber"
                    ControlSource ="PONumber"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4470
                    LayoutCachedWidth =5610
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =30
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9084
                    Width =588
                    Height =30
                    TabIndex =2
                    BorderColor =9211020
                    Name ="Text38"
                    ControlSource ="CountOfChildPartID"
                    GroupTable =17
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9084
                    LayoutCachedWidth =9672
                    LayoutCachedHeight =30
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =17
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7632
                    Height =30
                    TabIndex =1
                    BorderColor =9211020
                    Name ="Text39"
                    ControlSource ="childFinalCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =17
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7632
                    LayoutCachedWidth =9072
                    LayoutCachedHeight =30
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =17
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9684
                    Width =1740
                    Height =30
                    TabIndex =3
                    BorderColor =9211020
                    Name ="Text40"
                    ControlSource ="=[CountOfChildPartID]*[childFinalCost]"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =17
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9684
                    LayoutCachedWidth =11424
                    LayoutCachedHeight =30
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =17
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =1
                    IMESentenceMode =3
                    Left =72
                    Width =7548
                    Height =30
                    LeftMargin =72
                    BorderColor =9211020
                    Name ="Text34"
                    ControlSource ="PartDescription"
                    GroupTable =17
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =72
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =30
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =17
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
