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
    Width =10800
    DatasheetFontHeight =11
    ItemSuffix =247
    Left =2760
    Top =2580
    RecSrcDt = Begin
        0xfe8941795687e640
    End
    RecordSource ="SELECT ChildParts.ATFUnitSpecID, ChildParts.ATFChildSpecID, ATFChildPartSpec.ATF"
        "ID, [ChildParts].[ChildType] & \" \" & [ChildParts].[ChildYear] & \" \" & [Child"
        "Parts].[ChildMake] & \" \" & [ChildParts].[ChildClass] AS ChildDesc, Sum(ChildPa"
        "rts.childFinalCost) AS ChildPartCost, Vendors.VendorID AS ChildVendorID, Vendors"
        ".VendorSNm FROM (ATFChildPartSpec RIGHT JOIN ChildParts ON ATFChildPartSpec.ATFC"
        "hildSpecID = ChildParts.ATFChildSpecID) LEFT JOIN Vendors ON ChildParts.ChildVen"
        "dorREF = Vendors.VendorID GROUP BY ChildParts.ATFUnitSpecID, ChildParts.ATFChild"
        "SpecID, ATFChildPartSpec.ATFID, [ChildParts].[ChildType] & \" \" & [ChildParts]."
        "[ChildYear] & \" \" & [ChildParts].[ChildMake] & \" \" & [ChildParts].[ChildClas"
        "s], Vendors.VendorID, Vendors.VendorSNm;"
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="ATFUnitSpecID"
        End
        Begin BreakLevel
            ControlSource ="ChildDesc"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="GroupHeader0"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =230
            Name ="Detail"
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1125
                    Width =5100
                    Height =230
                    FontSize =8
                    BorderColor =12566463
                    ForeColor =0
                    Name ="ChildDesc"
                    ControlSource ="ChildDesc"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1125
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6225
                    Width =2850
                    Height =230
                    FontSize =8
                    TabIndex =1
                    BorderColor =12566463
                    ForeColor =0
                    Name ="VendorSNm"
                    ControlSource ="VendorSNm"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =6225
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Width =600
                    Height =0
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =2
                    BorderColor =14277081
                    ForeColor =0
                    Name ="ATFGroupSpecID"
                    ControlSource ="ATFChildSpecID"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =180
                    LayoutCachedWidth =780
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9072
                    Width =1110
                    Height =230
                    FontSize =8
                    TabIndex =3
                    BorderColor =12566463
                    ForeColor =0
                    Name ="ChildPartCost"
                    ControlSource ="ChildPartCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =9072
                    LayoutCachedWidth =10182
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Visible = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Width =600
                    Height =0
                    FontSize =9
                    TabIndex =4
                    BorderColor =14277081
                    ForeColor =0
                    Name ="ATFUnitSpecID"
                    ControlSource ="ATFUnitSpecID"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =600
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =255
                    Width =840
                    Height =216
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =5
                    BorderColor =16777215
                    Name ="COmmLBL"
                    ControlSource ="=IIf(Nz([ChildVendorId],0)=0,Null,\"Part\")"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =255
                    LayoutCachedWidth =1095
                    LayoutCachedHeight =216
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Line
                    Left =216
                    Width =0
                    Height =230
                    BorderColor =6108695
                    Name ="Line245"
                    LayoutCachedLeft =216
                    LayoutCachedWidth =216
                    LayoutCachedHeight =230
                    BorderThemeColorIndex =-1
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="GroupFooter0"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
