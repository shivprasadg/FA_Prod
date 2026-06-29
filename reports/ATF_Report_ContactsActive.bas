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
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =291
    Left =3060
    Top =2880
    Filter ="([ATF_Report_ContactsActive2].[SchID]=932)"
    RecSrcDt = Begin
        0x8a5d3e8b15f6e440
    End
    RecordSource ="ATF_Report_ContactsActive2"
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
            ControlSource ="SchID"
        End
        Begin BreakLevel
            ControlSource ="Sorter"
        End
        Begin BreakLevel
            ControlSource ="FullName"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =552
            Name ="GroupHeader0"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =1
                    Left =5844
                    Top =336
                    Width =1476
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label246"
                    Caption ="Signator Level"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5844
                    LayoutCachedTop =336
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =552
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =1
                    Left =60
                    Top =336
                    Width =2445
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label247"
                    Caption ="Contact"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =60
                    LayoutCachedTop =336
                    LayoutCachedWidth =2505
                    LayoutCachedHeight =552
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    Left =7320
                    Top =336
                    Width =516
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label248"
                    Caption ="Email"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7320
                    LayoutCachedTop =336
                    LayoutCachedWidth =7836
                    LayoutCachedHeight =552
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    Left =7836
                    Top =336
                    Width =1452
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label27"
                    Caption ="Office Phone"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7836
                    LayoutCachedTop =336
                    LayoutCachedWidth =9288
                    LayoutCachedHeight =552
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =5
                    OldBorderStyle =1
                    TextAlign =1
                    Width =11448
                    Height =330
                    FontWeight =700
                    BorderColor =7500402
                    ForeColor =0
                    Name ="Label15"
                    Caption ="Lessee Contacts"
                    LayoutCachedWidth =11448
                    LayoutCachedHeight =330
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    Left =9288
                    Top =336
                    Width =1452
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label269"
                    Caption ="Mobile"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9288
                    LayoutCachedTop =336
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =552
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =1
                    Left =2508
                    Top =336
                    Width =3336
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label289"
                    Caption ="Title"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2508
                    LayoutCachedTop =336
                    LayoutCachedWidth =5844
                    LayoutCachedHeight =552
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
            End
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
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5844
                    Width =1476
                    Height =230
                    FontSize =8
                    TabIndex =2
                    BorderColor =14211288
                    ForeColor =0
                    Name ="ChildDesc"
                    ControlSource ="SignerType"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =5844
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
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
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Width =2445
                    Height =230
                    FontSize =8
                    BorderColor =14211288
                    ForeColor =0
                    Name ="VendorSNm"
                    ControlSource ="FullName"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =60
                    LayoutCachedWidth =2505
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
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
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7320
                    Width =516
                    Height =230
                    FontSize =8
                    TabIndex =3
                    BorderColor =14211288
                    ForeColor =0
                    Name ="ChildPartCost"
                    ControlSource ="HasEmail"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =7320
                    LayoutCachedWidth =7836
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
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
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7836
                    Width =1452
                    Height =230
                    FontSize =8
                    TabIndex =4
                    BorderColor =14211288
                    ForeColor =0
                    Name ="contactPrimaryPhone"
                    ControlSource ="Phone"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =7836
                    LayoutCachedWidth =9288
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
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
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9288
                    Width =1452
                    Height =230
                    FontSize =8
                    TabIndex =5
                    BorderColor =14211288
                    ForeColor =0
                    Name ="Text270"
                    ControlSource ="Mobile"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =9288
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
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
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2508
                    Width =3336
                    Height =230
                    FontSize =8
                    TabIndex =1
                    BorderColor =14211288
                    ForeColor =0
                    Name ="Text283"
                    ControlSource ="Title"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2508
                    LayoutCachedWidth =5844
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
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
    End
End
CodeBehindForm
' See "ATF_Report_ContactsActive.cls"
