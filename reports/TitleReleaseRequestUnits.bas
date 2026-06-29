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
    Width =10692
    DatasheetFontHeight =11
    ItemSuffix =121
    Left =4815
    Top =5745
    RecSrcDt = Begin
        0xe4394815348ce640
    End
    RecordSource ="SELECT vw_CustodianTitleTrust2.VIN, vw_CustodianTitleTrust2.OrigAssetID, * FROM "
        "Pick_Units_WorkTable INNER JOIN vw_CustodianTitleTrust2 ON Pick_Units_WorkTable."
        "UnitID = vw_CustodianTitleTrust2.OrigAssetID WHERE (((vw_CustodianTitleTrust2.uS"
        "tatus)<>\"B\")) ORDER BY vw_CustodianTitleTrust2.VIN, vw_CustodianTitleTrust2.Or"
        "igAssetID;"
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
        Begin UnboundObjectFrame
            OldBorderStyle =1
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
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="BankID"
        End
        Begin BreakLevel
            ControlSource ="VIN"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
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
            RepeatSection = NotDefault
            Height =855
            Name ="GroupHeader0"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =2292
                    Width =660
                    Height =855
                    BackColor =16183539
                    BorderColor =0
                    ForeColor =0
                    Name ="Label64"
                    Caption ="State\015\012of \015\012Title"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2292
                    LayoutCachedWidth =2952
                    LayoutCachedHeight =855
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderTint =100.0
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Width =2280
                    Height =855
                    BackColor =16183539
                    BorderColor =0
                    ForeColor =0
                    Name ="Label66"
                    Caption ="Vehicle\015\012Identification \015\012Number"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =855
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderTint =100.0
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =2964
                    Width =4320
                    Height =855
                    BackColor =16183539
                    BorderColor =0
                    ForeColor =0
                    Name ="Label67"
                    Caption ="Equipment\015\012Description"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2964
                    LayoutCachedWidth =7284
                    LayoutCachedHeight =855
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderTint =100.0
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =7296
                    Width =2076
                    Height =855
                    BackColor =16183539
                    BorderColor =0
                    ForeColor =0
                    Name ="Label68"
                    Caption ="Portfolio Interest Certificate No."
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7296
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =855
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderTint =100.0
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    Left =9384
                    Width =1296
                    Height =855
                    BackColor =6737151
                    BorderColor =0
                    ForeColor =0
                    Name ="Label69"
                    Caption ="Account No\015\012COLL ID\015\012FA-Asset\015\012"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9384
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =855
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderTint =100.0
                    ForeTint =100.0
                    GroupTable =5
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =306
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2292
                    Width =660
                    Height =306
                    FontSize =10
                    TabIndex =1
                    ForeColor =0
                    Name ="Text2"
                    ControlSource ="State of Title"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2292
                    LayoutCachedWidth =2952
                    LayoutCachedHeight =306
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Width =2280
                    Height =306
                    FontSize =10
                    ForeColor =0
                    Name ="Text17"
                    ControlSource ="VIN"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =2280
                    LayoutCachedHeight =306
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2964
                    Width =4320
                    Height =306
                    FontSize =10
                    TabIndex =2
                    ForeColor =0
                    Name ="Text43"
                    ControlSource ="Equipment Description"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2964
                    LayoutCachedWidth =7284
                    LayoutCachedHeight =306
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7296
                    Width =2076
                    Height =306
                    FontSize =10
                    TabIndex =3
                    ForeColor =0
                    Name ="Text44"
                    ControlSource ="Portfolio Certificate No"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7296
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =306
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9384
                    Width =1296
                    Height =306
                    FontSize =10
                    TabIndex =4
                    BackColor =13434879
                    ForeColor =0
                    Name ="Text45"
                    ControlSource ="OrigAssetID"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9384
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =306
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =100.0
                    GroupTable =4
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
            Height =0
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
