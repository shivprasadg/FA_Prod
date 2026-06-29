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
    Width =11460
    DatasheetFontHeight =11
    ItemSuffix =27
    Left =4245
    Top =975
    RecSrcDt = Begin
        0x687783a210e8e440
    End
    RecordSource ="ATF_Report_GuarantorsActive"
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =72
            Name ="Detail"
            AlternateBackColor =16777215
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =924
                    Width =4320
                    Height =72
                    FontSize =9
                    BorderColor =14277081
                    ForeColor =0
                    Name ="LesseeName"
                    ControlSource ="GName"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =924
                    LayoutCachedWidth =5244
                    LayoutCachedHeight =72
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
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5820
                    Width =1920
                    Height =72
                    FontSize =9
                    TabIndex =1
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text13"
                    ControlSource ="clientTaxID"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5820
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =72
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
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
                    Left =10740
                    Width =600
                    Height =0
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =2
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text25"
                    ControlSource ="ATFID"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10740
                    LayoutCachedWidth =11340
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
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Width =885
                    Height =72
                    FontSize =8
                    TabIndex =3
                    BorderColor =14277081
                    Name ="Label304"
                    ControlSource ="=IIf(Nz([Gname],'')='',Null,\"Guarantor\")"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =60
                    LayoutCachedWidth =945
                    LayoutCachedHeight =72
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5244
                    Width =576
                    Height =72
                    FontSize =8
                    TabIndex =4
                    BorderColor =14277081
                    Name ="Label11"
                    ControlSource ="=IIf(Nz([Gname],'')='',Null,\"Tax ID\")"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5244
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =72
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
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
