Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =20
    GridY =24
    Width =4950
    DatasheetFontHeight =11
    ItemSuffix =3
    Left =6840
    Top =5445
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x55810b8892fbe540
    End
    RecordSource ="AssetIDBySchID"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =255
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
            AutoLabel = NotDefault
            AddColon = NotDefault
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            Height =480
            LabelX =-1800
            FontSize =18
            FontName ="Arial"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            ControlSource ="AssetID"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            Height =960
            Name ="Detail"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =288
                    Top =180
                    Width =4374
                    Height =540
                    FontSize =20
                    Name ="Text2"
                    ControlSource ="AssetID"

                    LayoutCachedLeft =288
                    LayoutCachedTop =180
                    LayoutCachedWidth =4662
                    LayoutCachedHeight =720
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
