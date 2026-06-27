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
    Width =6675
    DatasheetFontHeight =11
    Left =4035
    Top =195
    RecSrcDt = Begin
        0x5247db600c06e540
    End
    RecordSource ="SELECT ScheduleCoLessee.CoLesseeLabel, Clients.clientCompanyName, ScheduleCoLess"
        "ee.CoLesseeTaxID, ScheduleCoLessee.ScheduleID, ScheduleCoLessee.CoLesseeID, Sche"
        "duleCoLessee.HideOnCOA FROM ScheduleCoLessee INNER JOIN Clients ON ScheduleCoLes"
        "see.CoLesseeID = Clients.clientID WHERE (((ScheduleCoLessee.HideOnCOA)=False));"
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
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =300
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1260
                    Width =5414
                    Height =300
                    BorderColor =0
                    ForeColor =0
                    Name ="SubLesseeName"
                    ControlSource ="clientCompanyName"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1260
                    LayoutCachedWidth =6674
                    LayoutCachedHeight =300
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
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Width =1260
                    Height =300
                    TabIndex =1
                    BorderColor =0
                    ForeColor =0
                    Name ="SubLesseeLabel"
                    ControlSource ="=IIf(IsNull([CoLesseeLabel]),Null,[CoLesseeLabel] & \":\")"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =1260
                    LayoutCachedHeight =300
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
