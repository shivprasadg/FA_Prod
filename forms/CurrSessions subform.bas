Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10230
    DatasheetFontHeight =11
    ItemSuffix =8
    Left =7800
    Top =3030
    Right =14145
    Bottom =10380
    Filter ="(CurrSessions.Time_Out Is Null Or CurrSessions.Time_Out=\"\")"
    OrderBy ="[CurrSessions].[User_name], [CurrSessions].[Time_Out], [CurrSessions].[Time_In] "
        "DESC"
    RecSrcDt = Begin
        0xe66234927ac0e540
    End
    RecordSource ="CurrSessions"
    Caption ="CurrSessions subform"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin Section
            Height =2370
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =360
                    Width =7260
                    Height =600
                    ColumnWidth =1395
                    Name ="User_name"
                    ControlSource ="User_name"

                    LayoutCachedLeft =2910
                    LayoutCachedTop =360
                    LayoutCachedWidth =10170
                    LayoutCachedHeight =960
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =360
                            Width =2460
                            Height =330
                            Name ="User_name_Label"
                            Caption ="User_name"
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =690
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =1080
                    Width =1620
                    Height =330
                    ColumnWidth =2355
                    TabIndex =1
                    Name ="Time_In"
                    ControlSource ="Time_In"

                    LayoutCachedLeft =2910
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4530
                    LayoutCachedHeight =1410
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1080
                            Width =2460
                            Height =330
                            Name ="Time_In_Label"
                            Caption ="Time_In"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1080
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1410
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =1500
                    Width =3660
                    Height =330
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="Time_Out"
                    ControlSource ="Time_Out"

                    LayoutCachedLeft =2910
                    LayoutCachedTop =1500
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =1830
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1500
                            Width =2460
                            Height =330
                            Name ="Time_Out_Label"
                            Caption ="Time_Out"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1500
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1830
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =1920
                    Width =1050
                    Height =330
                    ColumnWidth =1050
                    TabIndex =3
                    Name ="FirstOfActive User"
                    ControlSource ="FirstOfActive User"
                    EventProcPrefix ="FirstOfActive_User"

                    LayoutCachedLeft =2910
                    LayoutCachedTop =1920
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =2250
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1920
                            Width =2460
                            Height =330
                            Name ="FirstOfActive User_Label"
                            Caption ="FirstOfActive User"
                            EventProcPrefix ="FirstOfActive_User_Label"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1920
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =2250
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
