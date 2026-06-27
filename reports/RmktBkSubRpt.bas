Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    ControlBox = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =0
    TabularCharSet =163
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8295
    DatasheetFontHeight =11
    ItemSuffix =10
    Left =4875
    Top =5715
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x5f5f613d1832e440
    End
    RecordSource ="RmktDirUnitq"
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
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =330
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Width =750
                    Height =30
                    Name ="clientID"
                    ControlSource ="clientID"

                    LayoutCachedWidth =750
                    LayoutCachedHeight =30
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15
                    Width =3600
                    Height =270
                    FontSize =10
                    TabIndex =1
                    BackColor =-2147483610
                    Name ="Purchased"
                    ControlSource ="Purchased"
                    Format =">"

                    LayoutCachedLeft =15
                    LayoutCachedWidth =3615
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3600
                    Width =4545
                    Height =270
                    FontSize =10
                    TabIndex =2
                    Name ="Text8"
                    ControlSource ="=IIf(IsNull([engine series]),Null,\"Eng: \" & [engine series] & \" \" & [engine "
                        "capacity])"

                    LayoutCachedLeft =3600
                    LayoutCachedWidth =8145
                    LayoutCachedHeight =270
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
