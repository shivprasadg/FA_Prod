Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =163
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11715
    DatasheetFontHeight =11
    ItemSuffix =148
    Left =8355
    Top =3240
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0xb3583b508ca8e540
    End
    RecordSource ="SELECT vw_BuyersOrder.*, * FROM vw_BuyersOrder WHERE (((vw_BuyersOrder.InvoiceRe"
        "questID)=[Forms]![ReMarketInvoiceManager]![InvoiceRequestID]));"
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
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
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
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
            ControlSource ="BuyerID"
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
        End
        Begin Section
            KeepTogether = NotDefault
            Height =0
            Name ="Detail"
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =225
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8400
                    Width =3030
                    Height =225
                    FontSize =9
                    Name ="Text47"
                    ControlSource ="=\"Page: \" & [Page]"

                    LayoutCachedLeft =8400
                    LayoutCachedWidth =11430
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =4185
                    Height =225
                    FontSize =9
                    TabIndex =1
                    BackColor =15527148
                    Name ="Text45"
                    ControlSource ="=\"Printed on \" & Date() & \" @ \" & Time()"

                    LayoutCachedWidth =4185
                    LayoutCachedHeight =225
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =3660
            Name ="ReportFooter"
            Begin
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Width =11640
                    Height =2100
                    Name ="MoveORderPickUpPage"
                    SourceObject ="Report.MoveORderPickUpPage"

                    LayoutCachedWidth =11640
                    LayoutCachedHeight =2100
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =2220
                    Width =11640
                    TabIndex =1
                    Name ="MoveOrderDeliveryPage"
                    SourceObject ="Report.MoveOrderDeliveryPage"

                    LayoutCachedTop =2220
                    LayoutCachedWidth =11640
                    LayoutCachedHeight =3660
                End
                Begin PageBreak
                    Top =2160
                    Name ="PageBreak144"
                End
            End
        End
    End
End
CodeBehindForm
' See "MoveOrderRelease.cls"
