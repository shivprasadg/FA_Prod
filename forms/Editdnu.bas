Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7800
    DatasheetFontHeight =11
    ItemSuffix =13
    Left =2955
    Top =2775
    Right =11040
    Bottom =5790
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x8599b657f662e440
    End
    RecordSource ="SELECT Vendors.VendorID, Vendors.VendorName, Vendors.VENDORDBA, Vendors.VendorAd"
        "dress, Vendors.VendorCity, Vendors.VendorState, Vendors.VendorZip, Vendors.Vendo"
        "rPmtType FROM Vendors WHERE (((Vendors.VendorID)=[forms]![Invoice - Add/Edit]![v"
        "endornm]));"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin PageHeader
            DisplayWhen =1
            Height =360
            Name ="PageHeaderSection"
            AutoHeight =1
        End
        Begin Section
            Height =3000
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =1560
                    Top =540
                    Width =3000
                    ColumnWidth =5715
                    FontSize =9
                    Name ="VendorName"
                    ControlSource ="VendorName"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =540
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =780
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =5520
                    Top =540
                    Width =2220
                    FontSize =9
                    TabIndex =1
                    Name ="VENDORDBA"
                    ControlSource ="VENDORDBA"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =540
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =780
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextFontCharSet =238
                            Left =4680
                            Top =540
                            Width =660
                            Height =240
                            FontSize =9
                            BorderColor =12632256
                            Name ="Label8"
                            Caption ="D/B/A:"
                            LayoutCachedLeft =4680
                            LayoutCachedTop =540
                            LayoutCachedWidth =5340
                            LayoutCachedHeight =780
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =1560
                    Top =840
                    Width =3000
                    ColumnWidth =2670
                    FontSize =9
                    TabIndex =2
                    Name ="VendorAddress"
                    ControlSource ="VendorAddress"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =840
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =1080
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =1560
                    Top =1140
                    Width =4380
                    FontSize =9
                    TabIndex =3
                    Name ="Text10"
                    ControlSource ="=UCase([vendorcity] & \" \" & [vendorstate] & \" \" & [vendorzip])"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =1140
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =1380
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =3720
                    Top =1500
                    Height =315
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    Name ="VendorPmtType"
                    ControlSource ="VendorPmtType"

                    LayoutCachedLeft =3720
                    LayoutCachedTop =1500
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =1815
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            Left =1860
                            Top =1500
                            Width =1740
                            Height =315
                            FontSize =9
                            FontWeight =700
                            Name ="Label12"
                            Caption ="Payment Method:"
                            LayoutCachedLeft =1860
                            LayoutCachedTop =1500
                            LayoutCachedWidth =3600
                            LayoutCachedHeight =1815
                        End
                    End
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =360
            Name ="PageFooterSection"
            AutoHeight =1
        End
    End
End
