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
    Width =31680
    DatasheetFontHeight =11
    ItemSuffix =72
    Left =840
    Top =8625
    RecSrcDt = Begin
        0x883bdd9c911fe640
    End
    RecordSource ="SELECT * FROM vw_PayProceeds_Report WHERE (((vw_PayProceeds_Report.[SchID]) In ("
        "2220,2446)));"
    Caption ="vw_PayProceeds_Report"
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
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            GroupOn =1
            GroupInterval =3
            ControlSource ="Sch"
        End
        Begin BreakLevel
            ControlSource ="Type"
        End
        Begin BreakLevel
            ControlSource ="PayInterimItems"
        End
        Begin BreakLevel
            ControlSource ="Sch"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =960
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =4125
                    Height =540
                    FontSize =20
                    Name ="Label69"
                    Caption ="vw_PayProceeds_Report"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =4185
                    LayoutCachedHeight =600
                End
            End
        End
        Begin PageHeader
            Height =1260
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =1
                    Left =360
                    Top =60
                    Width =2085
                    Height =315
                    Name ="Sch by 3 Initial Letters_Label"
                    Caption ="Sch by 3 Initial Letters"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Sch_by_3_Initial_Letters_Label"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =1680
                    Top =60
                    Width =525
                    Height =315
                    Name ="Type_Label"
                    Caption ="Type"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =1680
                    LayoutCachedTop =60
                    LayoutCachedWidth =2205
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =900
                    Top =60
                    Width =1620
                    Height =315
                    Name ="PayInterimItems_Label"
                    Caption ="PayInterimItems"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =900
                    LayoutCachedTop =60
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =2580
                    Top =60
                    Width =630
                    Height =315
                    Name ="Client_Label"
                    Caption ="Client"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =2580
                    LayoutCachedTop =60
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =3600
                    Top =60
                    Width =390
                    Height =315
                    Name ="Sch_Label"
                    Caption ="Sch"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =3600
                    LayoutCachedTop =60
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =4440
                    Top =60
                    Width =600
                    Height =315
                    Name ="Items_Label"
                    Caption ="Items"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =4440
                    LayoutCachedTop =60
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =5100
                    Top =60
                    Width =510
                    Height =315
                    Name ="Desc_Label"
                    Caption ="Desc"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =5100
                    LayoutCachedTop =60
                    LayoutCachedWidth =5610
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =7020
                    Top =60
                    Width =840
                    Height =315
                    Name ="Invoices_Label"
                    Caption ="Invoices"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =7020
                    LayoutCachedTop =60
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =9420
                    Top =60
                    Width =480
                    Height =315
                    Name ="Cost_Label"
                    Caption ="Cost"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =9420
                    LayoutCachedTop =60
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =9165
                    Top =60
                    Width =1095
                    Height =315
                    Name ="PayInterim_Label"
                    Caption ="PayInterim"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =9165
                    LayoutCachedTop =60
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =9525
                    Top =60
                    Width =1095
                    Height =315
                    Name ="PayVendor_Label"
                    Caption ="PayVendor"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =9525
                    LayoutCachedTop =60
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =10335
                    Top =60
                    Width =645
                    Height =315
                    Name ="PayFA_Label"
                    Caption ="PayFA"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =10335
                    LayoutCachedTop =60
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =11040
                    Top =60
                    Width =420
                    Height =315
                    Name ="VendorID_Label"
                    Caption ="VendorID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =11040
                    LayoutCachedTop =60
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =11520
                    Top =60
                    Width =420
                    Height =315
                    Name ="SchID_Label"
                    Caption ="SchID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =11520
                    LayoutCachedTop =60
                    LayoutCachedWidth =11940
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =12000
                    Top =60
                    Width =420
                    Height =315
                    Name ="iFbankID_Label"
                    Caption ="iFbankID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =12000
                    LayoutCachedTop =60
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =12480
                    Top =60
                    Width =420
                    Height =315
                    Name ="FACashID_Label"
                    Caption ="FACashID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =12480
                    LayoutCachedTop =60
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =12960
                    Top =60
                    Width =420
                    Height =315
                    Name ="PayVendorID_Label"
                    Caption ="PayVendorID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =12960
                    LayoutCachedTop =60
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =13440
                    Top =60
                    Width =420
                    Height =315
                    Name ="PayVendorItems_Label"
                    Caption ="PayVendorItems"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =13440
                    LayoutCachedTop =60
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =13920
                    Top =60
                    Width =420
                    Height =315
                    Name ="PayFAItems_Label"
                    Caption ="PayFAItems"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =13920
                    LayoutCachedTop =60
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =14400
                    Top =60
                    Width =1860
                    Height =315
                    Name ="VendorPayInst_Label"
                    Caption ="VendorPayInst"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =14400
                    LayoutCachedTop =60
                    LayoutCachedWidth =16260
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =16320
                    Top =60
                    Width =1860
                    Height =315
                    Name ="InterimPayInst_Label"
                    Caption ="InterimPayInst"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =16320
                    LayoutCachedTop =60
                    LayoutCachedWidth =18180
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =18240
                    Top =60
                    Width =1860
                    Height =315
                    Name ="FACashInst_Label"
                    Caption ="FACashInst"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =18240
                    LayoutCachedTop =60
                    LayoutCachedWidth =20100
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =20160
                    Top =60
                    Width =1860
                    Height =315
                    Name ="SyndBankLegalName_Label"
                    Caption ="SyndBankLegalName"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =20160
                    LayoutCachedTop =60
                    LayoutCachedWidth =22020
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =22080
                    Top =60
                    Width =1860
                    Height =315
                    Name ="SyndBankAddr_Label"
                    Caption ="SyndBankAddr"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =22080
                    LayoutCachedTop =60
                    LayoutCachedWidth =23940
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =24000
                    Top =60
                    Width =960
                    Height =315
                    Name ="SyndBank_Label"
                    Caption ="SyndBank"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =24000
                    LayoutCachedTop =60
                    LayoutCachedWidth =24960
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =25020
                    Top =60
                    Width =960
                    Height =315
                    Name ="SchShortDesc_Label"
                    Caption ="SchShortDesc"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =25020
                    LayoutCachedTop =60
                    LayoutCachedWidth =25980
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =26040
                    Top =60
                    Width =960
                    Height =315
                    Name ="MLNo_Label"
                    Caption ="MLNo"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =26040
                    LayoutCachedTop =60
                    LayoutCachedWidth =27000
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =27060
                    Top =60
                    Width =420
                    Height =315
                    Name ="clientID_Label"
                    Caption ="clientID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =27060
                    LayoutCachedTop =60
                    LayoutCachedWidth =27480
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =27540
                    Top =60
                    Width =4080
                    Height =315
                    Name ="clientCompanyName_Label"
                    Caption ="clientCompanyName"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    LayoutCachedLeft =27540
                    LayoutCachedTop =60
                    LayoutCachedWidth =31620
                    LayoutCachedHeight =375
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =375
            Name ="GroupHeader0"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =360
                    Width =960
                    Height =315
                    Name ="Sch by 3 Initial Letters"
                    ControlSource ="=Left$([Sch], 3)"
                    EventProcPrefix ="Sch_by_3_Initial_Letters"

                    LayoutCachedLeft =360
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =315
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =390
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =1680
                    Width =360
                    Height =315
                    Name ="Type"
                    ControlSource ="Type"

                    LayoutCachedLeft =1680
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =2100
                    Width =420
                    Height =315
                    TabIndex =1
                    Name ="PayInterimItems"
                    ControlSource ="PayInterimItems"

                    LayoutCachedLeft =2100
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =2580
                    Width =960
                    Height =315
                    ColumnWidth =3675
                    TabIndex =2
                    Name ="Client"
                    ControlSource ="Client"

                    LayoutCachedLeft =2580
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =3600
                    Width =960
                    Height =315
                    TabIndex =3
                    Name ="Sch"
                    ControlSource ="Sch"

                    LayoutCachedLeft =3600
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =4620
                    Width =420
                    Height =315
                    TabIndex =4
                    Name ="Items"
                    ControlSource ="Items"

                    LayoutCachedLeft =4620
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =5100
                    Width =1860
                    Height =315
                    ColumnWidth =6270
                    TabIndex =5
                    Name ="Desc"
                    ControlSource ="Desc"

                    LayoutCachedLeft =5100
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =7020
                    Width =1860
                    Height =315
                    ColumnWidth =10440
                    TabIndex =6
                    Name ="Invoices"
                    ControlSource ="Invoices"

                    LayoutCachedLeft =7020
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =8940
                    Width =960
                    Height =315
                    ColumnWidth =1620
                    TabIndex =7
                    Name ="Cost"
                    ControlSource ="Cost"

                    LayoutCachedLeft =8940
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =9960
                    Width =300
                    Height =315
                    ColumnWidth =1620
                    TabIndex =8
                    Name ="PayInterim"
                    ControlSource ="PayInterim"

                    LayoutCachedLeft =9960
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =10320
                    Width =300
                    Height =315
                    ColumnWidth =3585
                    TabIndex =9
                    Name ="PayVendor"
                    ControlSource ="PayVendor"

                    LayoutCachedLeft =10320
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =10680
                    Width =300
                    Height =315
                    TabIndex =10
                    Name ="PayFA"
                    ControlSource ="PayFA"

                    LayoutCachedLeft =10680
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =11040
                    Width =420
                    Height =330
                    TabIndex =11
                    Name ="VendorID"
                    ControlSource ="VendorID"

                    LayoutCachedLeft =11040
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =11520
                    Width =420
                    Height =330
                    TabIndex =12
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =11520
                    LayoutCachedWidth =11940
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =12000
                    Width =420
                    Height =330
                    TabIndex =13
                    Name ="iFbankID"
                    ControlSource ="iFbankID"

                    LayoutCachedLeft =12000
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =12480
                    Width =420
                    Height =330
                    TabIndex =14
                    Name ="FACashID"
                    ControlSource ="FACashID"

                    LayoutCachedLeft =12480
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =12960
                    Width =420
                    Height =330
                    TabIndex =15
                    Name ="PayVendorID"
                    ControlSource ="PayVendorID"

                    LayoutCachedLeft =12960
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =13440
                    Width =420
                    Height =330
                    TabIndex =16
                    Name ="PayVendorItems"
                    ControlSource ="PayVendorItems"

                    LayoutCachedLeft =13440
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Width =420
                    Height =330
                    TabIndex =17
                    Name ="PayFAItems"
                    ControlSource ="PayFAItems"

                    LayoutCachedLeft =13920
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =14400
                    Width =1860
                    Height =330
                    ColumnWidth =3150
                    TabIndex =18
                    Name ="VendorPayInst"
                    ControlSource ="VendorPayInst"

                    LayoutCachedLeft =14400
                    LayoutCachedWidth =16260
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =16320
                    Width =1860
                    Height =330
                    ColumnWidth =4020
                    TabIndex =19
                    Name ="InterimPayInst"
                    ControlSource ="InterimPayInst"

                    LayoutCachedLeft =16320
                    LayoutCachedWidth =18180
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =18240
                    Width =1860
                    Height =330
                    TabIndex =20
                    Name ="FACashInst"
                    ControlSource ="FACashInst"

                    LayoutCachedLeft =18240
                    LayoutCachedWidth =20100
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =20160
                    Width =1860
                    Height =330
                    TabIndex =21
                    Name ="SyndBankLegalName"
                    ControlSource ="SyndBankLegalName"

                    LayoutCachedLeft =20160
                    LayoutCachedWidth =22020
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =22080
                    Width =1860
                    Height =330
                    TabIndex =22
                    Name ="SyndBankAddr"
                    ControlSource ="SyndBankAddr"

                    LayoutCachedLeft =22080
                    LayoutCachedWidth =23940
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =24000
                    Width =960
                    Height =330
                    TabIndex =23
                    Name ="SyndBank"
                    ControlSource ="SyndBank"

                    LayoutCachedLeft =24000
                    LayoutCachedWidth =24960
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =25020
                    Width =960
                    Height =330
                    TabIndex =24
                    Name ="SchShortDesc"
                    ControlSource ="SchShortDesc"

                    LayoutCachedLeft =25020
                    LayoutCachedWidth =25980
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =26040
                    Width =960
                    Height =330
                    TabIndex =25
                    Name ="MLNo"
                    ControlSource ="MLNo"

                    LayoutCachedLeft =26040
                    LayoutCachedWidth =27000
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =27060
                    Width =420
                    Height =330
                    TabIndex =26
                    Name ="clientID"
                    ControlSource ="clientID"

                    LayoutCachedLeft =27060
                    LayoutCachedWidth =27480
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =27540
                    Width =4080
                    Height =330
                    TabIndex =27
                    Name ="clientCompanyName"
                    ControlSource ="clientCompanyName"

                    LayoutCachedLeft =27540
                    LayoutCachedWidth =31620
                    LayoutCachedHeight =330
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =690
            Name ="GroupFooter1"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =360
                    Width =31260
                    Name ="Text58"
                    ControlSource ="=\"Summary for \" & \"'Sch' = \" & \" \" & [Sch] & \" (\" & Count(*) & \" \" & I"
                        "If(Count(*) = 1, \"detail record\", \"detail records\") & \")\""

                    LayoutCachedLeft =360
                    LayoutCachedWidth =31620
                    LayoutCachedHeight =240
                End
                Begin Label
                    Left =360
                    Top =315
                    Width =480
                    Height =315
                    Name ="Label59"
                    Caption ="Sum"
                    LayoutCachedLeft =360
                    LayoutCachedTop =315
                    LayoutCachedWidth =840
                    LayoutCachedHeight =630
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9960
                    Top =315
                    Width =300
                    Height =315
                    TabIndex =1
                    Name ="Sum Of PayInterim"
                    ControlSource ="=Sum([PayInterim])"
                    EventProcPrefix ="Sum_Of_PayInterim"

                    LayoutCachedLeft =9960
                    LayoutCachedTop =315
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =630
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =10320
                    Top =315
                    Width =300
                    Height =315
                    TabIndex =2
                    Name ="Sum Of PayVendor"
                    ControlSource ="=Sum([PayVendor])"
                    EventProcPrefix ="Sum_Of_PayVendor"

                    LayoutCachedLeft =10320
                    LayoutCachedTop =315
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =630
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =10680
                    Top =315
                    Width =300
                    Height =315
                    TabIndex =3
                    Name ="Sum Of PayFA"
                    ControlSource ="=Sum([PayFA])"
                    EventProcPrefix ="Sum_Of_PayFA"

                    LayoutCachedLeft =10680
                    LayoutCachedTop =315
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =630
                End
            End
        End
        Begin PageFooter
            Height =570
            Name ="PageFooterSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =240
                    Width =5040
                    Height =330
                    Name ="Text70"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                    LayoutCachedLeft =60
                    LayoutCachedTop =240
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =570
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =26580
                    Top =240
                    Width =5040
                    Height =330
                    TabIndex =1
                    Name ="Text71"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =26580
                    LayoutCachedTop =240
                    LayoutCachedWidth =31620
                    LayoutCachedHeight =570
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =390
            Name ="ReportFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =9960
                    Width =300
                    Height =330
                    Name ="PayInterim Grand Total Sum"
                    ControlSource ="=Sum([PayInterim])"
                    EventProcPrefix ="PayInterim_Grand_Total_Sum"

                    LayoutCachedLeft =9960
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =330
                End
                Begin Label
                    Left =60
                    Width =1140
                    Height =315
                    Name ="Label61"
                    Caption ="Grand Total"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =10320
                    Width =300
                    Height =330
                    TabIndex =1
                    Name ="PayVendor Grand Total Sum"
                    ControlSource ="=Sum([PayVendor])"
                    EventProcPrefix ="PayVendor_Grand_Total_Sum"

                    LayoutCachedLeft =10320
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =330
                End
                Begin Label
                    Left =60
                    Width =1140
                    Height =315
                    Name ="Label64"
                    Caption ="Grand Total"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =10680
                    Width =300
                    Height =330
                    TabIndex =2
                    Name ="PayFA Grand Total Sum"
                    ControlSource ="=Sum([PayFA])"
                    EventProcPrefix ="PayFA_Grand_Total_Sum"

                    LayoutCachedLeft =10680
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =330
                End
                Begin Label
                    Left =60
                    Width =1140
                    Height =315
                    Name ="Label67"
                    Caption ="Grand Total"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =315
                End
            End
        End
    End
End
