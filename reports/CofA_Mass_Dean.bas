Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11618
    DatasheetFontHeight =11
    ItemSuffix =192
    Left =5265
    Top =10635
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x0b448af3d172e440
    End
    RecordSource ="CofA_Report_Mass"
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
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
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
        Begin BreakLevel
            ControlSource ="DelSt2"
        End
        Begin BreakLevel
            ControlSource ="UnitID"
        End
        Begin BreakLevel
            ControlSource ="vin4"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="grpcd"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =8310
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =8745
                    Height =315
                    ColumnOrder =0
                    TabIndex =6
                    Name ="SGrpID"
                    ControlSource ="SGrpID"

                    LayoutCachedLeft =8745
                    LayoutCachedWidth =10185
                    LayoutCachedHeight =315
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =900
                    Top =1635
                    Width =1095
                    Height =315
                    FontSize =12
                    Name ="Label6"
                    Caption ="Lessor:"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =900
                    LayoutCachedTop =1635
                    LayoutCachedWidth =1995
                    LayoutCachedHeight =1950
                End
                Begin Label
                    TextFontFamily =18
                    Left =2040
                    Top =1635
                    Width =3180
                    Height =315
                    FontSize =12
                    Name ="Label7"
                    Caption ="Fleet Advantage, LLC"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =2040
                    LayoutCachedTop =1635
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =1950
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2070
                    Top =2085
                    Width =8640
                    Height =315
                    FontSize =12
                    TabIndex =8
                    Name ="AddressBox"
                    ControlSource ="=[clientCompanyName] & \", an \" & [CorpState] & \" corporation\""
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2070
                    LayoutCachedTop =2085
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =2400
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    IMESentenceMode =3
                    Left =900
                    Top =2085
                    Width =1125
                    Height =315
                    FontSize =12
                    TabIndex =12
                    Name ="LesseeBox"
                    ControlSource ="=\"Lessee:\""
                    FontName ="Times New Roman"

                    LayoutCachedLeft =900
                    LayoutCachedTop =2085
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =2400
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Top =285
                    Width =10755
                    Height =360
                    FontWeight =700
                    Name ="Label168"
                    Caption ="CERTIFICATE OF ACCEPTANCE"
                    FontName ="Times New Roman"
                    LayoutCachedTop =285
                    LayoutCachedWidth =10755
                    LayoutCachedHeight =645
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =870
                    Top =2655
                    Width =9825
                    Height =315
                    FontSize =12
                    TabIndex =5
                    Name ="Text142"
                    ControlSource ="=\"Master Vehicle Lease Agreement No. <strong>\" & [MLNo] & \" </strong> dated a"
                        "s of  \" & Format([MLDt],\"mmmm d\\,yyyy\") & \", (the <strong>''<u>Lease</u>''<"
                        "/strong>)\""
                    FontName ="Times New Roman"
                    TextFormat =1

                    LayoutCachedLeft =870
                    LayoutCachedTop =2655
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =2970
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =855
                    Top =3060
                    Width =9825
                    Height =315
                    FontSize =12
                    TabIndex =11
                    Name ="Text169"
                    ControlSource ="=\"Schedule No: <strong>\" & [SchNo] & \"</strong> dated as of \" & Format([Sche"
                        "dDt],\"mmmm d\"\", \"\"yyyy\") & \" (the <strong>''<u>Schedule</u>''</strong>)\""
                    FontName ="Times New Roman"
                    TextFormat =1

                    LayoutCachedLeft =855
                    LayoutCachedTop =3060
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =3375
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =870
                    Top =3570
                    Width =9825
                    Height =315
                    FontSize =12
                    TabIndex =3
                    Name ="Text170"
                    ControlSource ="=\"Lessee certifies to Lessor (without prejudice to any rights Lessee or Lessor "
                        "may have against the manufacturer, supplier or any other third party) that, on t"
                        "he Certificate Date set forth above, the Equipment set forth and described in <s"
                        "trong><u>Exhibit A</u></strong> attached hereto and incorporated herein by this "
                        "reference: \""
                    FontName ="Times New Roman"
                    TextFormat =1

                    LayoutCachedLeft =870
                    LayoutCachedTop =3570
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =3885
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =900
                    Top =1140
                    Width =1800
                    Height =315
                    FontSize =12
                    Name ="Label172"
                    Caption ="Certificate Date: "
                    FontName ="Times New Roman"
                    LayoutCachedLeft =900
                    LayoutCachedTop =1140
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =1455
                End
                Begin TextBox
                    TextFontFamily =18
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2745
                    Top =1155
                    Width =3975
                    Height =345
                    FontSize =12
                    TabIndex =10
                    Name ="docdate"
                    ControlSource ="=[AskMonth] & \" _____, 2014\""
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2745
                    LayoutCachedTop =1155
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =1500
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =1740
                    Top =4050
                    Width =8955
                    Height =315
                    FontSize =12
                    TabIndex =4
                    Name ="Text174"
                    ControlSource ="=\"(1) has been delivered to Lessee, inspected by an agent of Lessee, found to b"
                        "e in good operating condition and repair and accepted under the Lease, (2) is av"
                        "ailable for use and service by Lessee, (3) is of a size, design, function and ma"
                        "nufacture selected by Lessee, and (4) is suitable for Lessee’s intended purposes"
                        ".\""
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1740
                    LayoutCachedTop =4050
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =4365
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =855
                    Top =4545
                    Width =9825
                    Height =315
                    FontSize =12
                    TabIndex =2
                    Name ="Text175"
                    ControlSource ="=\"Lessee represents and warrants to Lessor that, as of the Certificate Date set"
                        " forth above and with respect to the Equipment set forth and described in <stron"
                        "g><u>Exhibit A</u></strong> attached hereto:\""
                    FontName ="Times New Roman"
                    TextFormat =1

                    LayoutCachedLeft =855
                    LayoutCachedTop =4545
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =4860
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =1740
                    Top =5025
                    Width =8955
                    Height =315
                    FontSize =12
                    TabIndex =7
                    Name ="Text176"
                    ControlSource ="=\"(1) the representations and warranties of Lessee set forth in the Lease and S"
                        "chedule are true and correct in all material respects, (2) Lessee has satisfied "
                        "or complied with all requirements set forth in the Lease and Schedule to be sati"
                        "sfied or complied with on or prior to the date hereof, and (3) Lessee has obtain"
                        "ed, and there are in full force and effect, such insurance policies with respect"
                        " to such Equipment as are required to be obtained under the Lease and Schedule.\""
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1740
                    LayoutCachedTop =5025
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =5340
                End
                Begin Label
                    TextFontFamily =18
                    Left =870
                    Top =6300
                    Width =4650
                    Height =615
                    FontSize =12
                    Name ="Label177"
                    Caption ="FLEET ADVANTAGE, LLC\015\012(Lessor)"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =870
                    LayoutCachedTop =6300
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =6915
                End
                Begin Label
                    TextFontFamily =18
                    Left =6045
                    Top =6300
                    Width =4110
                    Height =615
                    FontSize =12
                    Name ="Label178"
                    Caption ="DEAN TRANSPORTATION, INC\015\012(Lessee)"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =6045
                    LayoutCachedTop =6300
                    LayoutCachedWidth =10155
                    LayoutCachedHeight =6915
                End
                Begin Label
                    TextFontFamily =18
                    Left =1575
                    Top =7215
                    Width =4065
                    Height =345
                    FontSize =12
                    Name ="Label179"
                    Caption ="_________________________________"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =1575
                    LayoutCachedTop =7215
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =7560
                End
                Begin Label
                    TextFontFamily =18
                    Left =1575
                    Top =7530
                    Width =4065
                    Height =345
                    FontSize =12
                    Name ="Label180"
                    Caption ="_________________________________"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =1575
                    LayoutCachedTop =7530
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =7875
                End
                Begin Label
                    TextFontFamily =18
                    Left =1575
                    Top =7830
                    Width =4065
                    Height =345
                    FontSize =12
                    Name ="Label181"
                    Caption ="_________________________________"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =1575
                    LayoutCachedTop =7830
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =8175
                End
                Begin Label
                    TextFontFamily =18
                    Left =855
                    Top =7230
                    Width =690
                    Height =915
                    FontSize =12
                    Name ="Label182"
                    Caption ="By:\015\012Name:\015\012Title:"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =855
                    LayoutCachedTop =7230
                    LayoutCachedWidth =1545
                    LayoutCachedHeight =8145
                End
                Begin Label
                    TextFontFamily =18
                    Left =6720
                    Top =7260
                    Width =4065
                    Height =345
                    FontSize =12
                    Name ="Label183"
                    Caption ="_________________________________"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =6720
                    LayoutCachedTop =7260
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =7605
                End
                Begin Label
                    TextFontFamily =18
                    Left =6000
                    Top =7275
                    Width =690
                    Height =885
                    FontSize =12
                    Name ="Label186"
                    Caption ="By:\015\012Name:\015\012Title:"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =6000
                    LayoutCachedTop =7275
                    LayoutCachedWidth =6690
                    LayoutCachedHeight =8160
                End
                Begin PageBreak
                    Top =8310
                    Name ="PageBreak187"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =855
                    Top =5625
                    Width =9825
                    Height =315
                    FontSize =12
                    TabIndex =9
                    Name ="Text188"
                    ControlSource ="=\"IN WITNESS WHEREOF, each of the undersigned has caused this Certificate of Ac"
                        "ceptance to be duly executed by its authorized signatory as of the Certificate D"
                        "ate set forth above.\""
                    FontName ="Times New Roman"
                    TextFormat =1

                    LayoutCachedLeft =855
                    LayoutCachedTop =5625
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =5940
                End
                Begin TextBox
                    TextFontFamily =18
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6720
                    Top =7575
                    Width =4065
                    Height =315
                    FontSize =12
                    TabIndex =1
                    Name ="ClientSignFullName"
                    ControlSource ="ClientSignFullName"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =7575
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =7890
                End
                Begin TextBox
                    TextFontFamily =18
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6720
                    Top =7860
                    Width =4065
                    Height =315
                    FontSize =12
                    Name ="ClientSignTitle"
                    ControlSource ="ClientSignTitle"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =7860
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =8175
                End
            End
        End
        Begin PageHeader
            Height =1470
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    BackStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =4188
                    Top =972
                    Width =996
                    Height =468
                    FontSize =9
                    BackColor =10855845
                    Name ="Label157"
                    Caption ="\015\012 Type"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =4188
                    LayoutCachedTop =972
                    LayoutCachedWidth =5184
                    LayoutCachedHeight =1440
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =36
                    Top =972
                    Width =732
                    Height =468
                    FontSize =9
                    BackColor =10855845
                    Name ="Label158"
                    Caption ="\015\012 UnitNo:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =36
                    LayoutCachedTop =972
                    LayoutCachedWidth =768
                    LayoutCachedHeight =1440
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =840
                    Top =972
                    Width =1875
                    Height =468
                    FontSize =9
                    BackColor =10855845
                    Name ="Label159"
                    Caption ="\015\012 VIN:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =840
                    LayoutCachedTop =972
                    LayoutCachedWidth =2715
                    LayoutCachedHeight =1440
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =2784
                    Top =972
                    Width =1332
                    Height =468
                    FontSize =9
                    BackColor =10855845
                    Name ="Label160"
                    Caption ="\015\012 Make:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =2784
                    LayoutCachedTop =972
                    LayoutCachedWidth =4116
                    LayoutCachedHeight =1440
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =5256
                    Top =972
                    Width =588
                    Height =468
                    FontSize =9
                    BackColor =10855845
                    Name ="Label161"
                    Caption ="\015\012Year"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =5256
                    LayoutCachedTop =972
                    LayoutCachedWidth =5844
                    LayoutCachedHeight =1440
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =7164
                    Top =972
                    Width =1860
                    Height =468
                    FontSize =9
                    BackColor =10855845
                    Name ="Label163"
                    Caption ="Delivery | Domicile Address"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =7164
                    LayoutCachedTop =972
                    LayoutCachedWidth =9024
                    LayoutCachedHeight =1440
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =9096
                    Top =972
                    Width =492
                    Height =468
                    FontSize =9
                    BackColor =10855845
                    Name ="Label164"
                    Caption ="Title State"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =9096
                    LayoutCachedTop =972
                    LayoutCachedWidth =9588
                    LayoutCachedHeight =1440
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =9660
                    Top =972
                    Width =1860
                    Height =468
                    FontSize =9
                    BackColor =10855845
                    Name ="Label165"
                    Caption ="Registration \015\012Address"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =9660
                    LayoutCachedTop =972
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =1440
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    Left =5916
                    Top =972
                    Width =1176
                    Height =468
                    FontSize =9
                    BackColor =10855845
                    Name ="Label167"
                    Caption ="\015\012 Model"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =5916
                    LayoutCachedTop =972
                    LayoutCachedWidth =7092
                    LayoutCachedHeight =1440
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Width =11490
                    Height =315
                    FontWeight =700
                    Name ="Label2"
                    Caption ="EXHIBIT A"
                    LayoutCachedWidth =11490
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =240
                    Width =11355
                    Height =315
                    Name ="2ndheader"
                    ControlSource ="=\"To Certificate of Acceptance Relating to Master Vehicle Lease Agreement No. <"
                        "strong>\" & [MLNo] & \" </strong> Schedule No: <strong>\" & [SchNo] & \"</strong"
                        ">\""
                    EventProcPrefix ="Ctl2ndheader"
                    TextFormat =1

                    LayoutCachedLeft =60
                    LayoutCachedTop =240
                    LayoutCachedWidth =11415
                    LayoutCachedHeight =555
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =75
                    Top =645
                    Width =11355
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="Text190"
                    ControlSource ="=\"<u> DESCRIPTION OF EQUIPMENT</u>\""
                    TextFormat =1

                    LayoutCachedLeft =75
                    LayoutCachedTop =645
                    LayoutCachedWidth =11430
                    LayoutCachedHeight =945
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            BreakLevel =3
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =10590
                    Width =180
                    Height =0
                    Name ="grpcd"
                    ControlSource ="grpcd"

                    LayoutCachedLeft =10590
                    LayoutCachedWidth =10770
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =482
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =36
                    Top =36
                    Width =732
                    Height =408
                    FontSize =8
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BorderColor =9605778
                    Name ="UnitID"
                    ControlSource ="unitunitnum"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =768
                    LayoutCachedHeight =444
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =840
                    Top =36
                    Width =1875
                    Height =408
                    ColumnWidth =2130
                    FontSize =9
                    TabIndex =1
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BorderColor =9605778
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =840
                    LayoutCachedTop =36
                    LayoutCachedWidth =2715
                    LayoutCachedHeight =444
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =5256
                    Top =36
                    Width =588
                    Height =408
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BorderColor =9605778
                    Name ="MLYr"
                    ControlSource ="MLYr"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =5256
                    LayoutCachedTop =36
                    LayoutCachedWidth =5844
                    LayoutCachedHeight =444
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4188
                    Top =36
                    Width =996
                    Height =408
                    FontSize =9
                    TabIndex =8
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BorderColor =9605778
                    Name ="UnitType"
                    ControlSource ="UnitType"
                    StatusBarText ="UNIT TYPE"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =4188
                    LayoutCachedTop =36
                    LayoutCachedWidth =5184
                    LayoutCachedHeight =444
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2784
                    Top =36
                    Width =1332
                    Height =408
                    FontSize =9
                    TabIndex =2
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BorderColor =9605778
                    Name ="UnitMake"
                    ControlSource ="UnitMake"
                    StatusBarText ="unit make"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2784
                    LayoutCachedTop =36
                    LayoutCachedWidth =4116
                    LayoutCachedHeight =444
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9660
                    Top =36
                    Width =1860
                    Height =408
                    FontSize =9
                    TabIndex =7
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =9605778
                    Name ="Text116"
                    ControlSource ="RegLoc"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =9660
                    LayoutCachedTop =36
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =444
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7164
                    Top =36
                    Width =1860
                    Height =408
                    FontSize =9
                    TabIndex =5
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =9605778
                    Name ="Text121"
                    ControlSource ="DeliveryLoc"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =7164
                    LayoutCachedTop =36
                    LayoutCachedWidth =9024
                    LayoutCachedHeight =444
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9096
                    Top =36
                    Width =492
                    Height =408
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BorderColor =9605778
                    Name ="TitledState"
                    ControlSource ="TitledState"
                    StatusBarText ="lookup in lookup_state table"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =9096
                    LayoutCachedTop =36
                    LayoutCachedWidth =9588
                    LayoutCachedHeight =444
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =5916
                    Top =36
                    Width =1176
                    Height =408
                    FontSize =9
                    TabIndex =4
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BorderColor =9605778
                    Name ="Text166"
                    ControlSource ="UnitModel"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =5916
                    LayoutCachedTop =36
                    LayoutCachedWidth =7092
                    LayoutCachedHeight =444
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            BreakLevel =3
            Name ="GroupFooter0"
        End
        Begin PageFooter
            Height =255
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7845
                    Width =3600
                    ColumnOrder =0
                    FontSize =9
                    Name ="Text66"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =7845
                    LayoutCachedWidth =11445
                    LayoutCachedHeight =240
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "CofA_Mass_Dean.cls"
