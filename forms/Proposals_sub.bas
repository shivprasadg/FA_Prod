Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5880
    DatasheetFontHeight =11
    ItemSuffix =49
    Left =3345
    Top =6045
    Right =9765
    Bottom =10065
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x33085e18d696e340
    End
    RecordSource ="Quotes"
    Caption ="Quotes_sub"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    NavigationCaption ="Proposals"
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =238
            FontSize =11
            BorderColor =-2147483609
            ForeColor =11830108
            FontName ="Calibri"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =-2147483609
        End
        Begin CommandButton
            TextFontCharSet =238
            FontSize =11
            FontWeight =400
            ForeColor =-2147483609
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =-2147483609
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin BoundObjectFrame
            SizeMode =3
            BorderLineStyle =0
            BackStyle =0
            LabelX =-1800
            BorderColor =-2147483609
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin CustomControl
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin ToggleButton
            TextFontCharSet =238
            FontSize =11
            FontWeight =400
            ForeColor =-2147483609
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =238
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Attachment
            BackStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =4800
            Height =3840
            BorderColor =-2147483609
            LabelX =-1800
        End
        Begin FormHeader
            Visible = NotDefault
            Height =285
            BackColor =-2147483609
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3180
                    Width =900
                    Height =285
                    ColumnWidth =1440
                    FontSize =10
                    ForeColor =255
                    Name ="dealOptID"
                    ControlSource ="dealOptID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =3180
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =285
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =2100
                            Width =1020
                            Height =285
                            FontSize =10
                            Name ="dealOptID_Label"
                            Caption ="dealOptID"
                            FontName ="Tahoma"
                            LayoutCachedLeft =2100
                            LayoutCachedWidth =3120
                            LayoutCachedHeight =285
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1140
                    Width =900
                    Height =285
                    ColumnWidth =1515
                    FontSize =10
                    TabIndex =1
                    ForeColor =255
                    Name ="dealRef"
                    ControlSource ="dealRef"
                    StatusBarText ="link to deals table"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1140
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =285
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =60
                            Width =1020
                            Height =285
                            FontSize =10
                            Name ="dealRef_Label"
                            Caption ="dealRef"
                            FontName ="Tahoma"
                            LayoutCachedLeft =60
                            LayoutCachedWidth =1080
                            LayoutCachedHeight =285
                        End
                    End
                End
            End
        End
        Begin Section
            Height =1740
            BackColor =16249583
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1260
                    Top =780
                    Width =1020
                    Height =255
                    ColumnWidth =1515
                    FontSize =9
                    TabIndex =4
                    BorderColor =15062992
                    Name ="dealLeaseTerm"
                    ControlSource ="dealLeaseTerm"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1260
                    LayoutCachedTop =780
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =1035
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =3
                            Left =180
                            Top =780
                            Width =1020
                            Height =255
                            FontSize =9
                            ForeColor =0
                            Name ="dealLeaseTerm_Label"
                            Caption ="Lease term"
                            FontName ="Tahoma"
                            LayoutCachedLeft =180
                            LayoutCachedTop =780
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =1035
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =4
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1260
                    Top =1080
                    Width =1020
                    Height =255
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =5
                    BorderColor =15062992
                    Name ="dealLRF"
                    ControlSource ="dealLRF"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1260
                    LayoutCachedTop =1080
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =1335
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =3
                            Left =180
                            Top =1080
                            Width =1020
                            Height =255
                            FontSize =9
                            ForeColor =0
                            Name ="dealLRF_Label"
                            Caption ="LRF"
                            FontName ="Tahoma"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1080
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =1335
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1260
                    Top =1380
                    Width =1020
                    Height =255
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =6
                    BorderColor =15062992
                    Name ="dealPayment"
                    ControlSource ="dealPayment"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1260
                    LayoutCachedTop =1380
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =1635
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =3
                            Left =180
                            Top =1380
                            Width =1020
                            Height =255
                            FontSize =9
                            ForeColor =0
                            Name ="dealPayment_Label"
                            Caption ="Payment"
                            FontName ="Tahoma"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1380
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =1635
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4200
                    Top =1080
                    Width =1500
                    Height =255
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =7
                    BorderColor =15062992
                    Name ="dealOverMiCharge"
                    ControlSource ="dealOverMiCharge"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Tahoma"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =1080
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =1335
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =3
                            Left =2580
                            Top =1080
                            Width =1575
                            Height =255
                            FontSize =9
                            ForeColor =0
                            Name ="dealOverMiCharge_Label"
                            Caption ="Over Mi/Hr Charge"
                            FontName ="Tahoma"
                            LayoutCachedLeft =2580
                            LayoutCachedTop =1080
                            LayoutCachedWidth =4155
                            LayoutCachedHeight =1335
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4200
                    Top =1380
                    Width =1500
                    Height =255
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =8
                    BorderColor =15062992
                    Name ="dealAdvArr"
                    ControlSource ="dealAdvArr"
                    RowSourceType ="Value List"
                    RowSource ="Advance;Arrears"
                    StatusBarText ="drop down"
                    FontName ="Tahoma"
                    AllowValueListEdits =0

                    LayoutCachedLeft =4200
                    LayoutCachedTop =1380
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =1635
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =3
                            Left =2580
                            Top =1380
                            Width =1560
                            Height =255
                            FontSize =9
                            ForeColor =0
                            Name ="dealAdvArr_Label"
                            Caption ="Adv/Arr"
                            FontName ="Tahoma"
                            LayoutCachedLeft =2580
                            LayoutCachedTop =1380
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =1635
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    ListRows =8
                    ListWidth =1440
                    Left =180
                    Top =60
                    Width =1320
                    Height =255
                    ColumnWidth =3000
                    FontSize =9
                    Name ="dealIndex"
                    ControlSource ="dealIndex"
                    RowSourceType ="Value List"
                    RowSource ="\"SWAP\";\"T. Bill\""
                    ColumnWidths ="1440"
                    StatusBarText ="drop down"
                    FontName ="Tahoma"
                    AllowValueListEdits =0

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2940
                    Top =60
                    Width =900
                    Height =255
                    ColumnWidth =1515
                    FontSize =9
                    TabIndex =2
                    Name ="dealYears"
                    ControlSource ="dealYears"
                    FontName ="Tahoma"

                    LayoutCachedLeft =2940
                    LayoutCachedTop =60
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =2940
                            Top =360
                            Width =900
                            Height =255
                            FontSize =9
                            ForeColor =9605778
                            Name ="dealYears_Label"
                            Caption ="Years"
                            FontName ="Tahoma"
                            LayoutCachedLeft =2940
                            LayoutCachedTop =360
                            LayoutCachedWidth =3840
                            LayoutCachedHeight =615
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =4
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3900
                    Top =60
                    Width =1320
                    Height =255
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =3
                    Name ="dealRate"
                    ControlSource ="dealRate"
                    FontName ="Tahoma"

                    LayoutCachedLeft =3900
                    LayoutCachedTop =60
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =3900
                            Top =360
                            Width =1095
                            Height =255
                            FontSize =9
                            ForeColor =9605778
                            Name ="dealRate_Label"
                            Caption ="Rate"
                            FontName ="Tahoma"
                            LayoutCachedLeft =3900
                            LayoutCachedTop =360
                            LayoutCachedWidth =4995
                            LayoutCachedHeight =615
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1560
                    Top =60
                    Width =1320
                    Height =255
                    ColumnWidth =2940
                    FontSize =9
                    TabIndex =1
                    Name ="dealDate"
                    ControlSource ="dealDate"
                    Format ="General Date"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =60
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =1560
                            Top =360
                            Width =1095
                            Height =255
                            FontSize =9
                            ForeColor =9605778
                            Name ="dealDate_Label"
                            Caption ="Date"
                            FontName ="Tahoma"
                            LayoutCachedLeft =1560
                            LayoutCachedTop =360
                            LayoutCachedWidth =2655
                            LayoutCachedHeight =615
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    Left =180
                    Top =360
                    Width =1095
                    Height =255
                    FontSize =9
                    ForeColor =9605778
                    Name ="dealIndex_Label"
                    Caption ="Index"
                    FontName ="Tahoma"
                    LayoutCachedLeft =180
                    LayoutCachedTop =360
                    LayoutCachedWidth =1275
                    LayoutCachedHeight =615
                End
                Begin Line
                    OverlapFlags =85
                    Left =2460
                    Top =720
                    Width =0
                    Height =960
                    Name ="Line48"
                    LayoutCachedLeft =2460
                    LayoutCachedTop =720
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =1680
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
        End
    End
End
