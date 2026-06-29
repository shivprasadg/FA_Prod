Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =20528
    DatasheetFontHeight =11
    ItemSuffix =44
    Left =435
    Top =4935
    Right =21915
    Bottom =15765
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x553a0aa39272e440
    End
    RecordSource ="SELECT * FROM ReqInvMSO WHERE (((ReqInvMSO.[Days2Delivery])<=15));"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
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
        Begin FormHeader
            Height =1950
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =1590
                    Width =19740
                    Height =360
                    BackColor =8210719
                    BorderColor =10855845
                    Name ="Box22"
                    HorizontalAnchor =2
                    LayoutCachedTop =1590
                    LayoutCachedWidth =19740
                    LayoutCachedHeight =1950
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    Left =16800
                    Top =510
                    Width =1950
                    Height =510
                    FontSize =20
                    FontWeight =700
                    ForeColor =-2147483614
                    Name ="Label33"
                    Caption ="AND / OR"
                    HorizontalAnchor =1
                    LayoutCachedLeft =16800
                    LayoutCachedTop =510
                    LayoutCachedWidth =18750
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =215
                    Left =132
                    Top =1680
                    Width =888
                    Height =240
                    ForeColor =16777215
                    Name ="Label5"
                    Caption ="FAID:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =132
                    LayoutCachedTop =1680
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =1920
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    Left =16860
                    Top =1680
                    Width =2988
                    Height =240
                    ForeColor =16777215
                    Name ="Label6"
                    Caption ="VIN:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =16860
                    LayoutCachedTop =1680
                    LayoutCachedWidth =19848
                    LayoutCachedHeight =1920
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =6060
                    Top =1680
                    Width =1065
                    Height =240
                    ForeColor =16777215
                    Name ="Label7"
                    Caption ="AssetID:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =6060
                    LayoutCachedTop =1680
                    LayoutCachedWidth =7125
                    LayoutCachedHeight =1920
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =7200
                    Top =1680
                    Width =630
                    Height =240
                    ForeColor =16777215
                    Name ="Label8"
                    Caption ="Year:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =7200
                    LayoutCachedTop =1680
                    LayoutCachedWidth =7830
                    LayoutCachedHeight =1920
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =7908
                    Top =1680
                    Width =1440
                    Height =240
                    ForeColor =16777215
                    Name ="Label9"
                    Caption ="Make:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =7908
                    LayoutCachedTop =1680
                    LayoutCachedWidth =9348
                    LayoutCachedHeight =1920
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    Left =9420
                    Top =1680
                    Width =1440
                    Height =240
                    ForeColor =16777215
                    Name ="Label10"
                    Caption ="Model:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =9420
                    LayoutCachedTop =1680
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =1920
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    Left =10932
                    Top =1680
                    Width =1596
                    Height =240
                    ForeColor =16777215
                    Name ="Label11"
                    Caption ="Type:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =10932
                    LayoutCachedTop =1680
                    LayoutCachedWidth =12528
                    LayoutCachedHeight =1920
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =4872
                    Top =1680
                    Width =1116
                    Height =240
                    ForeColor =16777215
                    Name ="Label12"
                    Caption ="Est Date"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =4872
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5988
                    LayoutCachedHeight =1920
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =3360
                    Top =1680
                    Width =1440
                    Height =240
                    ForeColor =16777215
                    Name ="Label14"
                    Caption ="Invoice No:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =3360
                    LayoutCachedTop =1680
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =1920
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    Left =12600
                    Top =1680
                    Width =615
                    Height =240
                    ForeColor =16777215
                    Name ="Label17"
                    Caption ="State:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =12600
                    LayoutCachedTop =1680
                    LayoutCachedWidth =13215
                    LayoutCachedHeight =1920
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =1680
                    Width =0
                    Height =240
                    ForeColor =16777215
                    Name ="Label18"
                    Caption ="Client:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =60
                    LayoutCachedTop =1680
                    LayoutCachedWidth =60
                    LayoutCachedHeight =1920
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    Left =13284
                    Top =1680
                    Width =1152
                    Height =240
                    ForeColor =16777215
                    Name ="Label19"
                    Caption ="SchNo:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =13284
                    LayoutCachedTop =1680
                    LayoutCachedWidth =14436
                    LayoutCachedHeight =1920
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =6180
                    Top =90
                    Width =1320
                    Height =1020
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8210719
                    Name ="MakeMSOEmail"
                    Caption ="Create MSO Request Email"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6180
                    LayoutCachedTop =90
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =1110
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =2010
                    Top =60
                    Width =4125
                    Height =1500
                    FontSize =10
                    BoundColumn =2
                    Name ="Client"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ReqInvMSO.Regarding AS Client, ReqInvMSO.[FAID/(PO)] AS FAID, ReqInvMSO.P"
                        "onum1 AS PO, Count(ReqInvMSO.AssetID) AS Units, ReqInvMSO.contactEmail FROM ReqI"
                        "nvMSO WHERE (((ReqInvMSO.Days2Delivery)<15)) GROUP BY ReqInvMSO.Regarding, ReqIn"
                        "vMSO.[FAID/(PO)], ReqInvMSO.Ponum1, ReqInvMSO.contactEmail;"
                    ColumnWidths ="1656;864;720;0;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2010
                    LayoutCachedTop =60
                    LayoutCachedWidth =6135
                    LayoutCachedHeight =1560
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =3
                            Left =930
                            Top =45
                            Width =1005
                            Height =855
                            FontWeight =700
                            ForeColor =5330263
                            Name ="Label3"
                            Caption ="Please \015\012Select ► \015\012 Company "
                            LayoutCachedLeft =930
                            LayoutCachedTop =45
                            LayoutCachedWidth =1935
                            LayoutCachedHeight =900
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =8895
                    Width =3150
                    Height =240
                    FontSize =8
                    ForeColor =-2147483617
                    Name ="Label28"
                    Caption =" (double-click to remove)"
                    LayoutCachedLeft =8895
                    LayoutCachedWidth =12045
                    LayoutCachedHeight =240
                End
                Begin Label
                    Vertical = NotDefault
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =3
                    Left =18885
                    Top =480
                    Width =750
                    Height =1230
                    FontSize =28
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label29"
                    Caption =" MSO"
                    HorizontalAnchor =1
                    LayoutCachedLeft =18885
                    LayoutCachedTop =480
                    LayoutCachedWidth =19635
                    LayoutCachedHeight =1710
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =3
                    Left =17085
                    Top =525
                    Width =1680
                    Height =480
                    FontSize =20
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label32"
                    Caption ="AND / OR"
                    HorizontalAnchor =1
                    LayoutCachedLeft =17085
                    LayoutCachedTop =525
                    LayoutCachedWidth =18765
                    LayoutCachedHeight =1005
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    SpecialEffect =2
                    OverlapFlags =215
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =7575
                    Top =240
                    Width =4485
                    Height =1320
                    FontSize =10
                    TabIndex =2
                    BackColor =15590879
                    Name ="ReqMSOTracking"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT FAID, Datesent, VehicleCount, Index FROM ReqInvMSO_Tracking WHERE FAID ='"
                        "14-635' AND ReqType = 'MSOCopies' Order BY Datesent;"
                    ColumnWidths ="1440;1152;1728;0"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =7575
                    LayoutCachedTop =240
                    LayoutCachedWidth =12060
                    LayoutCachedHeight =1560
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            Left =7575
                            Width =1290
                            Height =285
                            FontSize =9
                            FontWeight =700
                            ForeColor =-2147483615
                            Name ="Label26"
                            Caption ="Previous Emails:"
                            LayoutCachedLeft =7575
                            LayoutCachedWidth =8865
                            LayoutCachedHeight =285
                        End
                    End
                End
                Begin Label
                    Vertical = NotDefault
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =3
                    Left =18915
                    Top =465
                    Width =750
                    Height =1230
                    FontSize =28
                    FontWeight =700
                    ForeColor =-2147483617
                    Name ="Label35"
                    Caption =" MSO"
                    HorizontalAnchor =1
                    LayoutCachedLeft =18915
                    LayoutCachedTop =465
                    LayoutCachedWidth =19665
                    LayoutCachedHeight =1695
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =3
                    Left =11910
                    Top =675
                    Width =6915
                    Height =1185
                    FontSize =48
                    FontWeight =700
                    ForeColor =15527148
                    Name ="Label23"
                    Caption ="Missing Invoices"
                    HorizontalAnchor =1
                    LayoutCachedLeft =11910
                    LayoutCachedTop =675
                    LayoutCachedWidth =18825
                    LayoutCachedHeight =1860
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =3
                    Left =12120
                    Top =690
                    Width =6735
                    Height =1200
                    FontSize =48
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label24"
                    Caption ="Missing Invoices"
                    HorizontalAnchor =1
                    LayoutCachedLeft =12120
                    LayoutCachedTop =690
                    LayoutCachedWidth =18855
                    LayoutCachedHeight =1890
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =14508
                    Top =1680
                    Width =2280
                    Height =240
                    ForeColor =16777215
                    Name ="Label20"
                    Caption ="Vendor Contact:"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =14508
                    LayoutCachedTop =1680
                    LayoutCachedWidth =16788
                    LayoutCachedHeight =1920
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =17535
                    Top =1620
                    Width =720
                    Height =288
                    TabIndex =3
                    Name ="GoHome"
                    Caption ="GoHome"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaad000000000000adda0fff0bb0fff0da ,
                        0xad0fff0bb0fff0adda0fff0bb0fff0daad0fff0bb0fff0adda0fff0000fff0da ,
                        0xad0ffffffffff0add0b0ffffffff0b0aad0b0ffffff0b0addad0b0ffff0b0ada ,
                        0xadad0b0ff0b00daddadad0b00b010adaadadad0bb0a10daddadadad00ad10ada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="\"Open Main Menu\""
                    HorizontalAnchor =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17535
                    LayoutCachedTop =1620
                    LayoutCachedWidth =18255
                    LayoutCachedHeight =1908
                    Overlaps =1
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =120
                    Top =405
                    Width =780
                    Height =315
                    FontWeight =700
                    TabIndex =4
                    Name ="DaytoDelivery"
                    RowSourceType ="Value List"
                    RowSource ="5;10;15;20;30;60;120;180;365"
                    DefaultValue ="\"15\""
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =120
                    LayoutCachedTop =405
                    LayoutCachedWidth =900
                    LayoutCachedHeight =720
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =1
                            Left =120
                            Top =135
                            Width =780
                            Height =225
                            FontSize =9
                            Name ="Label37"
                            Caption ="Days To"
                            LayoutCachedLeft =120
                            LayoutCachedTop =135
                            LayoutCachedWidth =900
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =12240
                    Top =105
                    Width =1065
                    Height =270
                    FontSize =10
                    TabIndex =5
                    Name ="FAID"

                    LayoutCachedLeft =12240
                    LayoutCachedTop =105
                    LayoutCachedWidth =13305
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =215
                    Left =1092
                    Top =1680
                    Width =1308
                    Height =240
                    ForeColor =16777215
                    Name ="Label40"
                    Caption ="PO"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =1092
                    LayoutCachedTop =1680
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =1920
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =120
                    Top =750
                    Width =780
                    Height =240
                    FontSize =9
                    Name ="Label42"
                    Caption ="Delivery"
                    LayoutCachedLeft =120
                    LayoutCachedTop =750
                    LayoutCachedWidth =900
                    LayoutCachedHeight =990
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =18300
                    Top =1620
                    Width =1035
                    Height =288
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    ForeColor =-2147483617
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadad0dadadadadaadad00adadadadaddad030dadadadada ,
                        0xad0330adadadadad0033300000000adaa03330ff0dadadadd03300ff0adad4da ,
                        0xa03330ff0dad44add03330ff0ad44444a03330ff0d444444d03330ff0ad44444 ,
                        0xa0330fff0dad44add030ffff0adad4daa00fffff0dadadadd00000000adadada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    HorizontalAnchor =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =18300
                    LayoutCachedTop =1620
                    LayoutCachedWidth =19335
                    LayoutCachedHeight =1908
                    PictureCaptionArrangement =5
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =2472
                    Top =1680
                    Width =816
                    Height =240
                    ForeColor =16777215
                    Name ="Label43"
                    Caption ="Days"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =2472
                    LayoutCachedTop =1680
                    LayoutCachedWidth =3288
                    LayoutCachedHeight =1920
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin Section
            Height =386
            Name ="Detail"
            AlternateBackColor =-2147483606
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =132
                    Top =36
                    Width =888
                    Height =312
                    FontSize =10
                    TabIndex =1
                    Name ="FAID_PO"
                    ControlSource ="FAID/(PO)"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =132
                    LayoutCachedTop =36
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontFamily =50
                    IMESentenceMode =3
                    Left =16860
                    Top =36
                    Width =2988
                    Height =312
                    ColumnWidth =2805
                    FontSize =10
                    TabIndex =14
                    ForeColor =-2147483617
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    FontName ="OCR A Extended"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =16860
                    LayoutCachedTop =36
                    LayoutCachedWidth =19848
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6060
                    Top =36
                    Width =1065
                    Height =312
                    ColumnWidth =1125
                    FontSize =10
                    TabIndex =6
                    Name ="AssetID"
                    ControlSource ="AssetID"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =6060
                    LayoutCachedTop =36
                    LayoutCachedWidth =7125
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7200
                    Top =36
                    Width =630
                    Height =312
                    FontSize =10
                    TabIndex =7
                    Name ="MYr"
                    ControlSource ="MYr"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =7200
                    LayoutCachedTop =36
                    LayoutCachedWidth =7830
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =7908
                    Top =36
                    Height =312
                    FontSize =10
                    TabIndex =8
                    Name ="Make"
                    ControlSource ="Make"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =7908
                    LayoutCachedTop =36
                    LayoutCachedWidth =9348
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =9420
                    Top =36
                    Height =312
                    FontSize =10
                    TabIndex =9
                    Name ="Model"
                    ControlSource ="Model"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =9420
                    LayoutCachedTop =36
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =10932
                    Top =36
                    Width =1596
                    Height =312
                    ColumnWidth =2910
                    FontSize =10
                    TabIndex =10
                    Name ="Type"
                    ControlSource ="Type"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =10932
                    LayoutCachedTop =36
                    LayoutCachedWidth =12528
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4872
                    Top =36
                    Width =1116
                    Height =312
                    ColumnWidth =1755
                    FontSize =10
                    TabIndex =5
                    Name ="Est Delivelry Dt"
                    ControlSource ="Est Delivelry Dt"
                    Format ="mm/dd/yy"
                    EventProcPrefix ="Est_Delivelry_Dt"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =4872
                    LayoutCachedTop =36
                    LayoutCachedWidth =5988
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3360
                    Top =36
                    Height =312
                    ColumnWidth =2040
                    FontSize =10
                    TabIndex =4
                    Name ="Invoice No"
                    ControlSource ="Invoice No"
                    EventProcPrefix ="Invoice_No"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =3360
                    LayoutCachedTop =36
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =12600
                    Top =36
                    Width =615
                    Height =312
                    FontSize =10
                    TabIndex =11
                    Name ="Delivery St"
                    ControlSource ="Delivery St"
                    EventProcPrefix ="Delivery_St"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =12600
                    LayoutCachedTop =36
                    LayoutCachedWidth =13215
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =36
                    Width =0
                    Height =312
                    ColumnWidth =2115
                    Name ="ClientShNm"
                    ControlSource ="ClientShNm"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =60
                    LayoutCachedTop =36
                    LayoutCachedWidth =60
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =13284
                    Top =36
                    Width =1152
                    Height =312
                    ColumnWidth =2145
                    FontSize =10
                    TabIndex =12
                    Name ="SchNo"
                    ControlSource ="SchNo"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =13284
                    LayoutCachedTop =36
                    LayoutCachedWidth =14436
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =14508
                    Top =36
                    Width =2280
                    Height =312
                    ColumnWidth =4380
                    FontSize =10
                    TabIndex =13
                    Name ="VendorContact"
                    ControlSource ="VendorContact"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =14508
                    LayoutCachedTop =36
                    LayoutCachedWidth =16788
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =1092
                    Top =36
                    Width =1308
                    Height =312
                    FontSize =10
                    TabIndex =2
                    Name ="Text39"
                    ControlSource ="PONum1"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1092
                    LayoutCachedTop =36
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2472
                    Top =36
                    Width =816
                    Height =312
                    FontSize =10
                    TabIndex =3
                    Name ="Days2Delivery"
                    ControlSource ="Days2Delivery"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2472
                    LayoutCachedTop =36
                    LayoutCachedWidth =3288
                    LayoutCachedHeight =348
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin FormFooter
            Height =360
            Name ="FormFooter"
            AutoHeight =1
        End
    End
End
CodeBehindForm
' See "ReqInvMSO_Form.cls"
