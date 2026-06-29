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
    Width =15354
    DatasheetFontHeight =11
    ItemSuffix =143
    Left =25395
    Top =2460
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0xb2a1254b4dc5e440
    End
    RecordSource ="RMKTInventoryFinal"
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
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            KeepTogether =2
            ControlSource ="HDG"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            GroupHeader = NotDefault
            KeepTogether =2
            ControlSource ="unitstatus"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =2
            ControlSource ="clientshnm"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            KeepTogether =2
            ControlSource ="LocSrt"
        End
        Begin BreakLevel
            KeepTogether =2
            ControlSource ="mlno"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="schno"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="SGrpNO"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =840
            Name ="ReportHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    BackStyle =0
                    SizeMode =0
                    PictureType =2
                    Top =660
                    Width =15336
                    Height =180
                    Name ="Image142"
                    Picture ="1_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =660
                    LayoutCachedWidth =15336
                    LayoutCachedHeight =840
                    TabIndex =3
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Rectangle
                    Top =555
                    Width =15336
                    Height =150
                    BackColor =10319446
                    BorderColor =10319446
                    Name ="Box30"
                    LayoutCachedTop =555
                    LayoutCachedWidth =15336
                    LayoutCachedHeight =705
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15
                    Top =15
                    Width =7275
                    Height =765
                    FontSize =32
                    FontWeight =700
                    TabIndex =1
                    ForeColor =16777215
                    Name ="Label31"
                    ControlSource ="=\"Remarketing Inventory\""

                    LayoutCachedLeft =15
                    LayoutCachedTop =15
                    LayoutCachedWidth =7290
                    LayoutCachedHeight =780
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7935
                    Top =525
                    Width =4545
                    Height =210
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    ForeColor =16777215
                    Name ="Text71"
                    ControlSource ="=\"| \" & Date() & \" | Units: \" & Count([unitvin]) & \" | Internal Use Only |\""

                    LayoutCachedLeft =7935
                    LayoutCachedTop =525
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =735
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =7260
                    Height =765
                    FontSize =32
                    FontWeight =700
                    ForeColor =10319446
                    Name ="Label231"
                    ControlSource ="=\"Remarketing Inventory \""

                    LayoutCachedWidth =7260
                    LayoutCachedHeight =765
                End
            End
        End
        Begin PageHeader
            Height =0
            BackColor =16768393
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =435
            BackColor =16249583
            Name ="GroupHeader5"
            Begin
                Begin TextBox
                    TextFontCharSet =186
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =45
                    Width =7155
                    Height =360
                    ColumnWidth =1965
                    FontSize =14
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =-2147483615
                    Name ="HDG2"
                    ControlSource ="=[Header] & \" | Units: \" & Count([UnitVIN]) & \" |\""

                    LayoutCachedLeft =60
                    LayoutCachedTop =45
                    LayoutCachedWidth =7215
                    LayoutCachedHeight =405
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =315
            BreakLevel =1
            Name ="GroupHeader2"
            Begin
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =405
                    Width =6780
                    Height =315
                    FontSize =12
                    Name ="Text45"
                    ControlSource ="=IIf([notfa]=1,\"BROKERED UNITS\",\"FA UNITS\")"

                    LayoutCachedLeft =405
                    LayoutCachedWidth =7185
                    LayoutCachedHeight =315
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =420
            BreakLevel =2
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =390
                    Width =10410
                    Height =420
                    FontSize =14
                    FontWeight =700
                    ForeColor =8210719
                    Name ="ClientShNmd"
                    ControlSource ="=\"♦\" & [ClientShNm] & \" | Units: \" & Count([UnitVIN]) & \" | \" & [Header] &"
                        " \"\""
                    StatusBarText ="Short Name for Reporting"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =390
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Visible = NotDefault
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11025
                    Top =15
                    TabIndex =1
                    Name ="clientID"
                    ControlSource ="clientID"

                    LayoutCachedLeft =11025
                    LayoutCachedTop =15
                    LayoutCachedWidth =12465
                    LayoutCachedHeight =255
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =294
            BreakLevel =3
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =660
                    Width =8760
                    Height =285
                    FontSize =12
                    FontWeight =700
                    Name ="LocationHDR"
                    ControlSource ="=[Location] & \"\""

                    LayoutCachedLeft =660
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =285
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            BreakLevel =5
            Name ="GroupHeader1"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =542
            BreakLevel =6
            Name ="GroupHeader4"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8376
                    Width =6960
                    Height =540
                    FontSize =9
                    BackColor =15527148
                    BorderColor =15527148
                    ForeColor =8210719
                    Name ="Text43"
                    ControlSource ="=\"Schdule: \" & ([SCHNO]) & \" | Group: \" & [SGRPNO] & \" |  Units: \" & Count"
                        "([unitvin]) & \"\""
                    Format =">"

                    LayoutCachedLeft =8376
                    LayoutCachedWidth =15336
                    LayoutCachedHeight =540
                End
                Begin Rectangle
                    Top =255
                    Width =15336
                    Height =285
                    BackColor =15527148
                    BorderColor =15527148
                    Name ="Box107"
                    LayoutCachedTop =255
                    LayoutCachedWidth =15336
                    LayoutCachedHeight =540
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =4668
                    Top =276
                    Width =1320
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="Label7"
                    Caption ="Engine"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =4668
                    LayoutCachedTop =276
                    LayoutCachedWidth =5988
                    LayoutCachedHeight =504
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =120
                    Top =276
                    Width =615
                    Height =228
                    FontSize =9
                    FontWeight =700
                    Name ="Label88"
                    Caption ="Asset"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =120
                    LayoutCachedTop =276
                    LayoutCachedWidth =735
                    LayoutCachedHeight =504
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =792
                    Top =276
                    Width =1836
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="Label89"
                    Caption ="VIN"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =792
                    LayoutCachedTop =276
                    LayoutCachedWidth =2628
                    LayoutCachedHeight =504
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =2688
                    Top =276
                    Width =1920
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="Label90"
                    Caption ="Description"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =2688
                    LayoutCachedTop =276
                    LayoutCachedWidth =4608
                    LayoutCachedHeight =504
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    Left =6444
                    Top =276
                    Width =1008
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="Label91"
                    Caption ="Parked"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =6444
                    LayoutCachedTop =276
                    LayoutCachedWidth =7452
                    LayoutCachedHeight =504
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    Left =8424
                    Top =276
                    Width =432
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="Label92"
                    Caption ="HRS"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =8424
                    LayoutCachedTop =276
                    LayoutCachedWidth =8856
                    LayoutCachedHeight =504
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    Left =8916
                    Top =276
                    Width =1008
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="Label93"
                    Caption ="Buy Out "
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =8916
                    LayoutCachedTop =276
                    LayoutCachedWidth =9924
                    LayoutCachedHeight =504
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    Left =9996
                    Top =276
                    Width =552
                    Height =228
                    FontSize =10
                    FontWeight =700
                    BorderColor =14211288
                    Name ="Label94"
                    Caption ="Days"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =9996
                    LayoutCachedTop =276
                    LayoutCachedWidth =10548
                    LayoutCachedHeight =504
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =10620
                    Top =276
                    Width =996
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="Label95"
                    Caption ="Target $"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =10620
                    LayoutCachedTop =276
                    LayoutCachedWidth =11616
                    LayoutCachedHeight =504
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =11676
                    Top =276
                    Width =396
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="Label96"
                    Caption ="UPD"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =11676
                    LayoutCachedTop =276
                    LayoutCachedWidth =12072
                    LayoutCachedHeight =504
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =12132
                    Top =276
                    Width =492
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="Label97"
                    Caption ="Title"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =12132
                    LayoutCachedTop =276
                    LayoutCachedWidth =12624
                    LayoutCachedHeight =504
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =12684
                    Top =276
                    Width =852
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="Label98"
                    Caption ="Buyer"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =12684
                    LayoutCachedTop =276
                    LayoutCachedWidth =13536
                    LayoutCachedHeight =504
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =13596
                    Top =276
                    Width =708
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="Label99"
                    Caption ="Deposit"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =13596
                    LayoutCachedTop =276
                    LayoutCachedWidth =14304
                    LayoutCachedHeight =504
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    Left =7512
                    Top =276
                    Width =852
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="Label102"
                    Caption ="Miles"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =7512
                    LayoutCachedTop =276
                    LayoutCachedWidth =8364
                    LayoutCachedHeight =504
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    Left =14364
                    Top =276
                    Width =960
                    Height =228
                    FontSize =10
                    FontWeight =700
                    Name ="Label103"
                    Caption ="Invoiced"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =14364
                    LayoutCachedTop =276
                    LayoutCachedWidth =15324
                    LayoutCachedHeight =504
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =6060
                    Top =276
                    Width =312
                    Height =228
                    FontSize =9
                    FontWeight =700
                    BackColor =15527148
                    BorderColor =14211288
                    Name ="Label106"
                    Caption ="STS"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =6060
                    LayoutCachedTop =276
                    LayoutCachedWidth =6372
                    LayoutCachedHeight =504
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =285
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9996
                    Top =36
                    Width =552
                    Height =204
                    FontSize =9
                    TabIndex =8
                    BorderColor =14211288
                    Name ="Text23"
                    ControlSource ="=IIf(IsDate([BuyOutDate]),DateDiff(\"d\",[BuyOutDate],Date())+1,Null)"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =9996
                    LayoutCachedTop =36
                    LayoutCachedWidth =10548
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =792
                    Top =36
                    Width =1836
                    Height =204
                    ColumnWidth =2085
                    FontSize =9
                    TabIndex =1
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =792
                    LayoutCachedTop =36
                    LayoutCachedWidth =2628
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4668
                    Top =36
                    Width =1320
                    Height =204
                    FontSize =9
                    TabIndex =4
                    Name ="Engine"
                    ControlSource ="ENGINETYPE"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =4668
                    LayoutCachedTop =36
                    LayoutCachedWidth =5988
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12132
                    Top =36
                    Width =492
                    Height =204
                    FontSize =9
                    TabIndex =11
                    Name ="Text34"
                    ControlSource ="=UCase([Action])"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =12132
                    LayoutCachedTop =36
                    LayoutCachedWidth =12624
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CheckBox
                    OldBorderStyle =0
                    Left =11676
                    Top =36
                    Width =396
                    Height =204
                    TabIndex =10
                    Name ="Check41"
                    ControlSource ="unpdloan"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =11676
                    LayoutCachedTop =36
                    LayoutCachedWidth =12072
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12684
                    Top =36
                    Width =852
                    Height =204
                    FontSize =9
                    TabIndex =12
                    Name ="Clients.ClientShNm"
                    ControlSource ="Buyer"
                    StatusBarText ="Short Name for Reporting"
                    EventProcPrefix ="Clients_ClientShNm"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =12684
                    LayoutCachedTop =36
                    LayoutCachedWidth =13536
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10620
                    Top =36
                    Width =996
                    Height =204
                    FontSize =9
                    TabIndex =9
                    Name ="TargetSale"
                    ControlSource ="TargetSale"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="price Rmk Sm should get for unit"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =10620
                    LayoutCachedTop =36
                    LayoutCachedWidth =11616
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8916
                    Top =36
                    Width =1008
                    Height =204
                    FontSize =9
                    TabIndex =7
                    Name ="BuyOutDate"
                    ControlSource ="BuyOutDate"
                    Format ="Short Date"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =1

                    LayoutCachedLeft =8916
                    LayoutCachedTop =36
                    LayoutCachedWidth =9924
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13596
                    Top =36
                    Width =708
                    Height =204
                    FontSize =9
                    TabIndex =13
                    Name ="RmktDeposit"
                    ControlSource ="RmktDeposit"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =13596
                    LayoutCachedTop =36
                    LayoutCachedWidth =14304
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6444
                    Top =36
                    Width =1008
                    Height =204
                    FontSize =9
                    TabIndex =5
                    Name ="ParkedX"
                    ControlSource ="ParkedX"
                    ControlTipText ="The 'Park Date' progression is Parked Date, BuyOut Date Off Lease Date, Projecte"
                        "d Park Date (future)."
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =6444
                    LayoutCachedTop =36
                    LayoutCachedWidth =7452
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8424
                    Top =36
                    Width =432
                    Height =204
                    FontSize =9
                    TabIndex =6
                    Name ="Text73"
                    ControlSource ="=nz([HrParked],\"--\")"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =8424
                    LayoutCachedTop =36
                    LayoutCachedWidth =8856
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2688
                    Top =36
                    Width =1920
                    Height =204
                    FontSize =9
                    TabIndex =2
                    Name ="Text75"
                    ControlSource ="Desc"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =2688
                    LayoutCachedTop =36
                    LayoutCachedWidth =4608
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =36
                    Width =615
                    Height =204
                    FontSize =9
                    Name ="Text84"
                    ControlSource ="UnitID"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =120
                    LayoutCachedTop =36
                    LayoutCachedWidth =735
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =163
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14364
                    Top =36
                    Width =960
                    Height =204
                    FontSize =9
                    TabIndex =15
                    Name ="InvoiceSentDate"
                    ControlSource ="InvoiceSentDate"
                    Format ="Short Date"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =1

                    LayoutCachedLeft =14364
                    LayoutCachedTop =36
                    LayoutCachedWidth =15324
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7512
                    Top =36
                    Width =852
                    Height =204
                    FontSize =9
                    TabIndex =14
                    Name ="MILEAGEX"
                    ControlSource ="=CDbl(nz([MILEAGE],0))"
                    Format ="Standard"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =7512
                    LayoutCachedTop =36
                    LayoutCachedWidth =8364
                    LayoutCachedHeight =240
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
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6060
                    Top =36
                    Width =312
                    Height =204
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    BorderColor =14211288
                    Name ="Text105"
                    ControlSource ="unitstatus"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =6060
                    LayoutCachedTop =36
                    LayoutCachedWidth =6372
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =240
            BreakLevel =2
            Name ="GroupFooter4"
            Begin
                Begin Rectangle
                    Top =90
                    Width =15336
                    Height =45
                    BackColor =12566463
                    BorderColor =12566463
                    Name ="Box123"
                    LayoutCachedTop =90
                    LayoutCachedWidth =15336
                    LayoutCachedHeight =135
                End
            End
        End
        Begin PageFooter
            Height =315
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13140
                    Width =2100
                    Height =315
                    FontSize =9
                    Name ="Text30"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =13140
                    LayoutCachedWidth =15240
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =3180
                    Height =315
                    FontSize =9
                    TabIndex =1
                    Name ="Text31"
                    ControlSource ="=\"printed on \" & Date() & \" @ \" & Time()"

                    LayoutCachedWidth =3180
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =2
                            TextFontFamily =34
                            Left =3180
                            Width =9840
                            Height =285
                            FontSize =9
                            FontWeight =700
                            Name ="Label63"
                            Caption ="*The 'Park Date' progression is Parked Date, BuyOut Date Off Lease Date, Project"
                                "ed Park Date (future)."
                            LayoutCachedLeft =3180
                            LayoutCachedWidth =13020
                            LayoutCachedHeight =285
                        End
                    End
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =285
            Name ="ReportFooter"
            AutoHeight =1
            Begin
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =14610
                    Width =600
                    Height =270
                    FontSize =10
                    ForeColor =-2147483607
                    Name ="Text50"
                    ControlSource ="=Count([unitvin])"

                    LayoutCachedLeft =14610
                    LayoutCachedWidth =15210
                    LayoutCachedHeight =270
                    BackThemeColorIndex =2
                    Begin
                        Begin Label
                            TextFontCharSet =186
                            TextAlign =3
                            TextFontFamily =34
                            Left =13440
                            Top =15
                            Width =1140
                            Height =270
                            FontSize =10
                            FontWeight =700
                            Name ="Label51"
                            Caption ="Total Units:"
                            LayoutCachedLeft =13440
                            LayoutCachedTop =15
                            LayoutCachedWidth =14580
                            LayoutCachedHeight =285
                            ForeThemeColorIndex =2
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "RMKTInventory.cls"
