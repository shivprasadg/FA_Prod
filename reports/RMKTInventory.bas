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
    Width =15336
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
            Height =540
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
                    Left =4665
                    Top =270
                    Width =1320
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="Label7"
                    Caption ="Engine"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =4665
                    LayoutCachedTop =270
                    LayoutCachedWidth =5985
                    LayoutCachedHeight =495
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
                    Top =270
                    Width =615
                    Height =225
                    FontSize =9
                    FontWeight =700
                    Name ="Label88"
                    Caption ="Asset"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =120
                    LayoutCachedTop =270
                    LayoutCachedWidth =735
                    LayoutCachedHeight =495
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =795
                    Top =270
                    Width =1830
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="Label89"
                    Caption ="VIN"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =795
                    LayoutCachedTop =270
                    LayoutCachedWidth =2625
                    LayoutCachedHeight =495
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =2685
                    Top =270
                    Width =1923
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="Label90"
                    Caption ="Description"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =2685
                    LayoutCachedTop =270
                    LayoutCachedWidth =4608
                    LayoutCachedHeight =495
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    Left =6435
                    Top =270
                    Width =1005
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="Label91"
                    Caption ="Parked"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =6435
                    LayoutCachedTop =270
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =495
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    Left =8415
                    Top =270
                    Width =435
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="Label92"
                    Caption ="HRS"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =8415
                    LayoutCachedTop =270
                    LayoutCachedWidth =8850
                    LayoutCachedHeight =495
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    Left =8910
                    Top =270
                    Width =1005
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="Label93"
                    Caption ="Buy Out "
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =8910
                    LayoutCachedTop =270
                    LayoutCachedWidth =9915
                    LayoutCachedHeight =495
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
                    Left =9975
                    Top =270
                    Width =555
                    Height =225
                    FontSize =10
                    FontWeight =700
                    BorderColor =14211288
                    Name ="Label94"
                    Caption ="Days"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =9975
                    LayoutCachedTop =270
                    LayoutCachedWidth =10530
                    LayoutCachedHeight =495
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =10605
                    Top =270
                    Width =990
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="Label95"
                    Caption ="Target $"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =10605
                    LayoutCachedTop =270
                    LayoutCachedWidth =11595
                    LayoutCachedHeight =495
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =11655
                    Top =270
                    Width =390
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="Label96"
                    Caption ="UPD"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =11655
                    LayoutCachedTop =270
                    LayoutCachedWidth =12045
                    LayoutCachedHeight =495
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =12105
                    Top =270
                    Width =495
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="Label97"
                    Caption ="Title"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =12105
                    LayoutCachedTop =270
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =495
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =12660
                    Top =270
                    Width =855
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="Label98"
                    Caption ="Buyer"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =12660
                    LayoutCachedTop =270
                    LayoutCachedWidth =13515
                    LayoutCachedHeight =495
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    Left =13575
                    Top =270
                    Width =705
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="Label99"
                    Caption ="Deposit"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =13575
                    LayoutCachedTop =270
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =495
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    Left =7500
                    Top =270
                    Width =855
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="Label102"
                    Caption ="Miles"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =7500
                    LayoutCachedTop =270
                    LayoutCachedWidth =8355
                    LayoutCachedHeight =495
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontCharSet =163
                    TextAlign =2
                    TextFontFamily =34
                    Left =14340
                    Top =270
                    Width =960
                    Height =225
                    FontSize =10
                    FontWeight =700
                    Name ="Label103"
                    Caption ="Invoiced"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =14340
                    LayoutCachedTop =270
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =495
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
                    Left =6045
                    Top =270
                    Width =315
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =15527148
                    BorderColor =14211288
                    Name ="Label106"
                    Caption ="STS"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =6045
                    LayoutCachedTop =270
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =495
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
                    Left =9975
                    Top =30
                    Width =555
                    Height =209
                    FontSize =9
                    TabIndex =8
                    BorderColor =14211288
                    Name ="Text23"
                    ControlSource ="=IIf(IsDate([BuyOutDate]),DateDiff(\"d\",[BuyOutDate],Date())+1,Null)"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =9975
                    LayoutCachedTop =30
                    LayoutCachedWidth =10530
                    LayoutCachedHeight =239
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
                    Left =795
                    Top =30
                    Width =1830
                    Height =209
                    ColumnWidth =2085
                    FontSize =9
                    TabIndex =1
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =795
                    LayoutCachedTop =30
                    LayoutCachedWidth =2625
                    LayoutCachedHeight =239
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
                    Left =4665
                    Top =30
                    Width =1320
                    Height =209
                    FontSize =9
                    TabIndex =4
                    Name ="Engine"
                    ControlSource ="ENGINETYPE"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =4665
                    LayoutCachedTop =30
                    LayoutCachedWidth =5985
                    LayoutCachedHeight =239
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
                    Left =12105
                    Top =30
                    Width =495
                    Height =209
                    FontSize =9
                    TabIndex =11
                    Name ="Text34"
                    ControlSource ="=UCase([Action])"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =12105
                    LayoutCachedTop =30
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =239
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CheckBox
                    OldBorderStyle =0
                    Left =11655
                    Top =30
                    Width =390
                    Height =209
                    TabIndex =10
                    Name ="Check41"
                    ControlSource ="unpdloan"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =11655
                    LayoutCachedTop =30
                    LayoutCachedWidth =12045
                    LayoutCachedHeight =239
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
                    Left =12660
                    Top =30
                    Width =855
                    Height =209
                    FontSize =9
                    TabIndex =12
                    Name ="Clients.ClientShNm"
                    ControlSource ="Buyer"
                    StatusBarText ="Short Name for Reporting"
                    EventProcPrefix ="Clients_ClientShNm"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =12660
                    LayoutCachedTop =30
                    LayoutCachedWidth =13515
                    LayoutCachedHeight =239
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
                    Left =10605
                    Top =30
                    Width =990
                    Height =209
                    FontSize =9
                    TabIndex =9
                    Name ="TargetSale"
                    ControlSource ="TargetSale"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="price Rmk Sm should get for unit"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =10605
                    LayoutCachedTop =30
                    LayoutCachedWidth =11595
                    LayoutCachedHeight =239
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
                    Left =8910
                    Top =30
                    Width =1005
                    Height =209
                    FontSize =9
                    TabIndex =7
                    Name ="BuyOutDate"
                    ControlSource ="BuyOutDate"
                    Format ="Short Date"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =1

                    LayoutCachedLeft =8910
                    LayoutCachedTop =30
                    LayoutCachedWidth =9915
                    LayoutCachedHeight =239
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
                    Left =13575
                    Top =30
                    Width =705
                    Height =209
                    FontSize =9
                    TabIndex =13
                    Name ="RmktDeposit"
                    ControlSource ="RmktDeposit"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =13575
                    LayoutCachedTop =30
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =239
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
                    Left =6435
                    Top =30
                    Width =1005
                    Height =209
                    FontSize =9
                    TabIndex =5
                    Name ="ParkedX"
                    ControlSource ="ParkedX"
                    ControlTipText ="The 'Park Date' progression is Parked Date, BuyOut Date Off Lease Date, Projecte"
                        "d Park Date (future)."
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =6435
                    LayoutCachedTop =30
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =239
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
                    Left =8415
                    Top =30
                    Width =435
                    Height =209
                    FontSize =9
                    TabIndex =6
                    Name ="Text73"
                    ControlSource ="=nz([HrParked],\"--\")"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =8415
                    LayoutCachedTop =30
                    LayoutCachedWidth =8850
                    LayoutCachedHeight =239
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
                    Left =2685
                    Top =30
                    Width =1923
                    Height =209
                    FontSize =9
                    TabIndex =2
                    Name ="Text75"
                    ControlSource ="Desc"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =2685
                    LayoutCachedTop =30
                    LayoutCachedWidth =4608
                    LayoutCachedHeight =239
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
                    Top =30
                    Width =615
                    Height =209
                    FontSize =9
                    Name ="Text84"
                    ControlSource ="UnitID"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =120
                    LayoutCachedTop =30
                    LayoutCachedWidth =735
                    LayoutCachedHeight =239
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
                    Left =14340
                    Top =30
                    Width =960
                    Height =209
                    FontSize =9
                    TabIndex =15
                    Name ="InvoiceSentDate"
                    ControlSource ="InvoiceSentDate"
                    Format ="Short Date"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =1

                    LayoutCachedLeft =14340
                    LayoutCachedTop =30
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =239
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
                    Left =7500
                    Top =30
                    Width =855
                    Height =209
                    FontSize =9
                    TabIndex =14
                    Name ="MILEAGEX"
                    ControlSource ="=CDbl(nz([MILEAGE],0))"
                    Format ="Standard"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =7500
                    LayoutCachedTop =30
                    LayoutCachedWidth =8355
                    LayoutCachedHeight =239
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
                    Left =6045
                    Top =30
                    Width =315
                    Height =209
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    BorderColor =14211288
                    Name ="Text105"
                    ControlSource ="unitstatus"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =6045
                    LayoutCachedTop =30
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =239
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
