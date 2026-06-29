Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11700
    DatasheetFontHeight =11
    ItemSuffix =52
    Left =9795
    Top =5325
    Right =25125
    Bottom =18390
    RecSrcDt = Begin
        0x2934e89c1d07e540
    End
    RecordSource ="SELECT LocationID, LocationType, ClientREF, LocationName, Locationsn, IIf(Len([L"
        "AddrLine1])>4,[LAddrLine1] & ' ' & [LCity] & '  ' & [LState] & ' ',Null) & IIf(L"
        "en([LZip])>5,Left([LZip],5) & '-' & Right([LZip],4),[LZip]) AS Addr, Locations.L"
        "AddrLine1, Locations.LAddrLine2, Locations.LAddrLine3, LCity, LState, LZip, Long"
        "itude, Latitude FROM Locations;"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
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
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CustomControl
            OldBorderStyle =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin WebBrowser
            OldBorderStyle =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =13080
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin WebBrowser
                    OverlapFlags =85
                    Left =120
                    Top =240
                    Width =8520
                    Height =6960
                    AutoActivate =1
                    Name ="WebBrowser0"
                    OleData = Begin
                        0x000e0000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000020000000000000000100000 ,
                        0x0400000001000000feffffff0000000003000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff0100000061f956880a34d011a96b00c0 ,
                        0x4fd705a2000000000000000000000000a0fecb292b9fd1010700000000010000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000001000000ac000000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000038000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fefffffffdfffffffffffffffffffffffffffffffffffffffeffffff ,
                        0xfeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff0100000061f956880a34d011a96b00c0 ,
                        0x4fd705a200000000000000000000000080bdaf1bb2b2d1010500000000010000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000001000000ac000000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000038000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fffffffffffffffffefffffffdfffffffefffffffeffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffeffffff0200000003000000feffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff38000000000000000100000000000000000000000000000000000000 ,
                        0x0000000038000000000000000000000000000000000000000000000000000000 ,
                        0x0000000061f956880a34d011a96b00c04fd705a24c0000009a3a0000da2f0000 ,
                        0x0000000000000000000000000000000000000000000000004c00000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4c0000000114020000000000c000000000000046800000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000001000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =120
                    LayoutCachedTop =240
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =7200
                    ControlSource ="=\"https://fleetadvantage.sharepoint.com/Customers/Forms/AllItems.aspx\""
                    HyperlinkBinderDescription ="1|=https://fleetadvantage.sharepoint.com/Customers/Forms/AllItems.aspx"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1452
                    Top =7320
                    Width =6072
                    Height =315
                    TabIndex =3
                    Name ="AddressToSearch"
                    ControlSource ="Addr"
                    GroupTable =1

                    LayoutCachedLeft =1452
                    LayoutCachedTop =7320
                    LayoutCachedWidth =7524
                    LayoutCachedHeight =7635
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =7320
                            Width =1020
                            Height =315
                            Name ="Label8"
                            Caption ="Address to Search"
                            GroupTable =1
                            LayoutCachedLeft =360
                            LayoutCachedTop =7320
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =7635
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8880
                    Top =5580
                    TabIndex =1
                    Name ="SetAddr"
                    Caption ="Set Addr"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8880
                    LayoutCachedTop =5580
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =5940
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =5
                    Gradient =0
                    BackColor =12419407
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    HoverColor =13277810
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    QuickStyle =9
                    QuickStyleMask =-129
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1452
                    Top =7704
                    Width =2460
                    Height =312
                    TabIndex =4
                    Name ="Latitude"
                    ControlSource ="Latitude"
                    Format ="General Number"
                    GroupTable =1

                    LayoutCachedLeft =1452
                    LayoutCachedTop =7704
                    LayoutCachedWidth =3912
                    LayoutCachedHeight =8016
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =7704
                            Width =1020
                            Height =312
                            Name ="Label11"
                            Caption ="Latitude"
                            GroupTable =1
                            LayoutCachedLeft =360
                            LayoutCachedTop =7704
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =8016
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5064
                    Top =7704
                    Width =2460
                    Height =312
                    TabIndex =5
                    Name ="Longitude"
                    ControlSource ="Longitude"
                    Format ="General Number"
                    GroupTable =1

                    LayoutCachedLeft =5064
                    LayoutCachedTop =7704
                    LayoutCachedWidth =7524
                    LayoutCachedHeight =8016
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =3972
                            Top =7704
                            Width =1020
                            Height =312
                            Name ="Label13"
                            Caption ="Longitude"
                            GroupTable =1
                            LayoutCachedLeft =3972
                            LayoutCachedTop =7704
                            LayoutCachedWidth =4992
                            LayoutCachedHeight =8016
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8880
                    Top =6060
                    TabIndex =2
                    Name ="GetLatLog"
                    Caption ="Get Lat-Log"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8880
                    LayoutCachedTop =6060
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =6420
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =5
                    Gradient =0
                    BackColor =12419407
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    HoverColor =13277810
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    QuickStyle =9
                    QuickStyleMask =-129
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7680
                    Top =7740
                    Width =720
                    Height =315
                    TabIndex =6
                    Name ="ID"
                    ControlSource ="=LocationID"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =7740
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =8055
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7680
                            Top =7320
                            Width =690
                            Height =315
                            Name ="Label40"
                            Caption ="ID"
                            LayoutCachedLeft =7680
                            LayoutCachedTop =7320
                            LayoutCachedWidth =8370
                            LayoutCachedHeight =7635
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8880
                    Top =6540
                    TabIndex =7
                    Name ="MoveNext"
                    Caption ="MoveNext"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8880
                    LayoutCachedTop =6540
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =6900
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =5
                    Gradient =0
                    BackColor =12419407
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    HoverColor =13277810
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    QuickStyle =9
                    QuickStyleMask =-129
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =10
                    Left =120
                    Top =8460
                    Width =11400
                    Height =4260
                    FontSize =9
                    TabIndex =8
                    Name ="LocationsList"
                    RowSourceType ="Table/Query"
                    RowSource ="Select * From GeoCodeLocations Order By LocationSN"
                    ColumnWidths ="0;2448;2592;2160;720;1728;864;720;720"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =120
                    LayoutCachedTop =8460
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =12720
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8880
                    Top =7260
                    Width =1980
                    Height =300
                    TabIndex =9
                    Name ="SortBlank"
                    Caption ="Sort Blank"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8880
                    LayoutCachedTop =7260
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =7560
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =12419407
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    HoverColor =13277810
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    QuickStyle =9
                    QuickStyleMask =-129
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8880
                    Top =7680
                    Width =1980
                    Height =300
                    TabIndex =10
                    Name ="BlankWithUnits"
                    Caption ="Blank With Units"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8880
                    LayoutCachedTop =7680
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =7980
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =12419407
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    HoverColor =13277810
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    QuickStyle =9
                    QuickStyleMask =-129
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =120
                    Top =8220
                    Width =1152
                    FontSize =9
                    TabIndex =11
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchLocation"
                    OnChange ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =120
                    LayoutCachedTop =8220
                    LayoutCachedWidth =1272
                    LayoutCachedHeight =8460
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9000
                    Top =420
                    Width =2340
                    Height =315
                    TabIndex =12
                    Name ="VinToDecode"

                    LayoutCachedLeft =9000
                    LayoutCachedTop =420
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =735
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8880
                            Width =690
                            Height =315
                            Name ="Label49"
                            Caption ="Text48"
                            LayoutCachedLeft =8880
                            LayoutCachedWidth =9570
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8940
                    Top =780
                    TabIndex =13
                    Name ="DecodeVIN"
                    Caption ="Decode VIN"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8940
                    LayoutCachedTop =780
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =1140
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =5
                    Gradient =0
                    BackColor =12419407
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    HoverColor =13277810
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    QuickStyle =9
                    QuickStyleMask =-129
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8940
                    Top =1260
                    TabIndex =14
                    Name ="DecodeVIN2"
                    Caption ="Decode VIN2"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8940
                    LayoutCachedTop =1260
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =1620
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =5
                    Gradient =0
                    BackColor =12419407
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    HoverColor =13277810
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    QuickStyle =9
                    QuickStyleMask =-129
                End
            End
        End
    End
End
CodeBehindForm
' See "VIN_Decoder.cls"
