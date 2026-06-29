Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =11
    ItemSuffix =215
    Left =25185
    Top =4215
    Right =-29326
    Bottom =11520
    RecSrcDt = Begin
        0x322706b49305e540
    End
    Caption ="Add Address"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    PictureSizeMode =1
    DatasheetGridlinesColor12 =0
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
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
        Begin ComboBox
            AddColon = NotDefault
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
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin ToggleButton
            TextFontCharSet =177
            TextFontFamily =0
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
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
            PressedForeThemeColorIndex =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =315
            BackColor =14211288
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =45
                    Width =5970
                    Height =315
                    ColumnOrder =1
                    BackColor =14211288
                    ForeColor =0
                    Name ="Header"
                    ControlSource ="=\"Adding an address for Clientgroup:<b> \" & DLookUp(\"ClientGroupName\",\"Clie"
                        "ntGroups\",\"ClientGroupID=\" & [ClientGroupID] & \"\") & \"</b>\""
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedLeft =45
                    LayoutCachedWidth =6015
                    LayoutCachedHeight =315
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4185
                    Width =1155
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =1
                    ForeColor =0
                    Name ="MoveTo"
                    FontName ="Segoe UI"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4185
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =240
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =3210
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1275
                    Top =2385
                    Width =3030
                    Height =330
                    TabIndex =5
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LZip"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =1275
                    LayoutCachedTop =2385
                    LayoutCachedWidth =4305
                    LayoutCachedHeight =2715
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1275
                    Top =435
                    Width =5190
                    Height =330
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LAddrLine1"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =1275
                    LayoutCachedTop =435
                    LayoutCachedWidth =6465
                    LayoutCachedHeight =765
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =300
                            Top =435
                            Width =912
                            Height =330
                            FontSize =9
                            BorderColor =0
                            Name ="Label60"
                            Caption ="Address 1"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =300
                            LayoutCachedTop =435
                            LayoutCachedWidth =1212
                            LayoutCachedHeight =765
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1275
                    Top =825
                    Width =5190
                    Height =330
                    TabIndex =1
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LAddrLine2"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =1275
                    LayoutCachedTop =825
                    LayoutCachedWidth =6465
                    LayoutCachedHeight =1155
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =300
                            Top =825
                            Width =912
                            Height =330
                            FontSize =9
                            BorderColor =0
                            Name ="Label580"
                            Caption ="Address 2"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =300
                            LayoutCachedTop =825
                            LayoutCachedWidth =1212
                            LayoutCachedHeight =1155
                            RowStart =1
                            RowEnd =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =7200
                    Left =1275
                    Top =1605
                    Width =3030
                    Height =330
                    TabIndex =3
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LCity"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT City, CityID, County, CountyID From vw_CCSZ_Lookup WHERE StateID =3 GROUP"
                        " BY City, CityID, County, CountyID ORDER BY City, County;"
                    ColumnWidths ="3600;0;2880;0"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1275
                    LayoutCachedTop =1605
                    LayoutCachedWidth =4305
                    LayoutCachedHeight =1935
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =300
                            Top =1605
                            Width =912
                            Height =330
                            FontSize =9
                            BorderColor =0
                            Name ="Label42"
                            Caption ="City"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =300
                            LayoutCachedTop =1605
                            LayoutCachedWidth =1212
                            LayoutCachedHeight =1935
                            RowStart =3
                            RowEnd =3
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =3600
                    Left =1275
                    Top =1995
                    Width =3030
                    Height =330
                    TabIndex =4
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LCounty"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT County, CountyID From vw_CCSZ_Lookup WHERE StateID =3 GROUP BY County, Co"
                        "untyID ORDER BY County;"
                    ColumnWidths ="3600;0"
                    StatusBarText ="link to counties table"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1275
                    LayoutCachedTop =1995
                    LayoutCachedWidth =4305
                    LayoutCachedHeight =2325
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =300
                            Top =1995
                            Width =912
                            Height =330
                            FontSize =9
                            BorderColor =0
                            Name ="Label43"
                            Caption ="County"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =300
                            LayoutCachedTop =1995
                            LayoutCachedWidth =1212
                            LayoutCachedHeight =2325
                            RowStart =4
                            RowEnd =4
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =20
                    ListWidth =2880
                    Left =1275
                    Top =1215
                    Width =3030
                    Height =330
                    TabIndex =2
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TaxStates.PostalAbbreviation, IIf([RentReceiptAvailable]=True,\"RR\",\"--"
                        "\") & \"/\" & IIf([UpfrontAvailable]=True,\"UpFront\",\"--\") AS [Option], TaxSt"
                        "ates.ISOCountryCode2, TaxStates.ID FROM TaxStates GROUP BY TaxStates.PostalAbbre"
                        "viation, IIf([RentReceiptAvailable]=True,\"RR\",\"--\") & \"/\" & IIf([UpfrontAv"
                        "ailable]=True,\"UpFront\",\"--\"), TaxStates.ISOCountryCode2, TaxStates.ID, TaxS"
                        "tates.PostalAbbreviation HAVING (((TaxStates.ISOCountryCode2)=\"US\" Or (TaxStat"
                        "es.ISOCountryCode2)=\"CA\" Or (TaxStates.ISOCountryCode2)=\"MX\")) ORDER BY TaxS"
                        "tates.ISOCountryCode2 DESC , TaxStates.PostalAbbreviation;"
                    ColumnWidths ="864;0;432;0"
                    StatusBarText ="lookup in lookup_state table"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1275
                    LayoutCachedTop =1215
                    LayoutCachedWidth =4305
                    LayoutCachedHeight =1545
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =300
                    Top =2385
                    Width =912
                    Height =330
                    FontSize =9
                    BorderColor =0
                    Name ="Label778"
                    Caption ="Zip"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =300
                    LayoutCachedTop =2385
                    LayoutCachedWidth =1212
                    LayoutCachedHeight =2715
                    RowStart =5
                    RowEnd =5
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =300
                    Top =1215
                    Width =912
                    Height =330
                    FontSize =9
                    BorderColor =0
                    Name ="Label8"
                    Caption ="State"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =300
                    LayoutCachedTop =1215
                    LayoutCachedWidth =1212
                    LayoutCachedHeight =1545
                    RowStart =2
                    RowEnd =2
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1155
                    Top =2850
                    Width =660
                    FontSize =8
                    TabIndex =8
                    BorderColor =12632256
                    Name ="ClientID"
                    StatusBarText ="link to Clients table"
                    FontName ="Tahoma"
                    GridlineColor =0

                    LayoutCachedLeft =1155
                    LayoutCachedTop =2850
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =3090
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1815
                    Top =2850
                    Width =720
                    FontSize =8
                    TabIndex =9
                    BorderColor =12632256
                    Name ="LAddressID"
                    FontName ="Tahoma"
                    GridlineColor =0

                    LayoutCachedLeft =1815
                    LayoutCachedTop =2850
                    LayoutCachedWidth =2535
                    LayoutCachedHeight =3090
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2535
                    Top =2850
                    Width =720
                    FontSize =8
                    TabIndex =10
                    BorderColor =12632256
                    Name ="LStateID"
                    FontName ="Tahoma"
                    GridlineColor =0

                    LayoutCachedLeft =2535
                    LayoutCachedTop =2850
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =3090
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3255
                    Top =2850
                    Width =720
                    FontSize =8
                    TabIndex =11
                    BorderColor =12632256
                    Name ="LCountyID"
                    FontName ="Tahoma"
                    GridlineColor =0

                    LayoutCachedLeft =3255
                    LayoutCachedTop =2850
                    LayoutCachedWidth =3975
                    LayoutCachedHeight =3090
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3975
                    Top =2850
                    Width =720
                    FontSize =8
                    TabIndex =12
                    BorderColor =12632256
                    Name ="LCityID"
                    FontName ="Tahoma"
                    GridlineColor =0

                    LayoutCachedLeft =3975
                    LayoutCachedTop =2850
                    LayoutCachedWidth =4695
                    LayoutCachedHeight =3090
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Image
                    PictureTiling = NotDefault
                    BackStyle =1
                    BorderWidth =1
                    SizeMode =0
                    PictureType =2
                    Width =7200
                    Height =180
                    BackColor =15527148
                    BorderColor =14211288
                    Name ="Image142"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedWidth =7200
                    LayoutCachedHeight =180
                    TabIndex =13
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =450
                    Top =2850
                    Width =660
                    FontSize =8
                    TabIndex =7
                    BorderColor =12632256
                    Name ="ClientGroupID"
                    StatusBarText ="link to Clients table"
                    FontName ="Tahoma"
                    GridlineColor =0

                    LayoutCachedLeft =450
                    LayoutCachedTop =2850
                    LayoutCachedWidth =1110
                    LayoutCachedHeight =3090
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4905
                    Top =1785
                    Width =1110
                    Height =522
                    FontSize =8
                    TabIndex =6
                    ForeColor =8210719
                    Name ="SaveMe"
                    Caption ="Save Address"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000ab598980ab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989800000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffd9b4caffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffd9b4caffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4905
                    LayoutCachedTop =1785
                    LayoutCachedWidth =6015
                    LayoutCachedHeight =2307
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =8210719
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =8210719
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "LocationAddressesAdd.cls"
