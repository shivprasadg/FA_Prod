Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18732
    DatasheetFontHeight =11
    ItemSuffix =213
    Left =8535
    Top =8910
    Right =29055
    Bottom =20580
    Filter ="([Locations].[ClientGroupID]=703)"
    RecSrcDt = Begin
        0x6a85249d1029e640
    End
    RecordSource ="Locations"
    Caption ="Location Address Editor"
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
                    ColumnOrder =2
                    BackColor =14211288
                    ForeColor =0
                    Name ="Header"
                    ControlSource ="=\"Edit address for Clientgroup:<b> \" & DLookUp(\"ClientGroupName\",\"ClientGro"
                        "ups\",\"ClientGroupID=\" & [ClientGroupID] & \"\") & \"</b>\""
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
                    ColumnOrder =1
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
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9432
                    Top =58
                    Width =389
                    Height =225
                    ColumnOrder =0
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    BackColor =14211288
                    BorderColor =1643706
                    ForeColor =0
                    Name ="AddressCount"
                    GridlineColor =0

                    LayoutCachedLeft =9432
                    LayoutCachedTop =58
                    LayoutCachedWidth =9821
                    LayoutCachedHeight =283
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =7590
                            Top =58
                            Width =1845
                            Height =225
                            FontSize =8
                            BackColor =14211288
                            BorderColor =1643706
                            ForeColor =5855577
                            Name ="Label63"
                            Caption ="Addresses In Client Group"
                            GridlineColor =0
                            LayoutCachedLeft =7590
                            LayoutCachedTop =58
                            LayoutCachedWidth =9435
                            LayoutCachedHeight =283
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =8784
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =93
                    Left =180
                    Top =2314
                    Width =810
                    Height =1706
                    BorderColor =8872517
                    Name ="Box297"
                    LayoutCachedLeft =180
                    LayoutCachedTop =2314
                    LayoutCachedWidth =990
                    LayoutCachedHeight =4020
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =223
                    Left =180
                    Top =2674
                    Width =810
                    Height =1346
                    BorderColor =8872517
                    Name ="Box147"
                    LayoutCachedLeft =180
                    LayoutCachedTop =2674
                    LayoutCachedWidth =990
                    LayoutCachedHeight =4020
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =223
                    Left =180
                    Top =3030
                    Width =810
                    Height =990
                    BorderColor =8872517
                    Name ="Box148"
                    LayoutCachedLeft =180
                    LayoutCachedTop =3030
                    LayoutCachedWidth =990
                    LayoutCachedHeight =4020
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    BorderWidth =1
                    OverlapFlags =223
                    Left =885
                    Top =2025
                    Width =225
                    Height =2006
                    BorderColor =8872517
                    Name ="Box149"
                    LayoutCachedLeft =885
                    LayoutCachedTop =2025
                    LayoutCachedWidth =1110
                    LayoutCachedHeight =4031
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Left =5976
                    Top =144
                    Width =174
                    Height =8640
                    BorderColor =16777215
                    Name ="Image2243"
                    Picture ="LtGrey1_VertShadow"
                    VerticalAnchor =2

                    LayoutCachedLeft =5976
                    LayoutCachedTop =144
                    LayoutCachedWidth =6150
                    LayoutCachedHeight =8784
                    TabIndex =31
                    BorderShade =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1290
                    Top =3270
                    Width =3345
                    Height =300
                    FontSize =10
                    TabIndex =14
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LZip"
                    ControlSource ="LZip"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =1290
                    LayoutCachedTop =3270
                    LayoutCachedWidth =4635
                    LayoutCachedHeight =3570
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =2
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
                    Left =1290
                    Top =1500
                    Width =3345
                    Height =300
                    FontSize =10
                    TabIndex =6
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LAddrLine1"
                    ControlSource ="LAddrLine1"
                    FontName ="Segoe UI"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1290
                    LayoutCachedTop =1500
                    LayoutCachedWidth =4635
                    LayoutCachedHeight =1800
                    ColumnStart =1
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =315
                            Top =1500
                            Width =912
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="Label60"
                            Caption ="Address 1"
                            FontName ="Segoe UI"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =315
                            LayoutCachedTop =1500
                            LayoutCachedWidth =1227
                            LayoutCachedHeight =1740
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
                    Left =1290
                    Top =1854
                    Width =3345
                    Height =300
                    FontSize =10
                    TabIndex =7
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LAddrLine2"
                    ControlSource ="LAddrLine2"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =1290
                    LayoutCachedTop =1854
                    LayoutCachedWidth =4635
                    LayoutCachedHeight =2154
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =315
                            Top =1854
                            Width =912
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="Label580"
                            Caption ="Address 2"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =315
                            LayoutCachedTop =1854
                            LayoutCachedWidth =1227
                            LayoutCachedHeight =2094
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
                    ListWidth =4680
                    Left =1290
                    Top =2562
                    Width =3345
                    Height =300
                    FontSize =10
                    TabIndex =11
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LCity"
                    ControlSource ="LCity"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT City, CityID, County, CountyID From vw_CCSZ_Lookup WHERE StateID =3 GROUP"
                        " BY City, CityID, County, CountyID ORDER BY City, County;"
                    ColumnWidths ="2506;0;2160;0"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1290
                    LayoutCachedTop =2562
                    LayoutCachedWidth =4635
                    LayoutCachedHeight =2862
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =900
                            Top =2562
                            Width =327
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="Label42"
                            Caption ="City"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =900
                            LayoutCachedTop =2562
                            LayoutCachedWidth =1227
                            LayoutCachedHeight =2802
                            RowStart =3
                            RowEnd =3
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
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
                    Left =1290
                    Top =2916
                    Width =3345
                    Height =300
                    FontSize =10
                    TabIndex =13
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LCounty"
                    ControlSource ="LCounty"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT County, CountyID From vw_CCSZ_Lookup WHERE StateID =3 GROUP BY County, Co"
                        "untyID ORDER BY County;"
                    ColumnWidths ="3600;0"
                    StatusBarText ="link to counties table"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1290
                    LayoutCachedTop =2916
                    LayoutCachedWidth =4635
                    LayoutCachedHeight =3216
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =645
                            Top =2916
                            Width =582
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="Label43"
                            Caption ="County"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =645
                            LayoutCachedTop =2916
                            LayoutCachedWidth =1227
                            LayoutCachedHeight =3156
                            RowStart =4
                            RowEnd =4
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
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
                    ColumnCount =5
                    ListRows =20
                    ListWidth =4320
                    Left =1290
                    Top =2208
                    Width =1860
                    Height =300
                    FontSize =10
                    TabIndex =9
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LState"
                    ControlSource ="LState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TaxStates.PostalAbbreviation, TaxStates.StateName, IIf([RentReceiptAvaila"
                        "ble]=True,\"RR\",\"--\") & \"/\" & IIf([UpfrontAvailable]=True,\"UpFront\",\"--\""
                        ") AS [Option], TaxStates.ISOCountryCode2, TaxStates.ID FROM TaxStates GROUP BY T"
                        "axStates.PostalAbbreviation, TaxStates.StateName, IIf([RentReceiptAvailable]=Tru"
                        "e,\"RR\",\"--\") & \"/\" & IIf([UpfrontAvailable]=True,\"UpFront\",\"--\"), TaxS"
                        "tates.ISOCountryCode2, TaxStates.ID, TaxStates.PostalAbbreviation HAVING (((TaxS"
                        "tates.ISOCountryCode2)=\"USA\" Or (TaxStates.ISOCountryCode2)=\"CAN\" Or (TaxSta"
                        "tes.ISOCountryCode2)=\"MEX\")) ORDER BY TaxStates.ISOCountryCode2 DESC , TaxStat"
                        "es.PostalAbbreviation;"
                    ColumnWidths ="864;1728;0;432;0"
                    StatusBarText ="lookup in lookup_state table"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1290
                    LayoutCachedTop =2208
                    LayoutCachedWidth =3150
                    LayoutCachedHeight =2508
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =3
                    Left =465
                    Top =3270
                    Width =762
                    Height =240
                    FontSize =8
                    BorderColor =0
                    Name ="Label778"
                    Caption ="Zip"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =465
                    LayoutCachedTop =3270
                    LayoutCachedWidth =1227
                    LayoutCachedHeight =3510
                    RowStart =5
                    RowEnd =5
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =3
                    Left =795
                    Top =2208
                    Width =432
                    Height =240
                    FontSize =8
                    BorderColor =0
                    Name ="Label8"
                    Caption ="State"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =795
                    LayoutCachedTop =2208
                    LayoutCachedWidth =1227
                    LayoutCachedHeight =2448
                    RowStart =2
                    RowEnd =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Subform
                    CanShrink = NotDefault
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =315
                    Top =3855
                    Width =5325
                    Height =3090
                    TabIndex =15
                    Name ="LocationTax"
                    SourceObject ="Form.LocationTaxEdit"
                    LinkChildFields ="LocationID"
                    LinkMasterFields ="LAddressID"

                    LayoutCachedLeft =315
                    LayoutCachedTop =3855
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =6945
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4795
                    Top =2115
                    Width =1022
                    Height =285
                    TabIndex =8
                    BorderColor =12566463
                    ForeColor =5855577
                    Name ="UnitsAddress"
                    ControlSource ="=UnitsAtAddress([LAddressID])"
                    Format ="Standard"
                    GridlineColor =0

                    LayoutCachedLeft =4795
                    LayoutCachedTop =2115
                    LayoutCachedWidth =5817
                    LayoutCachedHeight =2400
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =4755
                            Top =1905
                            Width =1110
                            Height =525
                            FontSize =8
                            BackColor =15921906
                            BorderColor =12566463
                            ForeColor =8210719
                            Name ="Label77"
                            Caption ="Unit Links"
                            FontName ="Segoe UI"
                            GridlineColor =0
                            LayoutCachedLeft =4755
                            LayoutCachedTop =1905
                            LayoutCachedWidth =5865
                            LayoutCachedHeight =2430
                            ThemeFontIndex =-1
                            BackShade =95.0
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1140
                    Top =8520
                    Width =750
                    FontSize =8
                    TabIndex =25
                    BorderColor =15527148
                    Name ="ClientID"
                    ControlSource ="ClientREF"
                    StatusBarText ="link to Clients table"
                    FontName ="Tahoma"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    VerticalAnchor =1

                    LayoutCachedLeft =1140
                    LayoutCachedTop =8520
                    LayoutCachedWidth =1890
                    LayoutCachedHeight =8760
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1890
                    Top =8520
                    Width =510
                    FontSize =8
                    TabIndex =26
                    BorderColor =15527148
                    Name ="LAddressID"
                    ControlSource ="LocationID"
                    FontName ="Tahoma"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    VerticalAnchor =1

                    LayoutCachedLeft =1890
                    LayoutCachedTop =8520
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =8760
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2400
                    Top =8520
                    Width =510
                    FontSize =8
                    TabIndex =27
                    BorderColor =15527148
                    Name ="LStateID"
                    ControlSource ="LStateID"
                    FontName ="Tahoma"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    VerticalAnchor =1

                    LayoutCachedLeft =2400
                    LayoutCachedTop =8520
                    LayoutCachedWidth =2910
                    LayoutCachedHeight =8760
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2910
                    Top =8520
                    Width =510
                    FontSize =8
                    TabIndex =28
                    BorderColor =15527148
                    Name ="LCountyID"
                    ControlSource ="LCountyID"
                    FontName ="Tahoma"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    VerticalAnchor =1

                    LayoutCachedLeft =2910
                    LayoutCachedTop =8520
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =8760
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3420
                    Top =8520
                    Width =510
                    FontSize =8
                    TabIndex =29
                    BorderColor =15527148
                    Name ="LCityID"
                    ControlSource ="LCityID"
                    FontName ="Tahoma"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    VerticalAnchor =1

                    LayoutCachedLeft =3420
                    LayoutCachedTop =8520
                    LayoutCachedWidth =3930
                    LayoutCachedHeight =8760
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =65.0
                    GroupTable =3
                End
                Begin Image
                    PictureTiling = NotDefault
                    BackStyle =1
                    BorderWidth =1
                    SizeMode =0
                    PictureType =2
                    Width =18720
                    Height =180
                    BackColor =15527148
                    BorderColor =14211288
                    Name ="Image142"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedWidth =18720
                    LayoutCachedHeight =180
                    TabIndex =30
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =480
                    Top =8520
                    Width =660
                    FontSize =8
                    TabIndex =24
                    BorderColor =15527148
                    Name ="ClientGroupID"
                    ControlSource ="ClientGroupID"
                    StatusBarText ="link to Clients table"
                    FontName ="Tahoma"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    VerticalAnchor =1

                    LayoutCachedLeft =480
                    LayoutCachedTop =8520
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =8760
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =65.0
                    GroupTable =3
                End
                Begin Subform
                    OverlapFlags =247
                    Left =6060
                    Top =115
                    Width =12672
                    Height =8640
                    BorderColor =14211288
                    Name ="LocationAddressList"
                    SourceObject ="Form.LocationAddressList"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =6060
                    LayoutCachedTop =115
                    LayoutCachedWidth =18732
                    LayoutCachedHeight =8755
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4740
                    Top =7200
                    Width =1110
                    Height =522
                    FontSize =10
                    TabIndex =20
                    ForeColor =0
                    Name ="DeleteMe"
                    Caption ="Delete \015\012Location"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete This Base Tax"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000696969126969698a69696951000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000069696906 ,
                        0x696969a5696969276969696f696969ff696969fc6969694e0000000000000000 ,
                        0x00000000000000000000000000000000000000000000000069696906696969b7 ,
                        0x696969b40000000069696921696969f9696969ff696969f66969693000000000 ,
                        0x000000000000000000000000000000000000000069696906696969b4696969f9 ,
                        0x6969692400000000000000006969697e696969ff696969ff696969e46969691b ,
                        0x0000000000000000000000000000000069696906696969ae696969ff69696984 ,
                        0x00000000000000000000000000000000696969a5696969ff696969ff696969d2 ,
                        0x6969690f000000000000000069696906696969ae696969ff696969d269696906 ,
                        0x0000000000000000000000000000000069696906696969b4696969ff696969ff ,
                        0x696969c66969691269696906696969ab696969ff696969f96969693000000000 ,
                        0x00000000000000000000000000000000000000006969690369696993696969ff ,
                        0x696969ff696969d8696969c0696969ff696969ff696969630000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000006969695d ,
                        0x696969f3696969ff696969ff696969ff6969699c000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000006969694e ,
                        0x696969f3696969ff696969ff696969ff69696951000000000000000000000000 ,
                        0x00000000000000000000000000000000000000006969690c69696999696969ff ,
                        0x696969ff696969ff696969cc696969e1696969ff6969698a6969690600000000 ,
                        0x0000000000000000000000000000000069696948696969e1696969ff696969ff ,
                        0x696969ff696969ab696969096969690969696984696969f6696969cc69696921 ,
                        0x0000000000000000000000006969698d696969ff696969ff696969ff696969ff ,
                        0x69696987000000000000000000000000000000006969691e69696993696969ea ,
                        0x696969786969690669696903696969f6696969ff696969ff696969d569696942 ,
                        0x0000000000000000000000000000000000000000000000000000000069696912 ,
                        0x6969696369696927000000006969694e696969b76969696f6969690600000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4740
                    LayoutCachedTop =7200
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =7722
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =1
                    BackTint =100.0
                    BackShade =95.0
                    BorderColor =2366701
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =967423
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4755
                    Top =1134
                    Width =1095
                    Height =270
                    FontSize =8
                    TabIndex =4
                    ForeColor =8210719
                    Name ="SaveMe"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff00000000 ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xd5abc3ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff00000000727272ff727272ff727272ff727272ff00000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff
                    End

                    LayoutCachedLeft =4755
                    LayoutCachedTop =1134
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =1404
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =4755
                    Top =780
                    Width =1095
                    Height =270
                    FontSize =8
                    TabIndex =3
                    ForeColor =8210719
                    Name ="AddNew"
                    Caption ="New"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x0000000000000000000000000000000082c2ea0f82c2eaf00000000000000000 ,
                        0x00000000727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x00000000000000000000000082c2eaff82c2ea1582c2eaf382c2ea0982c2eaff ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000082c2ea1b82c2eaea82c2eaff82c2eaea82c2ea21 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000082c2eaf082c2eaff82c2eaff0000000082c2eaff82c2eaff ,
                        0x82c2eaff00000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000082c2ea2182c2eaf982c2eaff82c2eaf982c2ea2a ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000082c2eaff82c2ea3682c2eaf682c2ea2d82c2eaff ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000082c2ea0f82c2eaf00000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4755
                    LayoutCachedTop =780
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =1050
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
                Begin CheckBox
                    OverlapFlags =85
                    Left =2190
                    Top =7005
                    Width =240
                    Height =255
                    TabIndex =17
                    BorderColor =8210719
                    Name ="ObsoleteYN"
                    ControlSource ="LocObsolete"
                    GroupTable =2

                    LayoutCachedLeft =2190
                    LayoutCachedTop =7005
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =7260
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =1215
                            Top =7005
                            Width =915
                            Height =255
                            FontSize =8
                            BorderColor =0
                            Name ="Obsolete"
                            Caption ="Obsolete"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =1215
                            LayoutCachedTop =7005
                            LayoutCachedWidth =2130
                            LayoutCachedHeight =7260
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2970
                    Top =7005
                    Height =255
                    FontSize =9
                    TabIndex =18
                    ForeColor =0
                    Name ="ObsoleteDate"
                    ControlSource ="LocObsdate"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    GroupTable =2

                    LayoutCachedLeft =2970
                    LayoutCachedTop =7005
                    LayoutCachedWidth =4410
                    LayoutCachedHeight =7260
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =9
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =2490
                            Top =7005
                            Width =420
                            Height =255
                            FontSize =8
                            BorderColor =0
                            Name ="LTaxRate"
                            Caption ="Date"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =2490
                            LayoutCachedTop =7005
                            LayoutCachedWidth =2910
                            LayoutCachedHeight =7260
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2190
                    Top =7320
                    Width =2220
                    FontSize =9
                    TabIndex =19
                    ForeColor =0
                    Name ="ObsoleteReason"
                    ControlSource ="LocObsCmt"
                    FontName ="Segoe UI"
                    GroupTable =2

                    LayoutCachedLeft =2190
                    LayoutCachedTop =7320
                    LayoutCachedWidth =4410
                    LayoutCachedHeight =7560
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =9
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =1215
                            Top =7320
                            Width =915
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="Label76"
                            Caption ="   Reason "
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =1215
                            LayoutCachedTop =7320
                            LayoutCachedWidth =2130
                            LayoutCachedHeight =7560
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4740
                    Top =8040
                    Width =1095
                    Height =270
                    FontSize =8
                    TabIndex =23
                    ForeColor =8210719
                    Name ="MoveUnits"
                    Caption ="Move Units"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =4740
                    LayoutCachedTop =8040
                    LayoutCachedWidth =5835
                    LayoutCachedHeight =8310
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =4755
                    Top =2580
                    Width =1095
                    Height =270
                    FontSize =8
                    TabIndex =12
                    ForeColor =8210719
                    Name ="ExpandForm"
                    Caption ="Expand>>"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =4755
                    LayoutCachedTop =2580
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =2850
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
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3255
                    Top =2208
                    Width =1380
                    Height =300
                    FontSize =10
                    TabIndex =10
                    BorderColor =8355711
                    Name ="LCountry"
                    ControlSource ="LCountry"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =3255
                    LayoutCachedTop =2208
                    LayoutCachedWidth =4635
                    LayoutCachedHeight =2508
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =390
                    Top =7005
                    Width =1590
                    Height =240
                    FontSize =8
                    TabIndex =16
                    ForeColor =8210719
                    Name ="LookUpTaxZip"
                    Caption ="Look Up Tax (Zip)"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Edit Address"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000072727254727272b5 ,
                        0x737373f8727272ff737373f8727272b572727254000000000000000000000000 ,
                        0x0000000000000000000000000000000072727215737373af808080e69f9f9ff7 ,
                        0xa5a5a5fff3f3f3ffa0a0a0ff9f9f9ff7808080e6737373af7272721500000000 ,
                        0x00000000000000000000000072727224747474e5b3b3b3e3f3f3f3ff838383ff ,
                        0xf8f8f8fffffffffff8f8f8ff848484ffefefefffb3b3b3e3747474e572727218 ,
                        0x000000000000000072727203737373d1c1c1c1e9ffffffffc7c7c7ffb4b4b4ff ,
                        0xffffffffffffffffffffffffb4b4b4ffbebebeffffffffffc1c1c1e9757575b2 ,
                        0x00000000000000007272725d858585e2ffffffffffffffff9d9d9dffd6d6d6ff ,
                        0xffffffffffffffffffffffffd6d6d6ff9d9d9dffffffffffffffffff858585e2 ,
                        0x7272725700000000737373b5727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x737373b500000000747474edf0f0f0ffffffffffffffffff797979fff8f8f8ff ,
                        0xfffffffffffffffffffffffff8f8f8ff797979fffffffffffffffffff0f0f0ff ,
                        0x747474eb00000000727272fffcfcfcffffffffffffffffff727272fffdfdfdff ,
                        0xffffffffffffffffffffffffffffffff727272fffffffffffffffffffcfcfcff ,
                        0x737373fa00000000747474edf0f0f0ffffffffffffffffff797979fff3f3f3ff ,
                        0xfffffffffffffffffffffffff8f8f8ff797979fffffffffffffffffff0f0f0ff ,
                        0x747474eb00000000737373b5727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x737373b5000000007272725a858585e2ffffffffffffffff9d9d9dffd6d6d6ff ,
                        0xffffffffffffffffffffffffd6d6d6ff9d9d9dffffffffffffffffff858585e2 ,
                        0x727272570000000072727203737373cebebebee9ffffffffc8c8c8ffb4b4b4ff ,
                        0xffffffffffffffffffffffffb4b4b4ffbebebeffffffffffbebebee9757575b2 ,
                        0x00000000000000000000000072727224747474e5aeaeaee5f2f2f2ff7f7f7fff ,
                        0xf8f8f8fffffffffff8f8f8ff838383ffedededffaeaeaee5747474e772727218 ,
                        0x0000000000000000000000000000000072727215737373af7f7f7fe89f9f9ff7 ,
                        0xa4a4a4fff8f8f8ffa3a3a3ff9f9f9ff77f7f7fe8737373af7272721500000000 ,
                        0x0000000000000000000000000000000000000000000000007272724b72727297 ,
                        0x757575eb727272fe757575eb727272977272724b000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =390
                    LayoutCachedTop =7005
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =7245
                    PictureCaptionArrangement =5
                    ColumnEnd =2
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
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1275
                    Top =360
                    Width =4425
                    Height =300
                    FontSize =10
                    TabIndex =1
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LAddrName"
                    ControlSource ="LocationName"
                    FontName ="Segoe UI"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1275
                    LayoutCachedTop =360
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =660
                    ColumnStart =1
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =300
                            Top =360
                            Width =912
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="Label167"
                            Caption ="Name:"
                            FontName ="Segoe UI"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =300
                            LayoutCachedTop =360
                            LayoutCachedWidth =1212
                            LayoutCachedHeight =600
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
                    Top =780
                    Width =3345
                    Height =300
                    FontSize =10
                    TabIndex =2
                    BorderColor =8210719
                    ForeColor =0
                    Name ="Text168"
                    ControlSource ="LocationSN"
                    FontName ="Segoe UI"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1275
                    LayoutCachedTop =780
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =1080
                    ColumnStart =1
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =192
                            Top =780
                            Width =1020
                            Height =255
                            FontSize =8
                            BorderColor =0
                            Name ="Label169"
                            Caption ="Short Name:"
                            FontName ="Segoe UI"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =192
                            LayoutCachedTop =780
                            LayoutCachedWidth =1212
                            LayoutCachedHeight =1035
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
                    Top =1140
                    Width =3345
                    Height =300
                    FontSize =10
                    TabIndex =5
                    BorderColor =8210719
                    ForeColor =0
                    Name ="Text172"
                    ControlSource ="LocationStreet"
                    FontName ="Segoe UI"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1275
                    LayoutCachedTop =1140
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =1440
                    ColumnStart =1
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =300
                            Top =1140
                            Width =912
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="Label173"
                            Caption ="Company"
                            FontName ="Segoe UI"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =300
                            LayoutCachedTop =1140
                            LayoutCachedWidth =1212
                            LayoutCachedHeight =1380
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1035
                    Top =7620
                    Width =3345
                    Height =300
                    FontSize =10
                    TabIndex =21
                    BorderColor =8210719
                    ForeColor =0
                    Name ="LPhone"
                    ControlSource ="LPhone"
                    FontName ="Segoe UI"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =1035
                    LayoutCachedTop =7620
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =7920
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =60
                            Top =7620
                            Width =912
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="Label175"
                            Caption ="Phone"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =60
                            LayoutCachedTop =7620
                            LayoutCachedWidth =972
                            LayoutCachedHeight =7860
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
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1035
                    Top =7920
                    Width =3345
                    Height =300
                    FontSize =10
                    TabIndex =22
                    BorderColor =8210719
                    ForeColor =0
                    Name ="Text176"
                    ControlSource ="LEmail"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =1035
                    LayoutCachedTop =7920
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =8220
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =60
                            Top =7860
                            Width =912
                            Height =240
                            FontSize =8
                            BorderColor =0
                            Name ="Label177"
                            Caption ="Email"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =60
                            LayoutCachedTop =7860
                            LayoutCachedWidth =972
                            LayoutCachedHeight =8100
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
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =1890
                    Top =8280
                    Width =510
                    Height =240
                    FontSize =8
                    FontWeight =700
                    BorderColor =0
                    ForeColor =4754549
                    Name ="Label179"
                    Caption ="LocID"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    VerticalAnchor =1
                    LayoutCachedLeft =1890
                    LayoutCachedTop =8280
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =8520
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =480
                    Top =8280
                    Width =660
                    Height =240
                    FontSize =8
                    FontWeight =700
                    BorderColor =0
                    ForeColor =4754549
                    Name ="Label180"
                    Caption ="ClientId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    VerticalAnchor =1
                    LayoutCachedLeft =480
                    LayoutCachedTop =8280
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =8520
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =1140
                    Top =8280
                    Width =750
                    Height =240
                    FontSize =8
                    FontWeight =700
                    BorderColor =0
                    ForeColor =4754549
                    Name ="Label187"
                    Caption ="ClGrpID"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    VerticalAnchor =1
                    LayoutCachedLeft =1140
                    LayoutCachedTop =8280
                    LayoutCachedWidth =1890
                    LayoutCachedHeight =8520
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =2400
                    Top =8280
                    Width =510
                    Name ="EmptyCell203"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    VerticalAnchor =1
                    LayoutCachedLeft =2400
                    LayoutCachedTop =8280
                    LayoutCachedWidth =2910
                    LayoutCachedHeight =8520
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =2910
                    Top =8280
                    Width =510
                    Name ="EmptyCell207"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    VerticalAnchor =1
                    LayoutCachedLeft =2910
                    LayoutCachedTop =8280
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =8520
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =3420
                    Top =8280
                    Width =510
                    Name ="EmptyCell211"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    VerticalAnchor =1
                    LayoutCachedLeft =3420
                    LayoutCachedTop =8280
                    LayoutCachedWidth =3930
                    LayoutCachedHeight =8520
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    GroupTable =3
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
' See "LocationAddressesEdit.cls"
