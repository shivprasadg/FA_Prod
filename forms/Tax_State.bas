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
    ScrollBars =2
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =17295
    DatasheetFontHeight =11
    ItemSuffix =124
    Left =8580
    Top =5730
    Right =26130
    Bottom =15630
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x1508280a0df0e440
    End
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
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin Tab
            TextFontFamily =18
            FontSize =11
            FontName ="Cambria"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
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
        Begin FormHeader
            Height =600
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =8640
                    Width =2040
                    Height =210
                    FontSize =8
                    Name ="Label119"
                    Caption ="Cou    State  Cnty   City    Act"
                    LayoutCachedLeft =8640
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =210
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =420
                    Width =17280
                    Height =180
                    Name ="Image142"
                    Picture ="3_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =420
                    LayoutCachedWidth =17280
                    LayoutCachedHeight =600
                    TabIndex =7
                End
                Begin Label
                    OverlapFlags =255
                    Left =2760
                    Width =5790
                    Height =495
                    FontSize =20
                    FontWeight =700
                    ForeColor =12349952
                    Name ="Label50"
                    Caption ="- States"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =2760
                    LayoutCachedWidth =8550
                    LayoutCachedHeight =495
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9048
                    Top =173
                    Width =408
                    Height =231
                    ColumnOrder =3
                    FontSize =8
                    Name ="StateID"
                    TopPadding =0

                    LayoutCachedLeft =9048
                    LayoutCachedTop =173
                    LayoutCachedWidth =9456
                    LayoutCachedHeight =404
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8640
                    Top =173
                    Width =408
                    Height =231
                    ColumnOrder =4
                    FontSize =8
                    TabIndex =1
                    Name ="CountryID"
                    TopPadding =0

                    LayoutCachedLeft =8640
                    LayoutCachedTop =173
                    LayoutCachedWidth =9048
                    LayoutCachedHeight =404
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9456
                    Top =173
                    Width =408
                    Height =231
                    ColumnOrder =5
                    FontSize =8
                    TabIndex =2
                    Name ="CountyID"
                    TopPadding =0

                    LayoutCachedLeft =9456
                    LayoutCachedTop =173
                    LayoutCachedWidth =9864
                    LayoutCachedHeight =404
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    Left =120
                    Width =2640
                    Height =495
                    FontSize =20
                    FontWeight =700
                    Name ="Label105"
                    Caption ="Tax Manager"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =495
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10272
                    Top =173
                    Width =408
                    Height =231
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =3
                    Name ="TaxTypeID"
                    TopPadding =0

                    LayoutCachedLeft =10272
                    LayoutCachedTop =173
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =404
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9864
                    Top =173
                    Width =408
                    Height =231
                    ColumnOrder =0
                    FontSize =8
                    TabIndex =4
                    Name ="CityID"
                    TopPadding =0

                    LayoutCachedLeft =9864
                    LayoutCachedTop =173
                    LayoutCachedWidth =10272
                    LayoutCachedHeight =404
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =127
                    Left =14100
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =5
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="FAMainMenu"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"GoHome\" xmlns=\"http://schemas.microsoft.com/office/accesss"
                                "ervices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument Na"
                                "me=\"FormName\">FAMainMenu</Argum"
                        End
                        Begin
                            Comment ="_AXL:ent></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14100
                    LayoutCachedTop =120
                    LayoutCachedWidth =14655
                    LayoutCachedHeight =348
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    PictureType =1
                    Left =14667
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =6
                    Name ="CloseMe"
                    Caption ="Close"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1

                    LayoutCachedLeft =14667
                    LayoutCachedTop =120
                    LayoutCachedWidth =15222
                    LayoutCachedHeight =348
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =9
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =9317
            Name ="Detail"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =8759
                    Top =958
                    Width =8535
                    Height =446
                    BackColor =14211288
                    Name ="Box124"
                    HorizontalAnchor =2
                    LayoutCachedLeft =8759
                    LayoutCachedTop =958
                    LayoutCachedWidth =17294
                    LayoutCachedHeight =1404
                End
                Begin Subform
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =225
                    Top =960
                    Width =8640
                    Height =903
                    BorderColor =9211020
                    Name ="Tax_StateSub"
                    SourceObject ="Form.Tax_StateSub"
                    LinkChildFields ="Id"
                    LinkMasterFields ="StateID"
                    VerticalAnchor =2

                    LayoutCachedLeft =225
                    LayoutCachedTop =960
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =1863
                End
                Begin Label
                    OverlapFlags =95
                    Left =225
                    Top =571
                    Width =8640
                    Height =389
                    FontSize =14
                    ForeColor =8210719
                    Name ="StateLabel2"
                    Caption ="States"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =225
                    LayoutCachedTop =571
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =960
                End
                Begin Label
                    OverlapFlags =85
                    Left =225
                    Top =1944
                    Width =8580
                    Height =389
                    FontSize =14
                    ForeColor =8210719
                    Name ="StateLabel3"
                    Caption ="State Taxes"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =225
                    LayoutCachedTop =1944
                    LayoutCachedWidth =8805
                    LayoutCachedHeight =2333
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =12048
                    Top =1044
                    Width =2280
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="AddTaxRate"
                    Caption ="   Add A State Tax"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000505050ff505050ff00000000505050ff505050ff00000000 ,
                        0x505050ff505050ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000505050ff505050ff00000000505050ff505050ff00000000 ,
                        0x505050ff505050ff000000000000000000000000000000000000000072727248 ,
                        0x727272ff72727254000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000072727248727272f6 ,
                        0x727272ff727272ff1c75e3261c75e3ff1c75e326000000000000000000000000 ,
                        0x000000000000000000000000000000000000000072727248727272f6727272ff ,
                        0x727272fc727272541c75e3ff000000001c75e3ff000000001c75e3261c75e3ff ,
                        0x1c75e32600000000000000000000000072727248727272f6727272ff727272fc ,
                        0x72727254000000001c75e3261c75e3ff1c75e326000000001c75e3ff00000000 ,
                        0x1c75e3ff000000000000000072727248727272f6727272ff727272fc72727254 ,
                        0x0000000000000000000000000000000000000000000000001c75e3261c75e3ff ,
                        0x1c75e3260000000072727248727272f6727272ff727272fc7272725400000000 ,
                        0x727272ff727272ff000000001c75e3261c75e3ff1c75e3260000000000000000 ,
                        0x0000000072727248727272f6727272ff727272f972727254ffffff48fffffff9 ,
                        0xffffffff727272ff000000001c75e3ff000000001c75e3ff0000000000000000 ,
                        0x72727248727272f6727272ff727272f972727254f6f6f64ec8c8c8fdb3b3b3ff ,
                        0xffffffff727272ff000000001c75e3261c75e3ff1c75e3260000000072727248 ,
                        0x727272f6ffffffff727272f972727254f6f6f64efffffff9ffffffffb3b3b3ff ,
                        0xffffffff727272ff0000000000000000000000000000000072727248727272f6 ,
                        0xffffffff727272f972727254ffffff48c8c8c8fdb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xffffffff727272ff00000000000000000000000000000000727272ff727272ff ,
                        0x727272f972727254ffffff48fffffff9ffffffffffffffffffffffffffffffff ,
                        0xffffffff727272ff0000000000000000000000000000000072727254727272ff ,
                        0x72727254f6f6f64ec8c8c8fdffffffffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xffffffff727272ff000000000000000000000000000000000000000000000000 ,
                        0xffffff48fffffff9ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff727272ff0000000000000000000000000000000000000000b17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff0000000000000000000000000000000000000000b17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff
                    End

                    LayoutCachedLeft =12048
                    LayoutCachedTop =1044
                    LayoutCachedWidth =14328
                    LayoutCachedHeight =1329
                    PictureCaptionArrangement =5
                    RowStart =2
                    RowEnd =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderWidth =1
                    BorderColor =16777215
                    ThemeFontIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-753
                    Overlaps =1
                End
                Begin Subform
                    CanShrink = NotDefault
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =225
                    Top =2364
                    Width =17055
                    Height =2643
                    TabIndex =2
                    Name ="Tax_StateBaseSub"
                    SourceObject ="Form.Tax_AnyBaseSub"
                    LinkChildFields ="LinkID"
                    LinkMasterFields ="StateID"
                    HorizontalAnchor =2

                    LayoutCachedLeft =225
                    LayoutCachedTop =2364
                    LayoutCachedWidth =17280
                    LayoutCachedHeight =5007
                End
                Begin Subform
                    CanShrink = NotDefault
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =240
                    Top =5429
                    Width =17055
                    Height =3888
                    TabIndex =3
                    Name ="Tax_StateExceptionsSub"
                    SourceObject ="Form.Tax_Any_ExceptionsSub"
                    LinkChildFields ="TaxStatesID"
                    LinkMasterFields ="StateID"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =240
                    LayoutCachedTop =5429
                    LayoutCachedWidth =17295
                    LayoutCachedHeight =9317
                End
                Begin Label
                    OverlapFlags =93
                    Left =240
                    Top =5040
                    Width =8445
                    Height =389
                    FontSize =14
                    ForeColor =8210719
                    Name ="StateLabel4"
                    Caption ="State Exceptions"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =240
                    LayoutCachedTop =5040
                    LayoutCachedWidth =8685
                    LayoutCachedHeight =5429
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =4320
                    Left =1335
                    Top =75
                    Width =1051
                    Height =390
                    FontSize =12
                    TabIndex =4
                    Name ="SelectCountry"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TaxCountries.ISOCode3, TaxCountries.CountryName, TaxCountries.ISOCode2, T"
                        "axCountries.ISONumber FROM TaxCountries;"
                    ColumnWidths ="432;1440;0;0;0"
                    ValidationRule ="Is Not Null"
                    FontName ="Segoe UI"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1335
                    LayoutCachedTop =75
                    LayoutCachedWidth =2386
                    LayoutCachedHeight =465
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =225
                            Width =2217
                            Height =525
                            FontSize =10
                            FontWeight =700
                            LeftMargin =45
                            TopMargin =45
                            RightMargin =45
                            BottomMargin =45
                            Name ="Label1"
                            Caption ="Country"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =225
                            LayoutCachedWidth =2442
                            LayoutCachedHeight =525
                            BackThemeColorIndex =2
                            BackTint =60.0
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =5760
                    Left =3645
                    Top =75
                    Width =4320
                    Height =390
                    TabIndex =5
                    Name ="SelectState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TaxStates.Id, TaxStates.StateName, TaxStates.PostalAbbreviation AS State,"
                        " TaxStates.SalesTaxOrUseTaxNumber, TaxStates.ISOCountryCode2 FROM TaxStates WHER"
                        "E (((TaxStates.ISOCountryCode2)='US')) ORDER BY TaxStates.StateName, TaxStates.P"
                        "ostalAbbreviation;"
                    ColumnWidths ="0;2160;432;2880;0"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =3645
                    LayoutCachedTop =75
                    LayoutCachedWidth =7965
                    LayoutCachedHeight =465
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =2520
                            Width =5520
                            Height =525
                            FontSize =10
                            FontWeight =700
                            LeftMargin =45
                            TopMargin =45
                            RightMargin =45
                            BottomMargin =45
                            Name ="Label122"
                            Caption ="State"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =2520
                            LayoutCachedWidth =8040
                            LayoutCachedHeight =525
                            BackThemeColorIndex =4
                            BackShade =75.0
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =9648
                    Top =1044
                    Width =2280
                    Height =285
                    FontSize =10
                    TabIndex =6
                    Name ="UnitSummary"
                    Caption =" Future"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =9648
                    LayoutCachedTop =1044
                    LayoutCachedWidth =11928
                    LayoutCachedHeight =1329
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderColor =16777215
                    ThemeFontIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "Tax_State.cls"
