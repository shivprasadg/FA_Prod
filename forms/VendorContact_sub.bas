Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15600
    DatasheetFontHeight =11
    ItemSuffix =409
    Left =7935
    Top =9015
    Right =25140
    Bottom =15210
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xc03a04baf596e540
    End
    RecordSource ="VendorContacts"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
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
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin Tab
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =0
            BackColor =15709952
            Name ="FormHeader"
        End
        Begin Section
            Height =6555
            Name ="Detail"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Left =2880
                    Top =3540
                    Width =12714
                    Height =225
                    Name ="Image142"
                    Picture ="3_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedLeft =2880
                    LayoutCachedTop =3540
                    LayoutCachedWidth =15594
                    LayoutCachedHeight =3765
                    TabIndex =24
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =12690
                    Top =525
                    TabIndex =12
                    Name ="ContactObsolete"
                    ControlSource ="ContactObsolete"
                    StatusBarText ="Double Click the Unlock button to unlock this field"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Double Click the Unlock button to unlock this field"

                    LayoutCachedLeft =12690
                    LayoutCachedTop =525
                    LayoutCachedWidth =12950
                    LayoutCachedHeight =765
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =12975
                            Top =465
                            Width =1050
                            Height =270
                            FontSize =10
                            BackColor =15709952
                            Name ="Label49"
                            Caption ="   Obsolete "
                            LayoutCachedLeft =12975
                            LayoutCachedTop =465
                            LayoutCachedWidth =14025
                            LayoutCachedHeight =735
                            BorderThemeColorIndex =8
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =8652
                    Top =2508
                    Width =720
                    Height =252
                    TabIndex =10
                    Name ="contactOptOut"
                    ControlSource ="VCEMoptout"
                    StatusBarText ="email opt out"
                    GroupTable =1

                    LayoutCachedLeft =8652
                    LayoutCachedTop =2508
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =2760
                    RowStart =7
                    RowEnd =7
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =7200
                            Top =2508
                            Width =1392
                            Height =252
                            FontSize =9
                            Name ="Label25"
                            Caption ="Opt Out of Email"
                            GroupTable =1
                            LayoutCachedLeft =7200
                            LayoutCachedTop =2508
                            LayoutCachedWidth =8592
                            LayoutCachedHeight =2760
                            RowStart =7
                            RowEnd =7
                            ColumnStart =5
                            ColumnEnd =5
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3948
                    Top =1536
                    Width =1632
                    Height =252
                    FontSize =9
                    TabIndex =6
                    BorderColor =11830108
                    Name ="contactPrimaryPhone"
                    ControlSource ="contactPrimaryPhone"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =1

                    LayoutCachedLeft =3948
                    LayoutCachedTop =1536
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =1788
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2988
                            Top =1536
                            Width =888
                            Height =252
                            FontSize =9
                            Name ="Label27"
                            Caption ="Pri Phone:"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =1536
                            LayoutCachedWidth =3876
                            LayoutCachedHeight =1788
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3948
                    Top =1860
                    Width =1632
                    Height =252
                    FontSize =9
                    TabIndex =7
                    BorderColor =11830108
                    Name ="contactCell"
                    ControlSource ="contactCell"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =1

                    LayoutCachedLeft =3948
                    LayoutCachedTop =1860
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =2112
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2988
                            Top =1860
                            Width =888
                            Height =252
                            FontSize =9
                            Name ="Label38"
                            Caption ="Cellular:"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =1860
                            LayoutCachedWidth =3876
                            LayoutCachedHeight =2112
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =9570
                    Top =1605
                    Width =4740
                    Height =1653
                    FontSize =9
                    TabIndex =18
                    BackColor =13434879
                    BorderColor =11830108
                    Name ="contactNote"
                    ControlSource ="contactNote"
                    FontName ="Tahoma"

                    LayoutCachedLeft =9570
                    LayoutCachedTop =1605
                    LayoutCachedWidth =14310
                    LayoutCachedHeight =3258
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =9571
                            Top =1410
                            Width =480
                            Height =255
                            FontSize =9
                            FontWeight =700
                            Name ="Label39"
                            Caption ="Notes:"
                            LayoutCachedLeft =9571
                            LayoutCachedTop =1410
                            LayoutCachedWidth =10051
                            LayoutCachedHeight =1665
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3948
                    Top =2184
                    Width =1632
                    Height =252
                    FontSize =9
                    TabIndex =8
                    BorderColor =11830108
                    Name ="contactFax"
                    ControlSource ="contactFax"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =1

                    LayoutCachedLeft =3948
                    LayoutCachedTop =2184
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =2436
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2988
                            Top =2184
                            Width =888
                            Height =252
                            FontSize =9
                            Name ="Label48"
                            Caption ="Fax:"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =2184
                            LayoutCachedWidth =3876
                            LayoutCachedHeight =2436
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12324
                    Top =5988
                    Width =576
                    Height =192
                    FontSize =8
                    TabIndex =21
                    Name ="ClientId"
                    ControlSource ="VendorRef"
                    StatusBarText ="link to Clients table"
                    DefaultValue ="3840"
                    GroupTable =11
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12324
                    LayoutCachedTop =5988
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =6180
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =11676
                            Top =5988
                            Width =636
                            Height =192
                            FontSize =8
                            Name ="Label26"
                            Caption ="VendorId"
                            GroupTable =11
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =11676
                            LayoutCachedTop =5988
                            LayoutCachedWidth =12312
                            LayoutCachedHeight =6180
                            LayoutGroup =2
                            BorderThemeColorIndex =1
                            BorderShade =65.0
                            ForeThemeColorIndex =1
                            ForeShade =65.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =11
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13716
                    Top =5988
                    Width =576
                    Height =192
                    FontSize =8
                    TabIndex =22
                    Name ="contactID"
                    ControlSource ="contactID"
                    GroupTable =11
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13716
                    LayoutCachedTop =5988
                    LayoutCachedWidth =14292
                    LayoutCachedHeight =6180
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =11
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Left =12912
                            Top =5988
                            Width =792
                            Height =192
                            FontSize =8
                            Name ="Label43"
                            Caption ="contactID:"
                            GroupTable =11
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =12912
                            LayoutCachedTop =5988
                            LayoutCachedWidth =13704
                            LayoutCachedHeight =6180
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            BorderThemeColorIndex =1
                            BorderShade =65.0
                            ForeThemeColorIndex =1
                            ForeShade =65.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =11
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3948
                    Top =2508
                    Width =3180
                    Height =252
                    FontSize =9
                    TabIndex =9
                    BorderColor =11830108
                    Name ="contactEmail"
                    ControlSource ="contactEmail"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =3948
                    LayoutCachedTop =2508
                    LayoutCachedWidth =7128
                    LayoutCachedHeight =2760
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2988
                            Top =2508
                            Width =888
                            Height =252
                            FontSize =9
                            Name ="Label35"
                            Caption ="E-mail:"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =2508
                            LayoutCachedWidth =3876
                            LayoutCachedHeight =2760
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =119
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =8
                    Left =3974
                    Top =3300
                    Width =10350
                    FontSize =9
                    TabIndex =20
                    BorderColor =11830108
                    Name ="Locref"
                    ControlSource ="cLocationID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LocationStreet, L"
                        "ocations.LAddrLine1, Locations.LAddrLine2, Locations.LCity, Locations.LState, Lo"
                        "cations.LZip FROM Locations WHERE (((Locations.ClientREF)=[Forms]![VENDOR FORM]!"
                        "[VendorContact_sub].[Form]![contactID]) AND ((Locations.LocObsolete)=No));"
                    ColumnWidths ="0;2160;2160;2160;1800;1800;504;7200"
                    StatusBarText ="lookup to Location Table"
                    FontName ="Tahoma"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3974
                    LayoutCachedTop =3300
                    LayoutCachedWidth =14324
                    LayoutCachedHeight =3540
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextAlign =1
                            Left =3000
                            Top =3300
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label58"
                            Caption ="Address:"
                            LayoutCachedLeft =3000
                            LayoutCachedTop =3300
                            LayoutCachedWidth =3885
                            LayoutCachedHeight =3540
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3948
                    Top =540
                    Width =1008
                    Height =264
                    FontSize =9
                    BorderColor =11830108
                    Name ="contactFirst"
                    ControlSource ="contactFirst"
                    ValidationRule ="Is Not Null"
                    ValidationText ="First Name is required"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =3948
                    LayoutCachedTop =540
                    LayoutCachedWidth =4956
                    LayoutCachedHeight =804
                    DisplayAsHyperlink =2
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2988
                            Top =540
                            Width =888
                            Height =264
                            FontSize =9
                            Name ="Label394"
                            Caption ="Name:"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =540
                            LayoutCachedWidth =3876
                            LayoutCachedHeight =804
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5028
                    Top =540
                    Width =552
                    Height =264
                    FontSize =9
                    TabIndex =1
                    BorderColor =11830108
                    Name ="contactMiddle"
                    ControlSource ="contactMiddle"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =5028
                    LayoutCachedTop =540
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5652
                    Top =540
                    Width =1476
                    Height =264
                    FontSize =9
                    TabIndex =2
                    BorderColor =11830108
                    Name ="contactLast"
                    ControlSource ="contactLast"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Last Name is required."
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =5652
                    LayoutCachedTop =540
                    LayoutCachedWidth =7128
                    LayoutCachedHeight =804
                    DisplayAsHyperlink =2
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3948
                    Top =1200
                    Width =3180
                    Height =264
                    FontSize =9
                    TabIndex =5
                    BorderColor =11830108
                    Name ="contactTitle"
                    ControlSource ="contactTitle"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =3948
                    LayoutCachedTop =1200
                    LayoutCachedWidth =7128
                    LayoutCachedHeight =1464
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2988
                            Top =1200
                            Width =888
                            Height =264
                            FontSize =9
                            Name ="Label71"
                            Caption ="Title:"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =1200
                            LayoutCachedWidth =3876
                            LayoutCachedHeight =1464
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7200
                    Top =876
                    Width =1392
                    Height =252
                    FontSize =9
                    TabIndex =4
                    Name ="MakeFullContactBtn"
                    Caption ="◄ update"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7200
                    LayoutCachedTop =876
                    LayoutCachedWidth =8592
                    LayoutCachedHeight =1128
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    ThemeFontIndex =1
                    HoverThemeColorIndex =2
                    GroupTable =1
                    QuickStyle =9
                    QuickStyleMask =-241
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3948
                    Top =876
                    Width =3180
                    Height =252
                    FontSize =9
                    TabIndex =3
                    BackColor =-2147483615
                    BorderColor =15062992
                    ForeColor =16777215
                    Name ="contactFullName"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =3948
                    LayoutCachedTop =876
                    LayoutCachedWidth =7128
                    LayoutCachedHeight =1128
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2988
                            Top =876
                            Width =888
                            Height =252
                            FontSize =9
                            Name ="Label42"
                            Caption ="Full Name:"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =876
                            LayoutCachedWidth =3876
                            LayoutCachedHeight =1128
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12675
                    Top =1005
                    Width =1275
                    FontSize =9
                    TabIndex =16
                    Name ="ConObsoleteDt"
                    ControlSource ="ContactObsoletedt"
                    Format ="Short Date"
                    StatusBarText ="Double Click the Unlock button to unlock this field"
                    ControlTipText ="Double Click the Unlock button to unlock this field"

                    LayoutCachedLeft =12675
                    LayoutCachedTop =1005
                    LayoutCachedWidth =13950
                    LayoutCachedHeight =1245
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextFontCharSet =238
                            TextAlign =2
                            Left =12615
                            Top =765
                            Width =1425
                            Height =525
                            FontSize =10
                            FontWeight =700
                            BackColor =12493372
                            ForeColor =16777215
                            Name ="Label50"
                            Caption ="As of Date:"
                            LayoutCachedLeft =12615
                            LayoutCachedTop =765
                            LayoutCachedWidth =14040
                            LayoutCachedHeight =1290
                            BorderThemeColorIndex =8
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextFontCharSet =163
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9600
                    Top =4155
                    Width =1125
                    FontSize =9
                    TabIndex =19
                    Name ="Vclastchdt"
                    ControlSource ="Vclastchdt"
                    StatusBarText ="updated when contact unlocked to make change"

                    LayoutCachedLeft =9600
                    LayoutCachedTop =4155
                    LayoutCachedWidth =10725
                    LayoutCachedHeight =4395
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =163
                            TextAlign =2
                            Left =9570
                            Top =3915
                            Width =1215
                            Height =555
                            FontSize =9
                            BackColor =11525325
                            Name ="Label69"
                            Caption ="Class Change:"
                            LayoutCachedLeft =9570
                            LayoutCachedTop =3915
                            LayoutCachedWidth =10785
                            LayoutCachedHeight =4470
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =11655
                    Top =555
                    Width =720
                    Height =660
                    FontSize =9
                    FontWeight =700
                    TabIndex =13
                    Name ="Unlockit"
                    Caption ="Unlock"
                    StatusBarText ="Double Click to Unlock"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Double Click to Unlock"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11655
                    LayoutCachedTop =555
                    LayoutCachedWidth =12375
                    LayoutCachedHeight =1215
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =8
                    BorderThemeColorIndex =8
                    BorderShade =50.0
                    HoverThemeColorIndex =8
                    HoverShade =50.0
                    PressedThemeColorIndex =8
                    QuickStyle =34
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3948
                    Top =276
                    Width =1008
                    Height =192
                    FontSize =9
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label30"
                    Caption ="First¹"
                    GroupTable =1
                    LayoutCachedLeft =3948
                    LayoutCachedTop =276
                    LayoutCachedWidth =4956
                    LayoutCachedHeight =468
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5028
                    Top =276
                    Width =552
                    Height =192
                    FontSize =9
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label34"
                    Caption ="Middle"
                    GroupTable =1
                    LayoutCachedLeft =5028
                    LayoutCachedTop =276
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =468
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5652
                    Top =276
                    Width =1476
                    Height =192
                    FontSize =9
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label32"
                    Caption ="Last¹"
                    GroupTable =1
                    LayoutCachedLeft =5652
                    LayoutCachedTop =276
                    LayoutCachedWidth =7128
                    LayoutCachedHeight =468
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =11460
                    Top =360
                    Width =2820
                    Height =1005
                    Name ="Box75"
                    LayoutCachedLeft =11460
                    LayoutCachedTop =360
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =1365
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =119
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =75
                    Top =570
                    Width =2805
                    Height =5850
                    FontSize =9
                    TabIndex =14
                    Name ="ContactsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT VendorContacts.contactID, VendorContacts.VendorRef, VendorContacts.Contac"
                        "tTypeId, [contactFirst] & IIf(IsNull([ContactMiddle]),Null,\" \" & [ContactMiddl"
                        "e]) & [ContactLast] AS Contact FROM VendorContacts ORDER BY [contactFirst] & IIf"
                        "(IsNull([ContactMiddle]),Null,\" \" & [ContactMiddle]) & [ContactLast];"
                    ColumnWidths ="0;0;1008;1728"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =75
                    LayoutCachedTop =570
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =6420
                End
                Begin Label
                    OverlapFlags =93
                    Left =60
                    Top =30
                    Width =3015
                    Height =465
                    FontSize =18
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label80"
                    Caption ="Vendor Contacts"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =60
                    LayoutCachedTop =30
                    LayoutCachedWidth =3075
                    LayoutCachedHeight =495
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9900
                    Top =705
                    Width =1395
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =15
                    Name ="AddContact"
                    Caption ="Add Contact"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =9900
                    LayoutCachedTop =705
                    LayoutCachedWidth =11295
                    LayoutCachedHeight =960
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =2
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    PressedForeColor =62207
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9900
                    Top =375
                    Width =1395
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =11
                    Name ="SaveUpdate"
                    Caption ="Save | Update"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9900
                    LayoutCachedTop =375
                    LayoutCachedWidth =11295
                    LayoutCachedHeight =630
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =2
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    PressedForeColor =62207
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =2940
                    Top =3945
                    Width =1725
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =23
                    Name ="DeleteContact"
                    Caption ="DELETE  This Contact"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2940
                    LayoutCachedTop =3945
                    LayoutCachedWidth =4665
                    LayoutCachedHeight =4185
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =2366701
                    BorderThemeColorIndex =9
                    BorderShade =75.0
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedColor =2366701
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =3
                    Left =12615
                    Top =465
                    Width =375
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BackColor =15709952
                    Name ="Label84"
                    Caption ="   "
                    LayoutCachedLeft =12615
                    LayoutCachedTop =465
                    LayoutCachedWidth =12990
                    LayoutCachedHeight =735
                    BorderThemeColorIndex =8
                    ForeThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9900
                    Top =1020
                    Width =1395
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =17
                    Name ="Printcontacts"
                    Caption ="Print Contacts"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =9900
                    LayoutCachedTop =1020
                    LayoutCachedWidth =11295
                    LayoutCachedHeight =1275
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =2
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    PressedThemeColorIndex =2
                    PressedForeColor =62207
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =2988
                    Top =276
                    Width =888
                    Height =192
                    Name ="EmptyCell121"
                    GroupTable =1
                    LayoutCachedLeft =2988
                    LayoutCachedTop =276
                    LayoutCachedWidth =3876
                    LayoutCachedHeight =468
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8652
                    Top =876
                    Width =720
                    Height =252
                    Name ="EmptyCell134"
                    GroupTable =1
                    LayoutCachedLeft =8652
                    LayoutCachedTop =876
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =1128
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7200
                    Top =1200
                    Width =1392
                    Height =264
                    Name ="EmptyCell146"
                    GroupTable =1
                    LayoutCachedLeft =7200
                    LayoutCachedTop =1200
                    LayoutCachedWidth =8592
                    LayoutCachedHeight =1464
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8652
                    Top =1200
                    Width =720
                    Height =264
                    Name ="EmptyCell147"
                    GroupTable =1
                    LayoutCachedLeft =8652
                    LayoutCachedTop =1200
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =1464
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7200
                    Top =1536
                    Width =1392
                    Height =252
                    Name ="EmptyCell156"
                    GroupTable =1
                    LayoutCachedLeft =7200
                    LayoutCachedTop =1536
                    LayoutCachedWidth =8592
                    LayoutCachedHeight =1788
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8652
                    Top =1536
                    Width =720
                    Height =252
                    Name ="EmptyCell157"
                    GroupTable =1
                    LayoutCachedLeft =8652
                    LayoutCachedTop =1536
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =1788
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5652
                    Top =1860
                    Width =1476
                    Height =252
                    Name ="EmptyCell194"
                    GroupTable =1
                    LayoutCachedLeft =5652
                    LayoutCachedTop =1860
                    LayoutCachedWidth =7128
                    LayoutCachedHeight =2112
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7200
                    Top =1860
                    Width =1392
                    Height =252
                    Name ="EmptyCell196"
                    GroupTable =1
                    LayoutCachedLeft =7200
                    LayoutCachedTop =1860
                    LayoutCachedWidth =8592
                    LayoutCachedHeight =2112
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8652
                    Top =1860
                    Width =720
                    Height =252
                    Name ="EmptyCell197"
                    GroupTable =1
                    LayoutCachedLeft =8652
                    LayoutCachedTop =1860
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =2112
                    RowStart =5
                    RowEnd =5
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5652
                    Top =2184
                    Width =1476
                    Height =252
                    Name ="EmptyCell208"
                    GroupTable =1
                    LayoutCachedLeft =5652
                    LayoutCachedTop =2184
                    LayoutCachedWidth =7128
                    LayoutCachedHeight =2436
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7200
                    Top =2184
                    Width =1392
                    Height =252
                    Name ="EmptyCell209"
                    GroupTable =1
                    LayoutCachedLeft =7200
                    LayoutCachedTop =2184
                    LayoutCachedWidth =8592
                    LayoutCachedHeight =2436
                    RowStart =6
                    RowEnd =6
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8652
                    Top =2184
                    Width =720
                    Height =252
                    Name ="EmptyCell210"
                    GroupTable =1
                    LayoutCachedLeft =8652
                    LayoutCachedTop =2184
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =2436
                    RowStart =6
                    RowEnd =6
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2988
                    Top =2820
                    Width =888
                    Name ="EmptyCell386"
                    GroupTable =1
                    LayoutCachedLeft =2988
                    LayoutCachedTop =2820
                    LayoutCachedWidth =3876
                    LayoutCachedHeight =3060
                    RowStart =8
                    RowEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3948
                    Top =2820
                    Width =2304
                    Name ="EmptyCell387"
                    GroupTable =1
                    LayoutCachedLeft =3948
                    LayoutCachedTop =2820
                    LayoutCachedWidth =6252
                    LayoutCachedHeight =3060
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6312
                    Top =2820
                    Width =3060
                    Name ="EmptyCell388"
                    GroupTable =1
                    LayoutCachedLeft =6312
                    LayoutCachedTop =2820
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =3060
                    RowStart =8
                    RowEnd =8
                    ColumnStart =4
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8652
                    Top =540
                    Width =720
                    Height =264
                    Name ="EmptyCell389"
                    GroupTable =1
                    LayoutCachedLeft =8652
                    LayoutCachedTop =540
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8652
                    Top =276
                    Width =720
                    Height =192
                    Name ="EmptyCell390"
                    GroupTable =1
                    LayoutCachedLeft =8652
                    LayoutCachedTop =276
                    LayoutCachedWidth =9372
                    LayoutCachedHeight =468
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7200
                    Top =276
                    Width =1392
                    Height =192
                    Name ="EmptyCell401"
                    GroupTable =1
                    LayoutCachedLeft =7200
                    LayoutCachedTop =276
                    LayoutCachedWidth =8592
                    LayoutCachedHeight =468
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7200
                    Top =540
                    Width =1392
                    Height =264
                    Name ="EmptyCell402"
                    GroupTable =1
                    LayoutCachedLeft =7200
                    LayoutCachedTop =540
                    LayoutCachedWidth =8592
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5652
                    Top =1536
                    Width =1476
                    Height =252
                    Name ="EmptyCell404"
                    GroupTable =1
                    LayoutCachedLeft =5652
                    LayoutCachedTop =1536
                    LayoutCachedWidth =7128
                    LayoutCachedHeight =1788
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
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
CodeBehindForm
' See "VendorContact_sub.cls"
