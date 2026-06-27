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
    ItemSuffix =364
    Left =6930
    Top =7560
    Right =22320
    Bottom =14730
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x8207f34c7084e440
    End
    RecordSource ="Location_contacts"
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
                    TabIndex =42
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
                    TabIndex =17
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
                    Left =8550
                    Top =2655
                    Width =720
                    TabIndex =14
                    Name ="contactOptOut"
                    ControlSource ="contactOptOut"
                    StatusBarText ="email opt out"
                    GroupTable =1

                    LayoutCachedLeft =8550
                    LayoutCachedTop =2655
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =2895
                    RowStart =8
                    RowEnd =8
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
                            Left =7095
                            Top =2655
                            Width =1395
                            Height =240
                            FontSize =9
                            Name ="Label25"
                            Caption ="Opt Out of Email"
                            GroupTable =1
                            LayoutCachedLeft =7095
                            LayoutCachedTop =2655
                            LayoutCachedWidth =8490
                            LayoutCachedHeight =2895
                            RowStart =8
                            RowEnd =8
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
                    Left =3930
                    Top =1455
                    Width =1725
                    FontSize =8
                    TabIndex =8
                    BorderColor =11830108
                    Name ="contactPrimaryPhone"
                    ControlSource ="contactPrimaryPhone"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =1

                    LayoutCachedLeft =3930
                    LayoutCachedTop =1455
                    LayoutCachedWidth =5655
                    LayoutCachedHeight =1695
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
                            Left =2985
                            Top =1455
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label27"
                            Caption ="Pri Phone:"
                            GroupTable =1
                            LayoutCachedLeft =2985
                            LayoutCachedTop =1455
                            LayoutCachedWidth =3870
                            LayoutCachedHeight =1695
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
                    Left =6375
                    Top =1455
                    Width =660
                    FontSize =8
                    TabIndex =9
                    BorderColor =11830108
                    Name ="contactPhoneExt"
                    ControlSource ="contactPhoneExt"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =6375
                    LayoutCachedTop =1455
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =1695
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =5715
                            Top =1455
                            Width =600
                            Height =240
                            FontSize =9
                            Name ="Label31"
                            Caption ="Ext:"
                            GroupTable =1
                            LayoutCachedLeft =5715
                            LayoutCachedTop =1455
                            LayoutCachedWidth =6315
                            LayoutCachedHeight =1695
                            RowStart =4
                            RowEnd =4
                            ColumnStart =3
                            ColumnEnd =3
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
                    Left =3930
                    Top =1755
                    Width =1725
                    FontSize =8
                    TabIndex =10
                    BorderColor =11830108
                    Name ="contactPhone2"
                    ControlSource ="contactPhone2"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =1

                    LayoutCachedLeft =3930
                    LayoutCachedTop =1755
                    LayoutCachedWidth =5655
                    LayoutCachedHeight =1995
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
                            Left =2985
                            Top =1755
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label37"
                            Caption ="Alt Phone:"
                            GroupTable =1
                            LayoutCachedLeft =2985
                            LayoutCachedTop =1755
                            LayoutCachedWidth =3870
                            LayoutCachedHeight =1995
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
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3930
                    Top =2055
                    Width =1725
                    FontSize =8
                    TabIndex =11
                    BorderColor =11830108
                    Name ="contactCell"
                    ControlSource ="contactCell"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =1

                    LayoutCachedLeft =3930
                    LayoutCachedTop =2055
                    LayoutCachedWidth =5655
                    LayoutCachedHeight =2295
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
                            Left =2985
                            Top =2055
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label38"
                            Caption ="Cellular:"
                            GroupTable =1
                            LayoutCachedLeft =2985
                            LayoutCachedTop =2055
                            LayoutCachedWidth =3870
                            LayoutCachedHeight =2295
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
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =9570
                    Top =1605
                    Width =4740
                    Height =1653
                    FontSize =8
                    TabIndex =23
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
                            TextFontCharSet =238
                            TextAlign =3
                            Left =9571
                            Top =1410
                            Width =480
                            Height =255
                            FontSize =8
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
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListRows =8
                    ListWidth =1440
                    Left =3930
                    Top =2955
                    Width =2385
                    FontSize =9
                    TabIndex =15
                    BackColor =13434879
                    BorderColor =967423
                    Name ="contactClassification"
                    ControlSource ="contactClassification"
                    RowSourceType ="Value List"
                    RowSource ="BILLING;FINANCIAL;LEGAL;OPERATIONS;PRIMARY;PEOPLENET;PROCUREMENT;SIGNATOR;TOP MA"
                        "NAGEMENT"
                    ColumnWidths ="1440"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Classification is Required"
                    GroupTable =1
                    AllowValueListEdits =0

                    LayoutCachedLeft =3930
                    LayoutCachedTop =2955
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =3195
                    RowStart =9
                    RowEnd =9
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =2985
                            Top =2955
                            Width =885
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label3"
                            Caption ="Class¹ :"
                            GroupTable =1
                            LayoutCachedLeft =2985
                            LayoutCachedTop =2955
                            LayoutCachedWidth =3870
                            LayoutCachedHeight =3195
                            RowStart =9
                            RowEnd =9
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
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
                    Left =3930
                    Top =2355
                    Width =1725
                    FontSize =8
                    TabIndex =12
                    BorderColor =11830108
                    Name ="contactFax"
                    ControlSource ="contactFax"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =1

                    LayoutCachedLeft =3930
                    LayoutCachedTop =2355
                    LayoutCachedWidth =5655
                    LayoutCachedHeight =2595
                    RowStart =7
                    RowEnd =7
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
                            Left =2985
                            Top =2355
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label48"
                            Caption ="Fax:"
                            GroupTable =1
                            LayoutCachedLeft =2985
                            LayoutCachedTop =2355
                            LayoutCachedWidth =3870
                            LayoutCachedHeight =2595
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
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12300
                    Top =5985
                    Width =576
                    Height =195
                    FontSize =8
                    TabIndex =34
                    Name ="ClientId"
                    ControlSource ="ClientId"
                    StatusBarText ="link to Clients table"
                    DefaultValue ="3840"
                    GroupTable =11
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12300
                    LayoutCachedTop =5985
                    LayoutCachedWidth =12876
                    LayoutCachedHeight =6180
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
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
                            Left =11670
                            Top =5985
                            Width =634
                            Height =195
                            FontSize =8
                            Name ="Label26"
                            Caption ="clientid:"
                            GroupTable =11
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =11670
                            LayoutCachedTop =5985
                            LayoutCachedWidth =12304
                            LayoutCachedHeight =6180
                            LayoutGroup =4
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
                    Left =13665
                    Top =5985
                    Width =576
                    Height =195
                    FontSize =8
                    TabIndex =35
                    Name ="contactID"
                    ControlSource ="contactID"
                    GroupTable =11
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13665
                    LayoutCachedTop =5985
                    LayoutCachedWidth =14241
                    LayoutCachedHeight =6180
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
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
                            Left =12870
                            Top =5985
                            Width =795
                            Height =195
                            FontSize =8
                            Name ="Label43"
                            Caption ="contactID:"
                            GroupTable =11
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =12870
                            LayoutCachedTop =5985
                            LayoutCachedWidth =13665
                            LayoutCachedHeight =6180
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =4
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
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =6375
                    Top =2955
                    Width =2895
                    Height =240
                    FontSize =8
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label55"
                    Caption ="¹ Required or Must be Selected"
                    GroupTable =1
                    LayoutCachedLeft =6375
                    LayoutCachedTop =2955
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =3195
                    RowStart =9
                    RowEnd =9
                    ColumnStart =4
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3930
                    Top =2655
                    Width =3105
                    FontSize =8
                    TabIndex =13
                    BorderColor =11830108
                    Name ="contactEmail"
                    ControlSource ="contactEmail"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =3930
                    LayoutCachedTop =2655
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =2895
                    RowStart =8
                    RowEnd =8
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
                            Left =2985
                            Top =2655
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label35"
                            Caption ="E-mail:"
                            GroupTable =1
                            LayoutCachedLeft =2985
                            LayoutCachedTop =2655
                            LayoutCachedWidth =3870
                            LayoutCachedHeight =2895
                            RowStart =8
                            RowEnd =8
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
                    FontSize =8
                    TabIndex =25
                    BorderColor =11830108
                    Name ="Locref"
                    ControlSource ="cLocationID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LocationStreet, L"
                        "ocations.LAddrLine1, Locations.LAddrLine2, Locations.LCity, Locations.LState, Lo"
                        "cations.LZip FROM Locations WHERE (((Locations.ClientREF)=[forms]![Bank_view]![c"
                        "lientid]) AND ((Locations.LocObsolete)=No));"
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
                    Left =4785
                    Top =525
                    Width =1530
                    Height =270
                    FontSize =8
                    TabIndex =1
                    BorderColor =11830108
                    Name ="contactFirst"
                    ControlSource ="contactFirst"
                    ValidationRule ="Is Not Null"
                    ValidationText ="First Name is required"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =4785
                    LayoutCachedTop =525
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =795
                    DisplayAsHyperlink =2
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =3
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
                    Left =6375
                    Top =525
                    Width =660
                    Height =270
                    FontSize =8
                    TabIndex =2
                    BorderColor =11830108
                    Name ="contactMiddle"
                    ControlSource ="contactMiddle"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =6375
                    LayoutCachedTop =525
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =795
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
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
                    Left =7095
                    Top =525
                    Width =1395
                    Height =270
                    FontSize =8
                    TabIndex =3
                    BorderColor =11830108
                    Name ="contactLast"
                    ControlSource ="contactLast"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Last Name is required."
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =7095
                    LayoutCachedTop =525
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =795
                    DisplayAsHyperlink =2
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
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
                    Left =3930
                    Top =1155
                    Width =3105
                    FontSize =8
                    TabIndex =7
                    BorderColor =11830108
                    Name ="contactTitle"
                    ControlSource ="contactTitle"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =3930
                    LayoutCachedTop =1155
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =1395
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
                            Left =2985
                            Top =1155
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label71"
                            Caption ="Title:"
                            GroupTable =1
                            LayoutCachedLeft =2985
                            LayoutCachedTop =1155
                            LayoutCachedWidth =3870
                            LayoutCachedHeight =1395
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
                    Left =7095
                    Top =855
                    Width =1395
                    Height =240
                    FontSize =8
                    TabIndex =6
                    Name ="Command41"
                    Caption ="◄ update"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7095
                    LayoutCachedTop =855
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =1095
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
                    Left =3930
                    Top =855
                    Width =3105
                    FontSize =8
                    TabIndex =5
                    BackColor =-2147483615
                    BorderColor =15062992
                    ForeColor =16777215
                    Name ="contactFullName"
                    ControlSource ="contactFullName"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =3930
                    LayoutCachedTop =855
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =1095
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
                            Left =2985
                            Top =855
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label42"
                            Caption ="Full Name:"
                            GroupTable =1
                            LayoutCachedLeft =2985
                            LayoutCachedTop =855
                            LayoutCachedWidth =3870
                            LayoutCachedHeight =1095
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
                    TabIndex =21
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
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3930
                    Top =525
                    Width =795
                    Height =270
                    FontSize =8
                    BorderColor =11830108
                    Name ="contactPrefix"
                    ControlSource ="contactPrefix"
                    RowSourceType ="Value List"
                    RowSource ="Dr.;Mr.;Mrs.;Ms"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    GroupTable =1
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =3930
                    LayoutCachedTop =525
                    LayoutCachedWidth =4725
                    LayoutCachedHeight =795
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
                            Left =2985
                            Top =525
                            Width =885
                            Height =270
                            FontSize =9
                            Name ="Label28"
                            Caption ="Name:"
                            GroupTable =1
                            LayoutCachedLeft =2985
                            LayoutCachedTop =525
                            LayoutCachedWidth =3870
                            LayoutCachedHeight =795
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
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8550
                    Top =525
                    Width =720
                    Height =270
                    FontSize =8
                    TabIndex =4
                    BorderColor =11830108
                    Name ="contactSuffix"
                    ControlSource ="contactSuffix"
                    RowSourceType ="Value List"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =8550
                    LayoutCachedTop =525
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =795
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =163
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7800
                    Top =2115
                    Width =1125
                    FontSize =9
                    TabIndex =24
                    Name ="LastChangeDt"
                    ControlSource ="LastChangeDt"
                    StatusBarText ="updated when contact unlocked to make change"

                    LayoutCachedLeft =7800
                    LayoutCachedTop =2115
                    LayoutCachedWidth =8925
                    LayoutCachedHeight =2355
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =163
                            TextAlign =2
                            Left =7770
                            Top =1875
                            Width =1215
                            Height =555
                            FontSize =9
                            BackColor =11525325
                            Name ="Label69"
                            Caption ="Last Updated:"
                            LayoutCachedLeft =7770
                            LayoutCachedTop =1875
                            LayoutCachedWidth =8985
                            LayoutCachedHeight =2430
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =11655
                    Top =555
                    Width =720
                    Height =660
                    FontSize =10
                    FontWeight =700
                    TabIndex =18
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
                    QuickStyle =34
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =3930
                    Top =270
                    Width =795
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label29"
                    Caption ="Salutation"
                    GroupTable =1
                    LayoutCachedLeft =3930
                    LayoutCachedTop =270
                    LayoutCachedWidth =4725
                    LayoutCachedHeight =465
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =4785
                    Top =270
                    Width =1530
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label30"
                    Caption ="First¹"
                    GroupTable =1
                    LayoutCachedLeft =4785
                    LayoutCachedTop =270
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =465
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =6375
                    Top =270
                    Width =660
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label34"
                    Caption ="Middle"
                    GroupTable =1
                    LayoutCachedLeft =6375
                    LayoutCachedTop =270
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =465
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =7095
                    Top =270
                    Width =1395
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label32"
                    Caption ="Last¹"
                    GroupTable =1
                    LayoutCachedLeft =7095
                    LayoutCachedTop =270
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =465
                    ColumnStart =5
                    ColumnEnd =5
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
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =75
                    Top =570
                    Width =2805
                    Height =5850
                    FontSize =8
                    TabIndex =19
                    Name ="ContactsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT contactID, clientRef, contactClassification, contactFullName AS Contact F"
                        "ROM Location_contacts WHERE clientRef=5167 Order by contactFullName ;"
                    ColumnWidths ="0;0;1008;1728"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =75
                    LayoutCachedTop =570
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =6420
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =30
                    Width =2835
                    Height =465
                    FontSize =18
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label80"
                    Caption ="Bank Contacts"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =60
                    LayoutCachedTop =30
                    LayoutCachedWidth =2895
                    LayoutCachedHeight =495
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9900
                    Top =705
                    Width =1395
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =20
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
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =9900
                    Top =375
                    Width =1395
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =16
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
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =3075
                    Top =6165
                    Width =1725
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =41
                    Name ="DeleteContact"
                    Caption ="DELETE  This Contact"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3075
                    LayoutCachedTop =6165
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =6405
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =2366701
                    BorderThemeColorIndex =9
                    BorderShade =75.0
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
                    FontSize =8
                    FontWeight =700
                    TabIndex =22
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
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =8550
                    Top =270
                    Width =720
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label36"
                    Caption ="Suffix"
                    GroupTable =1
                    LayoutCachedLeft =8550
                    LayoutCachedTop =270
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =465
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2985
                    Top =270
                    Width =885
                    Height =195
                    Name ="EmptyCell121"
                    GroupTable =1
                    LayoutCachedLeft =2985
                    LayoutCachedTop =270
                    LayoutCachedWidth =3870
                    LayoutCachedHeight =465
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8550
                    Top =855
                    Width =720
                    Name ="EmptyCell134"
                    GroupTable =1
                    LayoutCachedLeft =8550
                    LayoutCachedTop =855
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =1095
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7095
                    Top =1155
                    Width =1395
                    Name ="EmptyCell146"
                    GroupTable =1
                    LayoutCachedLeft =7095
                    LayoutCachedTop =1155
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =1395
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8550
                    Top =1155
                    Width =720
                    Name ="EmptyCell147"
                    GroupTable =1
                    LayoutCachedLeft =8550
                    LayoutCachedTop =1155
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =1395
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7095
                    Top =1455
                    Width =1395
                    Name ="EmptyCell156"
                    GroupTable =1
                    LayoutCachedLeft =7095
                    LayoutCachedTop =1455
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =1695
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8550
                    Top =1455
                    Width =720
                    Name ="EmptyCell157"
                    GroupTable =1
                    LayoutCachedLeft =8550
                    LayoutCachedTop =1455
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =1695
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5715
                    Top =1755
                    Width =600
                    Name ="EmptyCell171"
                    GroupTable =1
                    LayoutCachedLeft =5715
                    LayoutCachedTop =1755
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =1995
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6375
                    Top =1755
                    Width =660
                    Name ="EmptyCell172"
                    GroupTable =1
                    LayoutCachedLeft =6375
                    LayoutCachedTop =1755
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =1995
                    RowStart =5
                    RowEnd =5
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7095
                    Top =1755
                    Width =1395
                    Name ="EmptyCell173"
                    GroupTable =1
                    LayoutCachedLeft =7095
                    LayoutCachedTop =1755
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =1995
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8550
                    Top =1755
                    Width =720
                    Name ="EmptyCell174"
                    GroupTable =1
                    LayoutCachedLeft =8550
                    LayoutCachedTop =1755
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =1995
                    RowStart =5
                    RowEnd =5
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5715
                    Top =2055
                    Width =600
                    Name ="EmptyCell194"
                    GroupTable =1
                    LayoutCachedLeft =5715
                    LayoutCachedTop =2055
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =2295
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6375
                    Top =2055
                    Width =660
                    Name ="EmptyCell195"
                    GroupTable =1
                    LayoutCachedLeft =6375
                    LayoutCachedTop =2055
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =2295
                    RowStart =6
                    RowEnd =6
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7095
                    Top =2055
                    Width =1395
                    Name ="EmptyCell196"
                    GroupTable =1
                    LayoutCachedLeft =7095
                    LayoutCachedTop =2055
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =2295
                    RowStart =6
                    RowEnd =6
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8550
                    Top =2055
                    Width =720
                    Name ="EmptyCell197"
                    GroupTable =1
                    LayoutCachedLeft =8550
                    LayoutCachedTop =2055
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =2295
                    RowStart =6
                    RowEnd =6
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5715
                    Top =2355
                    Width =600
                    Name ="EmptyCell207"
                    GroupTable =1
                    LayoutCachedLeft =5715
                    LayoutCachedTop =2355
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =2595
                    RowStart =7
                    RowEnd =7
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6375
                    Top =2355
                    Width =660
                    Name ="EmptyCell208"
                    GroupTable =1
                    LayoutCachedLeft =6375
                    LayoutCachedTop =2355
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =2595
                    RowStart =7
                    RowEnd =7
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7095
                    Top =2355
                    Width =1395
                    Name ="EmptyCell209"
                    GroupTable =1
                    LayoutCachedLeft =7095
                    LayoutCachedTop =2355
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =2595
                    RowStart =7
                    RowEnd =7
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8550
                    Top =2355
                    Width =720
                    Name ="EmptyCell210"
                    GroupTable =1
                    LayoutCachedLeft =8550
                    LayoutCachedTop =2355
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =2595
                    RowStart =7
                    RowEnd =7
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10305
                    Top =3960
                    Width =3300
                    Height =270
                    FontSize =10
                    TabIndex =26
                    BorderColor =11830108
                    Name ="Locationsn"
                    ControlSource ="Locations.Locationsn"
                    StatusBarText ="short name for locaition or Lessee number"
                    GroupTable =8

                    LayoutCachedLeft =10305
                    LayoutCachedTop =3960
                    LayoutCachedWidth =13605
                    LayoutCachedHeight =4230
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =9435
                            Top =3960
                            Width =810
                            Height =270
                            FontSize =9
                            Name ="Label90"
                            Caption ="Location:"
                            GroupTable =8
                            LayoutCachedLeft =9435
                            LayoutCachedTop =3960
                            LayoutCachedWidth =10245
                            LayoutCachedHeight =4230
                            LayoutGroup =2
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =8
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10305
                    Top =4290
                    Width =3300
                    FontSize =10
                    TabIndex =27
                    BorderColor =11830108
                    Name ="LocationStreet"
                    ControlSource ="LocationStreet"
                    StatusBarText ="default company name"
                    GroupTable =8

                    LayoutCachedLeft =10305
                    LayoutCachedTop =4290
                    LayoutCachedWidth =13605
                    LayoutCachedHeight =4530
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =9435
                            Top =4290
                            Width =810
                            Height =240
                            FontSize =9
                            Name ="Label91"
                            Caption ="Street:"
                            GroupTable =8
                            LayoutCachedLeft =9435
                            LayoutCachedTop =4290
                            LayoutCachedWidth =10245
                            LayoutCachedHeight =4530
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =8
                        End
                    End
                End
                Begin EmptyCell
                    Left =9435
                    Top =4590
                    Width =810
                    Name ="EmptyCell293"
                    GroupTable =8
                    LayoutCachedLeft =9435
                    LayoutCachedTop =4590
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =4830
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10305
                    Top =4590
                    Width =3300
                    FontSize =10
                    TabIndex =28
                    BorderColor =11830108
                    Name ="LAddrLine1"
                    ControlSource ="LAddrLine1"
                    GroupTable =8

                    LayoutCachedLeft =10305
                    LayoutCachedTop =4590
                    LayoutCachedWidth =13605
                    LayoutCachedHeight =4830
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =9435
                    Top =4890
                    Width =810
                    Name ="EmptyCell297"
                    GroupTable =8
                    LayoutCachedLeft =9435
                    LayoutCachedTop =4890
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =5130
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10305
                    Top =4890
                    Width =3300
                    FontSize =10
                    TabIndex =29
                    BorderColor =11830108
                    Name ="LAddrLine2"
                    ControlSource ="LAddrLine2"
                    GroupTable =8

                    LayoutCachedLeft =10305
                    LayoutCachedTop =4890
                    LayoutCachedWidth =13605
                    LayoutCachedHeight =5130
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =9435
                    Top =5190
                    Width =810
                    Name ="EmptyCell301"
                    GroupTable =8
                    LayoutCachedLeft =9435
                    LayoutCachedTop =5190
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =5430
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10305
                    Top =5190
                    Width =3300
                    FontSize =10
                    TabIndex =30
                    BorderColor =11830108
                    Name ="LAddrLine3"
                    ControlSource ="LAddrLine3"
                    GroupTable =8

                    LayoutCachedLeft =10305
                    LayoutCachedTop =5190
                    LayoutCachedWidth =13605
                    LayoutCachedHeight =5430
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =9435
                    Top =5490
                    Width =810
                    Name ="EmptyCell305"
                    GroupTable =8
                    LayoutCachedLeft =9435
                    LayoutCachedTop =5490
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =5730
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10305
                    Top =5490
                    Width =1980
                    FontSize =10
                    TabIndex =31
                    BorderColor =11830108
                    Name ="LCity"
                    ControlSource ="LCity"
                    GroupTable =8

                    LayoutCachedLeft =10305
                    LayoutCachedTop =5490
                    LayoutCachedWidth =12285
                    LayoutCachedHeight =5730
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12345
                    Top =5490
                    Width =465
                    FontSize =10
                    TabIndex =32
                    BorderColor =11830108
                    Name ="LState"
                    ControlSource ="StateAbbrev"
                    StatusBarText ="lookup in lookup_state table"
                    GroupTable =8

                    LayoutCachedLeft =12345
                    LayoutCachedTop =5490
                    LayoutCachedWidth =12810
                    LayoutCachedHeight =5730
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12870
                    Top =5490
                    Width =735
                    FontSize =10
                    TabIndex =33
                    BorderColor =11830108
                    Name ="LZip"
                    ControlSource ="LZip"
                    InputMask ="00000\\-9999;;_"
                    GroupTable =8

                    LayoutCachedLeft =12870
                    LayoutCachedTop =5490
                    LayoutCachedWidth =13605
                    LayoutCachedHeight =5730
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5610
                    Top =4830
                    Width =3330
                    Height =648
                    FontSize =8
                    TabIndex =40
                    BackColor =15858167
                    Name ="ContactNt"
                    ControlSource ="ContactNt"
                    FontName ="Tahoma"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5610
                    LayoutCachedTop =4830
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =5478
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =3
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin Label
                    OverlapFlags =255
                    Left =2985
                    Top =3645
                    Width =2055
                    Height =345
                    FontSize =12
                    FontWeight =700
                    Name ="Label73"
                    Caption ="Contact's Assistant"
                    LayoutCachedLeft =2985
                    LayoutCachedTop =3645
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =3990
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3990
                    Top =3975
                    Width =1590
                    Height =255
                    FontSize =8
                    TabIndex =36
                    Name ="ContactAsstFnm"
                    ControlSource ="ContactAsstFnm"
                    StatusBarText ="first name of contact's assistant"
                    FontName ="Tahoma"
                    GroupTable =9

                    LayoutCachedLeft =3990
                    LayoutCachedTop =3975
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =4230
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =2985
                            Top =3975
                            Width =945
                            Height =255
                            FontSize =9
                            Name ="Label74"
                            Caption ="Frist Name:"
                            GroupTable =9
                            LayoutCachedLeft =2985
                            LayoutCachedTop =3975
                            LayoutCachedWidth =3930
                            LayoutCachedHeight =4230
                            LayoutGroup =3
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =9
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6645
                    Top =3975
                    Width =2295
                    Height =255
                    FontSize =8
                    TabIndex =37
                    Name ="ContactAsstLnm"
                    ControlSource ="ContactAsstLnm"
                    StatusBarText ="Last name of contact's assistant"
                    FontName ="Tahoma"
                    GroupTable =9
                    RightPadding =0

                    LayoutCachedLeft =6645
                    LayoutCachedTop =3975
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =4230
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =5610
                            Top =3975
                            Width =1005
                            Height =255
                            FontSize =9
                            Name ="Label75"
                            Caption ="Last Name:"
                            GroupTable =9
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =5610
                            LayoutCachedTop =3975
                            LayoutCachedWidth =6615
                            LayoutCachedHeight =4230
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =3
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =9
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3990
                    Top =4290
                    Width =4950
                    Height =255
                    FontSize =8
                    TabIndex =38
                    Name ="ContactAsstE-mail"
                    ControlSource ="ContactAsstE-mail"
                    StatusBarText ="contact's assistants e-mail address"
                    FontName ="Tahoma"
                    EventProcPrefix ="ContactAsstE_mail"
                    GroupTable =9
                    RightPadding =0

                    LayoutCachedLeft =3990
                    LayoutCachedTop =4290
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =4545
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =3
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2985
                            Top =4290
                            Width =945
                            Height =255
                            FontSize =9
                            Name ="Label76"
                            Caption ="E-mail:"
                            GroupTable =9
                            LayoutCachedLeft =2985
                            LayoutCachedTop =4290
                            LayoutCachedWidth =3930
                            LayoutCachedHeight =4545
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =3
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =9
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3990
                    Top =4575
                    Width =1590
                    Height =255
                    FontSize =8
                    TabIndex =39
                    Name ="ContactAsstPH"
                    ControlSource ="ContactAsstPH"
                    StatusBarText ="contact assistants phone"
                    FontName ="Tahoma"
                    GroupTable =9
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3990
                    LayoutCachedTop =4575
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =4830
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2985
                            Top =4575
                            Width =945
                            Height =255
                            FontSize =9
                            Name ="Label77"
                            Caption ="Phone:"
                            GroupTable =9
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =2985
                            LayoutCachedTop =4575
                            LayoutCachedWidth =3930
                            LayoutCachedHeight =4830
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =3
                            ForeThemeColorIndex =6
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =9
                        End
                    End
                End
                Begin EmptyCell
                    Left =6645
                    Top =4575
                    Width =2295
                    Height =255
                    Name ="EmptyCell275"
                    GroupTable =9
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6645
                    LayoutCachedTop =4575
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =4830
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GroupTable =9
                End
                Begin EmptyCell
                    Left =2985
                    Top =4830
                    Width =945
                    Height =648
                    Name ="EmptyCell333"
                    GroupTable =9
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2985
                    LayoutCachedTop =4830
                    LayoutCachedWidth =3930
                    LayoutCachedHeight =5478
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =3
                    GroupTable =9
                End
                Begin EmptyCell
                    Left =3990
                    Top =4830
                    Width =1590
                    Height =648
                    Name ="EmptyCell337"
                    GroupTable =9
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3990
                    LayoutCachedTop =4830
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =5478
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GroupTable =9
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5610
                    Top =4575
                    Width =1005
                    Height =255
                    FontSize =8
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label346"
                    Caption ="Notes:"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5610
                    LayoutCachedTop =4575
                    LayoutCachedWidth =6615
                    LayoutCachedHeight =4830
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
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
' See "TempPlateContact_sub.cls"
