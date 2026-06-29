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
    Width =14328
    DatasheetFontHeight =11
    ItemSuffix =388
    Left =7260
    Top =7860
    Right =21360
    Bottom =17250
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x9767ae2a8cb6e540
    End
    RecordSource ="Select * From Location_contacts where ContactID =0;"
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
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =12690
                    Top =525
                    TabIndex =18
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
                    Left =8640
                    Top =2760
                    Width =720
                    TabIndex =15
                    Name ="contactOptOut"
                    ControlSource ="contactOptOut"
                    StatusBarText ="email opt out"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =8640
                    LayoutCachedTop =2760
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =3000
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
                            TextAlign =3
                            Left =7176
                            Top =2760
                            Width =1392
                            Height =240
                            FontSize =9
                            Name ="Label25"
                            Caption ="Do NOT Email"
                            GroupTable =1
                            LayoutCachedLeft =7176
                            LayoutCachedTop =2760
                            LayoutCachedWidth =8568
                            LayoutCachedHeight =3000
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
                    Left =3972
                    Top =1512
                    Width =1740
                    FontSize =8
                    TabIndex =8
                    BorderColor =11830108
                    Name ="contactPrimaryPhone"
                    ControlSource ="contactPrimaryPhone"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =1

                    LayoutCachedLeft =3972
                    LayoutCachedTop =1512
                    LayoutCachedWidth =5712
                    LayoutCachedHeight =1752
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
                            Top =1512
                            Width =912
                            Height =240
                            FontSize =9
                            Name ="Label27"
                            Caption ="Pri Phone"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =1512
                            LayoutCachedWidth =3900
                            LayoutCachedHeight =1752
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
                    Left =6444
                    Top =1512
                    Width =660
                    FontSize =8
                    TabIndex =9
                    BorderColor =11830108
                    Name ="contactPhoneExt"
                    ControlSource ="contactPhoneExt"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =6444
                    LayoutCachedTop =1512
                    LayoutCachedWidth =7104
                    LayoutCachedHeight =1752
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
                            Left =5772
                            Top =1512
                            Width =600
                            Height =240
                            FontSize =9
                            Name ="Label31"
                            Caption ="Ext:"
                            GroupTable =1
                            LayoutCachedLeft =5772
                            LayoutCachedTop =1512
                            LayoutCachedWidth =6372
                            LayoutCachedHeight =1752
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
                    Left =3972
                    Top =1824
                    Width =1740
                    FontSize =8
                    TabIndex =11
                    BorderColor =11830108
                    Name ="contactPhone2"
                    ControlSource ="contactPhone2"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =1

                    LayoutCachedLeft =3972
                    LayoutCachedTop =1824
                    LayoutCachedWidth =5712
                    LayoutCachedHeight =2064
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
                            Top =1824
                            Width =912
                            Height =240
                            FontSize =9
                            Name ="Label37"
                            Caption ="Alt Phone"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =1824
                            LayoutCachedWidth =3900
                            LayoutCachedHeight =2064
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
                    Left =3972
                    Top =2136
                    Width =1740
                    FontSize =8
                    TabIndex =12
                    BorderColor =11830108
                    Name ="contactCell"
                    ControlSource ="contactCell"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =1

                    LayoutCachedLeft =3972
                    LayoutCachedTop =2136
                    LayoutCachedWidth =5712
                    LayoutCachedHeight =2376
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
                            Top =2136
                            Width =912
                            Height =240
                            FontSize =9
                            Name ="Label38"
                            Caption ="Cellular"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =2136
                            LayoutCachedWidth =3900
                            LayoutCachedHeight =2376
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
                    TabIndex =24
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
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =3972
                    Top =3072
                    Width =2400
                    FontSize =9
                    TabIndex =16
                    BackColor =13434879
                    BorderColor =967423
                    Name ="contactClassification"
                    ControlSource ="contactClassification"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ContactType.ContactType FROM ContactType WHERE (((ContactType.ContactType"
                        ") Not Like \"*Incumbent*\" And (ContactType.ContactType) Not Like \"*Custodian*\""
                        "));"
                    ColumnWidths ="1440"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Classification is Required"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    AllowValueListEdits =0

                    LayoutCachedLeft =3972
                    LayoutCachedTop =3072
                    LayoutCachedWidth =6372
                    LayoutCachedHeight =3312
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
                            Left =2988
                            Top =3072
                            Width =912
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label3"
                            Caption ="Class¹ "
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =3072
                            LayoutCachedWidth =3900
                            LayoutCachedHeight =3312
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
                    Left =3972
                    Top =2448
                    Width =1740
                    FontSize =8
                    TabIndex =13
                    BorderColor =11830108
                    Name ="contactFax"
                    ControlSource ="contactFax"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =1

                    LayoutCachedLeft =3972
                    LayoutCachedTop =2448
                    LayoutCachedWidth =5712
                    LayoutCachedHeight =2688
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
                            Left =2988
                            Top =2448
                            Width =912
                            Height =240
                            FontSize =9
                            Name ="Label48"
                            Caption ="Fax"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =2448
                            LayoutCachedWidth =3900
                            LayoutCachedHeight =2688
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
                    Left =12288
                    Top =3720
                    Width =576
                    Height =195
                    FontSize =8
                    TabIndex =27
                    Name ="ClientId"
                    ControlSource ="ClientId"
                    StatusBarText ="link to Clients table"
                    DefaultValue ="3840"
                    GroupTable =11
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12288
                    LayoutCachedTop =3720
                    LayoutCachedWidth =12864
                    LayoutCachedHeight =3915
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
                            Left =11640
                            Top =3720
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
                            LayoutCachedLeft =11640
                            LayoutCachedTop =3720
                            LayoutCachedWidth =12274
                            LayoutCachedHeight =3915
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
                    Left =13680
                    Top =3720
                    Width =576
                    Height =195
                    FontSize =8
                    TabIndex =28
                    Name ="contactID"
                    ControlSource ="contactID"
                    GroupTable =11
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13680
                    LayoutCachedTop =3720
                    LayoutCachedWidth =14256
                    LayoutCachedHeight =3915
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
                            Left =12876
                            Top =3720
                            Width =792
                            Height =195
                            FontSize =8
                            Name ="Label43"
                            Caption ="contactID:"
                            GroupTable =11
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =12876
                            LayoutCachedTop =3720
                            LayoutCachedWidth =13668
                            LayoutCachedHeight =3915
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
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =6444
                    Top =3072
                    Width =2916
                    Height =240
                    FontSize =8
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label55"
                    Caption ="¹ Required or Must be Selected"
                    GroupTable =1
                    LayoutCachedLeft =6444
                    LayoutCachedTop =3072
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =3312
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
                    Left =3972
                    Top =2760
                    Width =3132
                    FontSize =8
                    TabIndex =14
                    BorderColor =11830108
                    Name ="contactEmail"
                    ControlSource ="contactEmail"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =3972
                    LayoutCachedTop =2760
                    LayoutCachedWidth =7104
                    LayoutCachedHeight =3000
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
                            Left =2988
                            Top =2760
                            Width =912
                            Height =240
                            FontSize =9
                            Name ="Label35"
                            Caption ="E-mail"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =2760
                            LayoutCachedWidth =3900
                            LayoutCachedHeight =3000
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
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =8
                    Left =3974
                    Top =3300
                    Width =10350
                    FontSize =8
                    TabIndex =26
                    BorderColor =11830108
                    Name ="Locref"
                    ControlSource ="cLocationID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LocationStreet, L"
                        "ocations.LAddrLine1, Locations.LAddrLine2, Locations.LCity, Locations.LState, Lo"
                        "cations.LZip FROM Locations WHERE (((Locations.ClientREF)=[Forms]![BuyersForm]!["
                        "BuyerID]) AND ((Locations.LocObsolete)=No));"
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
                            OverlapFlags =85
                            TextAlign =1
                            Left =3000
                            Top =3300
                            Width =885
                            Height =240
                            FontSize =9
                            Name ="Label58"
                            Caption ="Address"
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
                    Left =4836
                    Top =540
                    Width =1536
                    Height =276
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

                    LayoutCachedLeft =4836
                    LayoutCachedTop =540
                    LayoutCachedWidth =6372
                    LayoutCachedHeight =816
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
                    Left =6444
                    Top =540
                    Width =660
                    Height =276
                    FontSize =8
                    TabIndex =2
                    BorderColor =11830108
                    Name ="contactMiddle"
                    ControlSource ="contactMiddle"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =6444
                    LayoutCachedTop =540
                    LayoutCachedWidth =7104
                    LayoutCachedHeight =816
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
                    Left =7176
                    Top =540
                    Width =1392
                    Height =276
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

                    LayoutCachedLeft =7176
                    LayoutCachedTop =540
                    LayoutCachedWidth =8568
                    LayoutCachedHeight =816
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
                    Left =3972
                    Top =1200
                    Width =3132
                    FontSize =8
                    TabIndex =7
                    BorderColor =11830108
                    Name ="contactTitle"
                    ControlSource ="contactTitle"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =3972
                    LayoutCachedTop =1200
                    LayoutCachedWidth =7104
                    LayoutCachedHeight =1440
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
                            Width =912
                            Height =240
                            FontSize =9
                            Name ="Label71"
                            Caption ="Title"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =1200
                            LayoutCachedWidth =3900
                            LayoutCachedHeight =1440
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
                    Left =7176
                    Top =888
                    Width =1392
                    Height =240
                    FontSize =8
                    TabIndex =6
                    Name ="MakeFullContactBtn"
                    Caption ="◄ update"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7176
                    LayoutCachedTop =888
                    LayoutCachedWidth =8568
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
                    Left =3972
                    Top =888
                    Width =3132
                    FontSize =8
                    TabIndex =5
                    BackColor =-2147483615
                    BorderColor =15062992
                    ForeColor =16777215
                    Name ="contactFullName"
                    ControlSource ="contactFullName"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =3972
                    LayoutCachedTop =888
                    LayoutCachedWidth =7104
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
                            Top =888
                            Width =912
                            Height =240
                            FontSize =9
                            Name ="Label42"
                            Caption ="Full Name"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =888
                            LayoutCachedWidth =3900
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
                    TabIndex =22
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
                    Left =3972
                    Top =540
                    Width =792
                    Height =276
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

                    LayoutCachedLeft =3972
                    LayoutCachedTop =540
                    LayoutCachedWidth =4764
                    LayoutCachedHeight =816
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
                            Width =912
                            Height =276
                            FontSize =9
                            Name ="Label28"
                            Caption ="Name"
                            GroupTable =1
                            LayoutCachedLeft =2988
                            LayoutCachedTop =540
                            LayoutCachedWidth =3900
                            LayoutCachedHeight =816
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
                    Left =8640
                    Top =540
                    Width =720
                    Height =276
                    FontSize =8
                    TabIndex =4
                    BorderColor =11830108
                    Name ="contactSuffix"
                    ControlSource ="contactSuffix"
                    RowSourceType ="Value List"
                    FontName ="Tahoma"
                    GroupTable =1

                    LayoutCachedLeft =8640
                    LayoutCachedTop =540
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =816
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
                    TabIndex =25
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
                    TabIndex =19
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
                    HoverThemeColorIndex =7
                    QuickStyle =34
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =3972
                    Top =276
                    Width =792
                    Height =192
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label29"
                    Caption ="Salutation"
                    GroupTable =1
                    LayoutCachedLeft =3972
                    LayoutCachedTop =276
                    LayoutCachedWidth =4764
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
                    TextFontCharSet =238
                    TextAlign =1
                    Left =4836
                    Top =276
                    Width =1536
                    Height =192
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label30"
                    Caption ="First¹"
                    GroupTable =1
                    LayoutCachedLeft =4836
                    LayoutCachedTop =276
                    LayoutCachedWidth =6372
                    LayoutCachedHeight =468
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
                    Left =6444
                    Top =276
                    Width =660
                    Height =192
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label34"
                    Caption ="Middle"
                    GroupTable =1
                    LayoutCachedLeft =6444
                    LayoutCachedTop =276
                    LayoutCachedWidth =7104
                    LayoutCachedHeight =468
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
                    Left =7176
                    Top =276
                    Width =1392
                    Height =192
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label32"
                    Caption ="Last¹"
                    GroupTable =1
                    LayoutCachedLeft =7176
                    LayoutCachedTop =276
                    LayoutCachedWidth =8568
                    LayoutCachedHeight =468
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
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =75
                    Top =570
                    Width =2805
                    Height =5850
                    FontSize =8
                    TabIndex =20
                    Name ="ContactsList"
                    RowSourceType ="Table/Query"
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
                    Caption ="Contacts"
                    FontName ="Segoe UI"
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
                    TabIndex =21
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
                    TabIndex =17
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
                    TabIndex =29
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
                    HoverThemeColorIndex =9
                    HoverShade =75.0
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
                    TabIndex =23
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
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =8640
                    Top =276
                    Width =720
                    Height =192
                    FontSize =8
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label36"
                    Caption ="Suffix"
                    GroupTable =1
                    LayoutCachedLeft =8640
                    LayoutCachedTop =276
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =468
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2988
                    Top =276
                    Width =912
                    Height =192
                    Name ="EmptyCell121"
                    GroupTable =1
                    LayoutCachedLeft =2988
                    LayoutCachedTop =276
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =468
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8640
                    Top =888
                    Width =720
                    Name ="EmptyCell134"
                    GroupTable =1
                    LayoutCachedLeft =8640
                    LayoutCachedTop =888
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =1128
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7176
                    Top =1200
                    Width =1392
                    Name ="EmptyCell146"
                    GroupTable =1
                    LayoutCachedLeft =7176
                    LayoutCachedTop =1200
                    LayoutCachedWidth =8568
                    LayoutCachedHeight =1440
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8640
                    Top =1200
                    Width =720
                    Name ="EmptyCell147"
                    GroupTable =1
                    LayoutCachedLeft =8640
                    LayoutCachedTop =1200
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =1440
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5772
                    Top =1824
                    Width =600
                    Name ="EmptyCell171"
                    GroupTable =1
                    LayoutCachedLeft =5772
                    LayoutCachedTop =1824
                    LayoutCachedWidth =6372
                    LayoutCachedHeight =2064
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6444
                    Top =1824
                    Width =660
                    Name ="EmptyCell172"
                    GroupTable =1
                    LayoutCachedLeft =6444
                    LayoutCachedTop =1824
                    LayoutCachedWidth =7104
                    LayoutCachedHeight =2064
                    RowStart =5
                    RowEnd =5
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7176
                    Top =1824
                    Width =1392
                    Name ="EmptyCell173"
                    GroupTable =1
                    LayoutCachedLeft =7176
                    LayoutCachedTop =1824
                    LayoutCachedWidth =8568
                    LayoutCachedHeight =2064
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8640
                    Top =1824
                    Width =720
                    Name ="EmptyCell174"
                    GroupTable =1
                    LayoutCachedLeft =8640
                    LayoutCachedTop =1824
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =2064
                    RowStart =5
                    RowEnd =5
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5772
                    Top =2136
                    Width =600
                    Name ="EmptyCell194"
                    GroupTable =1
                    LayoutCachedLeft =5772
                    LayoutCachedTop =2136
                    LayoutCachedWidth =6372
                    LayoutCachedHeight =2376
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6444
                    Top =2136
                    Width =660
                    Name ="EmptyCell195"
                    GroupTable =1
                    LayoutCachedLeft =6444
                    LayoutCachedTop =2136
                    LayoutCachedWidth =7104
                    LayoutCachedHeight =2376
                    RowStart =6
                    RowEnd =6
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7176
                    Top =2136
                    Width =1392
                    Name ="EmptyCell196"
                    GroupTable =1
                    LayoutCachedLeft =7176
                    LayoutCachedTop =2136
                    LayoutCachedWidth =8568
                    LayoutCachedHeight =2376
                    RowStart =6
                    RowEnd =6
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8640
                    Top =2136
                    Width =720
                    Name ="EmptyCell197"
                    GroupTable =1
                    LayoutCachedLeft =8640
                    LayoutCachedTop =2136
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =2376
                    RowStart =6
                    RowEnd =6
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5772
                    Top =2448
                    Width =600
                    Name ="EmptyCell207"
                    GroupTable =1
                    LayoutCachedLeft =5772
                    LayoutCachedTop =2448
                    LayoutCachedWidth =6372
                    LayoutCachedHeight =2688
                    RowStart =7
                    RowEnd =7
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6444
                    Top =2448
                    Width =660
                    Name ="EmptyCell208"
                    GroupTable =1
                    LayoutCachedLeft =6444
                    LayoutCachedTop =2448
                    LayoutCachedWidth =7104
                    LayoutCachedHeight =2688
                    RowStart =7
                    RowEnd =7
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7176
                    Top =2448
                    Width =1392
                    Name ="EmptyCell209"
                    GroupTable =1
                    LayoutCachedLeft =7176
                    LayoutCachedTop =2448
                    LayoutCachedWidth =8568
                    LayoutCachedHeight =2688
                    RowStart =7
                    RowEnd =7
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8640
                    Top =2448
                    Width =720
                    Name ="EmptyCell210"
                    GroupTable =1
                    LayoutCachedLeft =8640
                    LayoutCachedTop =2448
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =2688
                    RowStart =7
                    RowEnd =7
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =8640
                    Top =1512
                    Width =720
                    TabIndex =10
                    Name ="DoNotCall"
                    ControlSource ="DoNotCall"
                    StatusBarText ="email opt out"
                    OnClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =8640
                    LayoutCachedTop =1512
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =1752
                    RowStart =4
                    RowEnd =4
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
                            TextAlign =3
                            Left =7176
                            Top =1512
                            Width =1392
                            Height =240
                            FontSize =9
                            Name ="Label365"
                            Caption ="Do NOT Call"
                            GroupTable =1
                            LayoutCachedLeft =7176
                            LayoutCachedTop =1512
                            LayoutCachedWidth =8568
                            LayoutCachedHeight =1752
                            RowStart =4
                            RowEnd =4
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
' See "RemarketingContact_sub.cls"
