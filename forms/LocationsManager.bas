Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =19152
    DatasheetFontHeight =11
    ItemSuffix =733
    Left =4350
    Top =3555
    Right =22320
    Bottom =10470
    Filter ="LocationID = 1559"
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x397d78c81c07e540
    End
    RecordSource ="SELECT Locations.*, lookup_States.StTaxType FROM Locations LEFT JOIN lookup_Stat"
        "es ON Locations.LState = lookup_States.StateAbbrev WHERE LocationId =2341 ORDER "
        "BY Locations.LocationType, Locations.LState, Locations.LCity;"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    NavigationCaption ="Contacts"
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16777215
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
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
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
            CanGrow = NotDefault
            Height =6930
            Name ="Detail"
            AlternateBackColor =14545355
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =13020
                    Top =285
                    Width =3495
                    Height =225
                    FontSize =8
                    Name ="Label88"
                    Caption =" ┌─ Search   ─┬────────┬───────┐"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =13020
                    LayoutCachedTop =285
                    LayoutCachedWidth =16515
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListRows =8
                    ListWidth =1440
                    Left =5835
                    Top =450
                    Width =3570
                    ColumnWidth =1650
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =7
                    Name ="Location Type"
                    ControlSource ="LocationType"
                    RowSourceType ="Value List"
                    RowSource ="\"Billing\";\"Garage\";\"Headquarters\";\"Registration\";\"Title\";\"Other\";Con"
                        "signment"
                    ColumnWidths ="1440"
                    EventProcPrefix ="Location_Type"
                    AllowValueListEdits =0

                    LayoutCachedLeft =5835
                    LayoutCachedTop =450
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =690
                    ColumnStart =1
                    ColumnEnd =4
                    ThemeFontIndex =1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =450
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label3"
                            Caption ="Classification"
                            LayoutCachedLeft =4635
                            LayoutCachedTop =450
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =690
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5835
                    Top =750
                    Width =3570
                    ColumnWidth =1800
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =9
                    Name ="LocationName"
                    ControlSource ="LocationName"

                    LayoutCachedLeft =5835
                    LayoutCachedTop =750
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =990
                    DatasheetCaption ="Location Name"
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =4
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =750
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label39"
                            Caption ="Location Name"
                            LayoutCachedLeft =4635
                            LayoutCachedTop =750
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =990
                            RowStart =1
                            RowEnd =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10425
                    Top =750
                    Width =2070
                    ColumnWidth =1920
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =10
                    Name ="Locationsn"
                    ControlSource ="Locationsn"
                    StatusBarText ="short name for locaition or Lessee number"

                    LayoutCachedLeft =10425
                    LayoutCachedTop =750
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =990
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =7
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =9465
                            Top =750
                            Width =900
                            Height =240
                            FontSize =8
                            Name ="Label40"
                            Caption ="Short Name"
                            LayoutCachedLeft =9465
                            LayoutCachedTop =750
                            LayoutCachedWidth =10365
                            LayoutCachedHeight =990
                            RowStart =1
                            RowEnd =1
                            ColumnStart =5
                            ColumnEnd =5
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5835
                    Top =1050
                    Width =3570
                    ColumnWidth =2985
                    ColumnOrder =15
                    FontSize =9
                    TabIndex =11
                    Name ="LocationStreet"
                    ControlSource ="LocationStreet"

                    LayoutCachedLeft =5835
                    LayoutCachedTop =1050
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =4
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =1050
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label41"
                            Caption ="Address 1"
                            LayoutCachedLeft =4635
                            LayoutCachedTop =1050
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =1290
                            RowStart =2
                            RowEnd =2
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =20
                    ListWidth =2880
                    Left =9465
                    Top =1950
                    Width =900
                    Height =270
                    ColumnWidth =855
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =15
                    Name ="LState"
                    ControlSource ="LState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TaxStates.PostalAbbreviation, IIf([RentReceiptAvailable]=True,\"RR\",\"--"
                        "\") & \"/\" & IIf([UpfrontAvailable]=True,\"UpFront\",\"--\") AS [Option], TaxSt"
                        "ates.ISOCountryCode2, TaxStates.ID FROM TaxStates GROUP BY TaxStates.PostalAbbre"
                        "viation, IIf([RentReceiptAvailable]=True,\"RR\",\"--\") & \"/\" & IIf([UpfrontAv"
                        "ailable]=True,\"UpFront\",\"--\"), TaxStates.ISOCountryCode2, TaxStates.ID, TaxS"
                        "tates.PostalAbbreviation HAVING (((TaxStates.ISOCountryCode2)=\"US\" Or (TaxStat"
                        "es.ISOCountryCode2)=\"CA\" Or (TaxStates.ISOCountryCode2)=\"MX\")) ORDER BY TaxS"
                        "tates.ISOCountryCode2 DESC , TaxStates.PostalAbbreviation;"
                    ColumnWidths ="432;1152;432;0"
                    StatusBarText ="lookup in lookup_state table"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =9465
                    LayoutCachedTop =1950
                    LayoutCachedWidth =10365
                    LayoutCachedHeight =2220
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =5
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10425
                    Top =1950
                    Width =1035
                    Height =270
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =16
                    Name ="LZip"
                    ControlSource ="LZip"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="00000\\-9999;;_"

                    LayoutCachedLeft =10425
                    LayoutCachedTop =1950
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =2220
                    RowStart =5
                    RowEnd =5
                    ColumnStart =6
                    ColumnEnd =6
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5835
                    Top =2610
                    Width =1785
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =18
                    Name ="LPhone"
                    ControlSource ="LPhone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =5835
                    LayoutCachedTop =2610
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =2850
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =2
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =2610
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label46"
                            Caption ="Phone "
                            LayoutCachedLeft =4635
                            LayoutCachedTop =2610
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =2850
                            RowStart =7
                            RowEnd =7
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5835
                    Top =2910
                    Width =1785
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =20
                    Name ="LFax"
                    ControlSource ="LFax"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =5835
                    LayoutCachedTop =2910
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =3150
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =2
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =2910
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label50"
                            Caption ="Fax "
                            LayoutCachedLeft =4635
                            LayoutCachedTop =2910
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =3150
                            RowStart =8
                            RowEnd =8
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10350
                    Top =2925
                    Width =210
                    ColumnOrder =11
                    TabIndex =23
                    Name ="LocObsolete"
                    ControlSource ="LocObsolete"

                    LayoutCachedLeft =10350
                    LayoutCachedTop =2925
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =3165
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9480
                            Top =2925
                            Width =810
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Obsolete"
                            Caption ="Obsolete "
                            LayoutCachedLeft =9480
                            LayoutCachedTop =2925
                            LayoutCachedWidth =10290
                            LayoutCachedHeight =3165
                            RowStart =1
                            RowEnd =1
                            ColumnStart =8
                            ColumnEnd =8
                            ThemeFontIndex =1
                            BackThemeColorIndex =9
                            BackTint =40.0
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11430
                    Top =2925
                    Width =1170
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =24
                    Name ="LocOBsoletDt"
                    ControlSource ="LocObsdate"
                    Format ="Short Date"

                    LayoutCachedLeft =11430
                    LayoutCachedTop =2925
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =3165
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    BackThemeColorIndex =9
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =10620
                            Top =2925
                            Width =810
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label55"
                            Caption ="As of "
                            LayoutCachedLeft =10620
                            LayoutCachedTop =2925
                            LayoutCachedWidth =11430
                            LayoutCachedHeight =3165
                            RowStart =2
                            RowEnd =2
                            ColumnStart =8
                            ColumnEnd =8
                            ThemeFontIndex =1
                            BackThemeColorIndex =9
                            BackTint =40.0
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5835
                    Top =1350
                    Width =3570
                    ColumnWidth =4605
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =12
                    Name ="Location Add2"
                    ControlSource ="LAddrLine1"
                    EventProcPrefix ="Location_Add2"

                    LayoutCachedLeft =5835
                    LayoutCachedTop =1350
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =1590
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =4
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =1350
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label60"
                            Caption ="Address 2"
                            LayoutCachedLeft =4635
                            LayoutCachedTop =1350
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =1590
                            RowStart =3
                            RowEnd =3
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9465
                    Top =2610
                    Width =3030
                    ColumnWidth =2340
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =19
                    Name ="Sub Lessee"
                    ControlSource ="LocSubLessee"
                    StatusBarText ="allowable sublesee if not Lessee"
                    EventProcPrefix ="Sub_Lessee"

                    LayoutCachedLeft =9465
                    LayoutCachedTop =2610
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =2850
                    RowStart =7
                    RowEnd =7
                    ColumnStart =5
                    ColumnEnd =7
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7680
                            Top =2610
                            Width =1725
                            Height =240
                            FontSize =8
                            Name ="Label62"
                            Caption ="Allowable SubLessee"
                            LayoutCachedLeft =7680
                            LayoutCachedTop =2610
                            LayoutCachedWidth =9405
                            LayoutCachedHeight =2850
                            RowStart =7
                            RowEnd =7
                            ColumnStart =3
                            ColumnEnd =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12615
                    Top =6435
                    Width =660
                    ColumnWidth =1260
                    ColumnOrder =13
                    FontSize =8
                    TabIndex =29
                    Name ="ClientID"
                    ControlSource ="ClientREF"
                    StatusBarText ="link to Clients table"
                    FontName ="Tahoma"

                    LayoutCachedLeft =12615
                    LayoutCachedTop =6435
                    LayoutCachedWidth =13275
                    LayoutCachedHeight =6675
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13275
                    Top =6435
                    Width =720
                    ColumnWidth =1290
                    ColumnOrder =14
                    FontSize =8
                    TabIndex =30
                    Name ="Locationid"
                    ControlSource ="LocationID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =13275
                    LayoutCachedTop =6435
                    LayoutCachedWidth =13995
                    LayoutCachedHeight =6675
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =90
                    Top =450
                    Width =4470
                    Height =6090
                    FontSize =9
                    TabIndex =6
                    ForeColor =0
                    Name ="LocationList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LocationID, LocationType AS Type, LocationName AS Name, LState AS State, "
                        "Units, LocObsolete, ClientREF FROM ParkedLocationsUnitCount WHERE ClientREF=5295"
                        " ORDER BY LocationType, LocObsolete, LocationName;"
                    ColumnWidths ="0;864;2448;0;576;432"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    AllowValueListEdits =0

                    LayoutCachedLeft =90
                    LayoutCachedTop =450
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =6540
                    ColumnEnd =3
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10455
                    Top =1290
                    Width =2040
                    Height =270
                    FontSize =8
                    TabIndex =1
                    Name ="SaveMe"
                    Caption ="Save / Update"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10455
                    LayoutCachedTop =1290
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =1560
                    ForeThemeColorIndex =2
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    HoverThemeColorIndex =1
                    HoverShade =85.0
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    QuickStyle =22
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =75
                    Width =1665
                    Height =420
                    FontSize =18
                    FontWeight =700
                    Name ="Label65"
                    Caption ="Locations"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =75
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    DecimalPlaces =4
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14385
                    Top =6120
                    Width =870
                    FontSize =9
                    FontWeight =700
                    TabIndex =27
                    Name ="LTaxRate"
                    ControlSource ="LTaxRate"
                    Format ="Percent"
                    ValidationRule =">=0"
                    ValidationText ="Tax rate must be reater than or equal to 0"
                    DefaultValue ="0"

                    LayoutCachedLeft =14385
                    LayoutCachedTop =6120
                    LayoutCachedWidth =15255
                    LayoutCachedHeight =6360
                    RowStart =10
                    RowEnd =10
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14160
                    Top =5835
                    Width =900
                    FontSize =9
                    TabIndex =26
                    Name ="StTaxType"
                    ControlSource ="StTaxType"

                    LayoutCachedLeft =14160
                    LayoutCachedTop =5835
                    LayoutCachedWidth =15060
                    LayoutCachedHeight =6075
                    RowStart =10
                    RowEnd =10
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =12960
                            Top =5835
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label69"
                            Caption ="Tax Type "
                            LayoutCachedLeft =12960
                            LayoutCachedTop =5835
                            LayoutCachedWidth =14100
                            LayoutCachedHeight =6075
                            RowStart =10
                            RowEnd =10
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ListRows =2
                    Left =16605
                    Top =6120
                    Width =1560
                    FontSize =9
                    TabIndex =28
                    Name ="LTaxOption"
                    ControlSource ="LTaxOption"
                    RowSourceType ="Value List"
                    RowSource ="\"RR\";\"Upfront\""
                    StatusBarText ="For Optional states, customer selection of Upfront or Rent Receipts"
                    ControlTipText ="For Optional states, customer selection of Upfront or Rent Receipts"

                    LayoutCachedLeft =16605
                    LayoutCachedTop =6120
                    LayoutCachedWidth =18165
                    LayoutCachedHeight =6360
                    RowStart =10
                    RowEnd =10
                    ColumnStart =6
                    ColumnEnd =6
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =15315
                            Top =6120
                            Width =1230
                            Height =240
                            FontSize =8
                            Name ="Label80"
                            Caption ="Option Selection "
                            ControlTipText ="For Optional states, customer selection of Upfront or Rent Receipts"
                            LayoutCachedLeft =15315
                            LayoutCachedTop =6120
                            LayoutCachedWidth =16545
                            LayoutCachedHeight =6360
                            RowStart =10
                            RowEnd =10
                            ColumnStart =4
                            ColumnEnd =5
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9480
                    Top =3225
                    Width =3120
                    Height =300
                    FontSize =9
                    TabIndex =25
                    Name ="LocObsCmmt"
                    ControlSource ="LocObsCmt"

                    LayoutCachedLeft =9480
                    LayoutCachedTop =3225
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =3525
                    RowStart =3
                    RowEnd =6
                    ColumnStart =8
                    ColumnEnd =9
                    BackThemeColorIndex =9
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =13500
                    Top =6120
                    Width =825
                    Height =240
                    FontSize =8
                    Name ="Label68"
                    Caption ="Tax Rate"
                    LayoutCachedLeft =13500
                    LayoutCachedTop =6120
                    LayoutCachedWidth =14325
                    LayoutCachedHeight =6360
                    RowStart =10
                    RowEnd =10
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =8
                    ListWidth =2880
                    Left =5835
                    Top =3210
                    Width =3570
                    FontSize =9
                    TabIndex =21
                    Name ="DivisionID"
                    ControlSource ="DivisionID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DivisionID, DivisionName, ClientID From ClientDivisions WHERE Deleted=Fal"
                        "se and ClientID=5295 Order by DivisionName"
                    ColumnWidths ="0;2160;0"
                    AllowValueListEdits =0

                    LayoutCachedLeft =5835
                    LayoutCachedTop =3210
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =3450
                    RowStart =9
                    RowEnd =9
                    ColumnStart =1
                    ColumnEnd =4
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =3210
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label566"
                            Caption ="Division"
                            LayoutCachedLeft =4635
                            LayoutCachedTop =3210
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =3450
                            RowStart =9
                            RowEnd =9
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5835
                    Top =1650
                    Width =3570
                    FontSize =9
                    TabIndex =13
                    Name ="Text579"
                    ControlSource ="LAddrLine2"

                    LayoutCachedLeft =5835
                    LayoutCachedTop =1650
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =1890
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =4
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =1650
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label580"
                            Caption ="Address 3"
                            LayoutCachedLeft =4635
                            LayoutCachedTop =1650
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =1890
                            RowStart =4
                            RowEnd =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =12960
                    Top =735
                    Width =6192
                    Height =5055
                    FontSize =8
                    TabIndex =8
                    ForeColor =0
                    Name ="UnitsList"
                    RowSourceType ="Table/Query"
                    RowSource ="Select AssetID, MLNO, Schedule, [Unit#], Desc From UnitsParkedLocations Where Ga"
                        "rageLocationID=2341 Order by MLOrig, AssetID;"
                    ColumnWidths ="648;1152;1152;720;2520"
                    OnDblClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    AllowValueListEdits =0

                    LayoutCachedLeft =12960
                    LayoutCachedTop =735
                    LayoutCachedWidth =19152
                    LayoutCachedHeight =5790
                    ColumnEnd =3
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12960
                    Top =480
                    Width =648
                    FontSize =9
                    TabIndex =2
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SearchAssetId"
                    OnChange ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12960
                    LayoutCachedTop =480
                    LayoutCachedWidth =13608
                    LayoutCachedHeight =720
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15912
                    Top =480
                    Width =720
                    FontSize =9
                    TabIndex =5
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SearchUnit"
                    OnChange ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15912
                    LayoutCachedTop =480
                    LayoutCachedWidth =16632
                    LayoutCachedHeight =720
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12960
                    Width =6192
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BorderColor =0
                    ForeColor =10319446
                    Name ="Label629"
                    ControlSource ="=\"Units at Location: \" & IIf([UnitsList].[ListCount]>0,[UnitsList].[ListCount]"
                        "-1,0) & \" \""
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12960
                    LayoutCachedWidth =19152
                    LayoutCachedHeight =285
                    ColumnEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =13620
                    Top =480
                    Width =1152
                    FontSize =9
                    TabIndex =3
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SearchLease"
                    OnChange ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13620
                    LayoutCachedTop =480
                    LayoutCachedWidth =14772
                    LayoutCachedHeight =720
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14760
                    Top =480
                    Width =1152
                    FontSize =9
                    TabIndex =4
                    BackColor =13434879
                    BorderColor =967423
                    Name ="SearchSchedule"
                    OnChange ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14760
                    LayoutCachedTop =480
                    LayoutCachedWidth =15912
                    LayoutCachedHeight =720
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5835
                    Top =1950
                    Width =3570
                    Height =270
                    FontSize =9
                    TabIndex =14
                    Name ="LCity"
                    ControlSource ="LCity"
                    RowSourceType ="Table/Query"
                    RowSource ="Select City From vw_CCSZ_Lookup Where State ='GA'  AND County= 'Hancock'  Group "
                        "By City Order By City;"
                    ColumnWidths ="2160"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5835
                    LayoutCachedTop =1950
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =2220
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =4
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =1950
                            Width =1140
                            Height =270
                            FontSize =8
                            Name ="Label42"
                            Caption ="City, State, Zip"
                            LayoutCachedLeft =4635
                            LayoutCachedTop =1950
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =2220
                            RowStart =5
                            RowEnd =5
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2880
                    Left =5835
                    Top =2280
                    Width =3570
                    Height =270
                    FontSize =9
                    TabIndex =17
                    Name ="LCounty"
                    ControlSource ="LCounty"
                    RowSourceType ="Table/Query"
                    RowSource ="Select County From vw_CCSZ_Lookup Where State ='GA' Group By County, CountyID Or"
                        "der By County;"
                    ColumnWidths ="2160;720"
                    StatusBarText ="link to counties table"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5835
                    LayoutCachedTop =2280
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =2550
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =4
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =2280
                            Width =1140
                            Height =270
                            FontSize =8
                            Name ="Label43"
                            Caption ="County"
                            LayoutCachedLeft =4635
                            LayoutCachedTop =2280
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =2550
                            RowStart =6
                            RowEnd =6
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =4710
                    Top =3750
                    Width =7815
                    Height =2760
                    TabIndex =22
                    Name ="LocationTax"
                    SourceObject ="Form.LocationTax"
                    LinkChildFields ="LocationID"
                    LinkMasterFields ="LocationID"

                    LayoutCachedLeft =4710
                    LayoutCachedTop =3750
                    LayoutCachedWidth =12525
                    LayoutCachedHeight =6510
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13995
                    Top =6435
                    Width =720
                    FontSize =8
                    TabIndex =31
                    Name ="LStateID"
                    ControlSource ="LStateID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =13995
                    LayoutCachedTop =6435
                    LayoutCachedWidth =14715
                    LayoutCachedHeight =6675
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14715
                    Top =6435
                    Width =720
                    FontSize =8
                    TabIndex =32
                    Name ="LCountyID"
                    ControlSource ="LCountyID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =14715
                    LayoutCachedTop =6435
                    LayoutCachedWidth =15435
                    LayoutCachedHeight =6675
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15435
                    Top =6435
                    Width =720
                    FontSize =8
                    TabIndex =33
                    Name ="LCityID"
                    ControlSource ="LCityID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =15435
                    LayoutCachedTop =6435
                    LayoutCachedWidth =16155
                    LayoutCachedHeight =6675
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
            End
        End
    End
End
CodeBehindForm
' See "LocationsManager.cls"
