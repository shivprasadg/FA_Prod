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
    Width =19695
    DatasheetFontHeight =11
    ItemSuffix =786
    Left =7080
    Top =8175
    Right =26805
    Bottom =14145
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x00879e812fcde540
    End
    RecordSource ="SELECT * FROM Locations WHERE (((Locations.ClientREF)=[forms]![bank_View]![BankI"
        "D]));"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
        Begin Section
            CanGrow = NotDefault
            Height =5940
            Name ="Detail"
            AlternateBackColor =14545355
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =5460
                    Top =2910
                    Width =1020
                    Height =1950
                    BorderColor =8210719
                    Name ="Box781"
                    LayoutCachedLeft =5460
                    LayoutCachedTop =2910
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =4860
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =5460
                    Top =2610
                    Width =1020
                    Height =2250
                    BorderColor =8210719
                    Name ="Box780"
                    LayoutCachedLeft =5460
                    LayoutCachedTop =2610
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =4860
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =5460
                    Top =2280
                    Width =1020
                    Height =2580
                    BorderColor =8210719
                    Name ="Box779"
                    LayoutCachedLeft =5460
                    LayoutCachedTop =2280
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =4860
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =223
                    Left =6345
                    Top =2280
                    Width =240
                    Height =2430
                    Name ="Box782"
                    LayoutCachedLeft =6345
                    LayoutCachedTop =2280
                    LayoutCachedWidth =6585
                    LayoutCachedHeight =4710
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
                    Left =6720
                    Top =720
                    Width =3570
                    ColumnWidth =1650
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =1
                    Name ="Location Type"
                    ControlSource ="LocationType"
                    RowSourceType ="Value List"
                    RowSource ="\"Billing\";\"Garage\";\"Headquarters\";\"Registration\";\"Title\";\"Other\";Con"
                        "signment"
                    ColumnWidths ="1440"
                    EventProcPrefix ="Location_Type"
                    AllowValueListEdits =0

                    LayoutCachedLeft =6720
                    LayoutCachedTop =720
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =960
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
                            Left =5520
                            Top =720
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label3"
                            Caption ="Classification"
                            LayoutCachedLeft =5520
                            LayoutCachedTop =720
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =960
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
                    Left =6720
                    Top =1020
                    Width =3570
                    ColumnWidth =1800
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =2
                    Name ="LocationName"
                    ControlSource ="LocationName"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =1020
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =1260
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
                            Left =5520
                            Top =1020
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label39"
                            Caption ="Location Name"
                            LayoutCachedLeft =5520
                            LayoutCachedTop =1020
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =1260
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
                    Left =10365
                    Top =1020
                    Width =2070
                    ColumnWidth =1920
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =3
                    Name ="Locationsn"
                    ControlSource ="Locationsn"
                    StatusBarText ="short name for locaition or Lessee number"

                    LayoutCachedLeft =10365
                    LayoutCachedTop =1020
                    LayoutCachedWidth =12435
                    LayoutCachedHeight =1260
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
                            Left =10365
                            Top =720
                            Width =900
                            Height =240
                            FontSize =8
                            Name ="Label40"
                            Caption ="Short Name"
                            LayoutCachedLeft =10365
                            LayoutCachedTop =720
                            LayoutCachedWidth =11265
                            LayoutCachedHeight =960
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
                    Left =6720
                    Top =1320
                    Width =3570
                    ColumnWidth =2985
                    ColumnOrder =15
                    FontSize =9
                    TabIndex =4
                    Name ="LocationStreet"
                    ControlSource ="LocationStreet"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =1320
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =1560
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
                            Left =5475
                            Top =1320
                            Width =1185
                            Height =240
                            FontSize =8
                            Name ="Label41"
                            Caption ="Company Name"
                            LayoutCachedLeft =5475
                            LayoutCachedTop =1320
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =1560
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
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6720
                    Top =3150
                    Width =1770
                    Height =270
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =10
                    Name ="LZip"
                    ControlSource ="LZip"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =3150
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =3420
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
                    Left =6720
                    Top =3480
                    Width =1785
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =11
                    Name ="LPhone"
                    ControlSource ="LPhone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =3480
                    LayoutCachedWidth =8505
                    LayoutCachedHeight =3720
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
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =5520
                            Top =3480
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label46"
                            Caption ="Phone  / FAX"
                            LayoutCachedLeft =5520
                            LayoutCachedTop =3480
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =3720
                            RowStart =7
                            RowEnd =7
                            ThemeFontIndex =1
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
                    Left =8565
                    Top =3480
                    Width =1725
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =12
                    Name ="LFax"
                    ControlSource ="LFax"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =8565
                    LayoutCachedTop =3480
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =3720
                    RowStart =7
                    RowEnd =7
                    ColumnStart =3
                    ColumnEnd =4
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =7380
                    Top =4380
                    Width =210
                    ColumnOrder =11
                    TabIndex =19
                    Name ="LocObsolete"
                    ControlSource ="LocObsolete"

                    LayoutCachedLeft =7380
                    LayoutCachedTop =4380
                    LayoutCachedWidth =7590
                    LayoutCachedHeight =4620
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =5520
                            Top =4380
                            Width =1800
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Obsolete"
                            Caption ="Obsolete /Date / Reason "
                            LayoutCachedLeft =5520
                            LayoutCachedTop =4380
                            LayoutCachedWidth =7320
                            LayoutCachedHeight =4620
                            RowStart =1
                            RowEnd =1
                            ColumnStart =8
                            ColumnEnd =8
                            ThemeFontIndex =1
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
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7590
                    Top =4365
                    Width =1170
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =20
                    Name ="LocOBsoletDt"
                    ControlSource ="LocObsdate"
                    Format ="Short Date"

                    LayoutCachedLeft =7590
                    LayoutCachedTop =4365
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =4605
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
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6720
                    Top =1620
                    Width =3570
                    ColumnWidth =4605
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =5
                    Name ="Location Add2"
                    ControlSource ="LAddrLine1"
                    EventProcPrefix ="Location_Add2"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =1620
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =1860
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
                            Left =5520
                            Top =1620
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label60"
                            Caption ="Address 1"
                            LayoutCachedLeft =5520
                            LayoutCachedTop =1620
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =1860
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
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4230
                    Top =5340
                    Width =1365
                    ColumnWidth =2340
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =22
                    Name ="Sub Lessee"
                    ControlSource ="LocSubLessee"
                    StatusBarText ="allowable sublesee if not Lessee"
                    EventProcPrefix ="Sub_Lessee"

                    LayoutCachedLeft =4230
                    LayoutCachedTop =5340
                    LayoutCachedWidth =5595
                    LayoutCachedHeight =5580
                    RowStart =11
                    RowEnd =11
                    ColumnStart =2
                    ColumnEnd =4
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            TextAlign =3
                            Left =2670
                            Top =5340
                            Width =1500
                            Height =240
                            FontSize =8
                            Name ="Label62"
                            Caption ="Allowable SubLessee"
                            LayoutCachedLeft =2670
                            LayoutCachedTop =5340
                            LayoutCachedWidth =4170
                            LayoutCachedHeight =5580
                            RowStart =11
                            RowEnd =11
                            ColumnEnd =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =90
                    Top =735
                    Width =5280
                    Height =4485
                    FontSize =9
                    TabIndex =15
                    ForeColor =0
                    Name ="LocationList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ParkedLocationsUnitCount.LocationID, ParkedLocationsUnitCount.LocationTyp"
                        "e AS Type, ParkedLocationsUnitCount.LocationName AS Name, ParkedLocationsUnitCou"
                        "nt.LState AS State, ParkedLocationsUnitCount.Units, ParkedLocationsUnitCount.Tax"
                        "Rate, ParkedLocationsUnitCount.LocObsolete, ParkedLocationsUnitCount.ClientREF F"
                        "ROM ParkedLocationsUnitCount WHERE (((ParkedLocationsUnitCount.ClientREF)=[Forms"
                        "]![Bank_View]![BankID])) ORDER BY ParkedLocationsUnitCount.LocationName;"
                    ColumnWidths ="0;1008;2448;432;432;547"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    AllowValueListEdits =0

                    LayoutCachedLeft =90
                    LayoutCachedTop =735
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =5220
                    ColumnEnd =3
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10140
                    Top =180
                    Width =1095
                    Height =270
                    FontSize =8
                    TabIndex =14
                    Name ="SaveMe"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10140
                    LayoutCachedTop =180
                    LayoutCachedWidth =11235
                    LayoutCachedHeight =450
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
                Begin TextBox
                    DecimalPlaces =4
                    SpecialEffect =3
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2100
                    Top =5520
                    Width =780
                    FontSize =9
                    FontWeight =700
                    TabIndex =16
                    BackColor =52479
                    Name ="LTaxRate"
                    ControlSource ="LTaxRate"
                    Format ="Percent"
                    ValidationRule =">=0"
                    ValidationText ="Tax rate must be reater than or equal to 0"
                    DefaultValue ="0"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =5520
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =5760
                    RowStart =10
                    RowEnd =10
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =795
                    Top =5520
                    Width =675
                    FontSize =9
                    TabIndex =17
                    Name ="StTaxType"

                    LayoutCachedLeft =795
                    LayoutCachedTop =5520
                    LayoutCachedWidth =1470
                    LayoutCachedHeight =5760
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
                            OverlapFlags =93
                            TextAlign =3
                            Left =120
                            Top =5520
                            Width =675
                            Height =240
                            FontSize =8
                            Name ="Label69"
                            Caption ="Tax Type "
                            LayoutCachedLeft =120
                            LayoutCachedTop =5520
                            LayoutCachedWidth =795
                            LayoutCachedHeight =5760
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
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ListRows =2
                    Left =3555
                    Top =5520
                    Width =960
                    FontSize =9
                    TabIndex =18
                    Name ="LTaxOption"
                    ControlSource ="LTaxOption"
                    RowSourceType ="Value List"
                    RowSource ="\"RR\";\"Upfront\""
                    StatusBarText ="For Optional states, customer selection of Upfront or Rent Receipts"
                    ControlTipText ="For Optional states, customer selection of Upfront or Rent Receipts"

                    LayoutCachedLeft =3555
                    LayoutCachedTop =5520
                    LayoutCachedWidth =4515
                    LayoutCachedHeight =5760
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
                            OverlapFlags =223
                            TextAlign =1
                            Left =2910
                            Top =5520
                            Width =645
                            Height =240
                            FontSize =8
                            Name ="Label80"
                            Caption ="Option Selection "
                            ControlTipText ="For Optional states, customer selection of Upfront or Rent Receipts"
                            LayoutCachedLeft =2910
                            LayoutCachedTop =5520
                            LayoutCachedWidth =3555
                            LayoutCachedHeight =5760
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
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8880
                    Top =4365
                    Width =3525
                    FontSize =9
                    TabIndex =21
                    Name ="LocObsCmmt"
                    ControlSource ="LocObsCmt"

                    LayoutCachedLeft =8880
                    LayoutCachedTop =4365
                    LayoutCachedWidth =12405
                    LayoutCachedHeight =4605
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
                    OverlapFlags =223
                    TextAlign =3
                    Left =1470
                    Top =5520
                    Width =630
                    Height =240
                    FontSize =8
                    Name ="Label68"
                    Caption ="Tax Rate"
                    LayoutCachedLeft =1470
                    LayoutCachedTop =5520
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =5760
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =8940
                    Top =180
                    Width =1095
                    Height =270
                    FontSize =8
                    Name ="AddNew"
                    Caption ="New"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8940
                    LayoutCachedTop =180
                    LayoutCachedWidth =10035
                    LayoutCachedHeight =450
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
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6720
                    Top =1920
                    Width =3570
                    FontSize =9
                    TabIndex =6
                    Name ="LAddrLine2"
                    ControlSource ="LAddrLine2"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =1920
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =2160
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
                            Left =5520
                            Top =1920
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label580"
                            Caption ="Address 2"
                            LayoutCachedLeft =5520
                            LayoutCachedTop =1920
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =2160
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
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6720
                    Top =3780
                    Width =3570
                    FontSize =9
                    TabIndex =13
                    Name ="LEmail"
                    ControlSource ="LEmail"
                    StatusBarText ="link to counties table"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =3780
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =4020
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =4
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =5520
                            Top =3780
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label723"
                            Caption ="Email"
                            LayoutCachedLeft =5520
                            LayoutCachedTop =3780
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =4020
                            RowStart =8
                            RowEnd =8
                            ThemeFontIndex =1
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
                    Left =12840
                    Top =4380
                    Width =660
                    FontSize =8
                    TabIndex =23
                    Name ="ClientID"
                    ControlSource ="ClientREF"
                    StatusBarText ="link to Clients table"
                    FontName ="Tahoma"

                    LayoutCachedLeft =12840
                    LayoutCachedTop =4380
                    LayoutCachedWidth =13500
                    LayoutCachedHeight =4620
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13500
                    Top =4380
                    Width =720
                    FontSize =8
                    TabIndex =24
                    Name ="Locationid"
                    ControlSource ="LocationID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =13500
                    LayoutCachedTop =4380
                    LayoutCachedWidth =14220
                    LayoutCachedHeight =4620
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14220
                    Top =4380
                    Width =720
                    FontSize =8
                    TabIndex =25
                    Name ="LStateID"
                    ControlSource ="LStateID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =14220
                    LayoutCachedTop =4380
                    LayoutCachedWidth =14940
                    LayoutCachedHeight =4620
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14940
                    Top =4380
                    Width =720
                    FontSize =8
                    TabIndex =26
                    Name ="LCountyID"
                    ControlSource ="LCountyID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =14940
                    LayoutCachedTop =4380
                    LayoutCachedWidth =15660
                    LayoutCachedHeight =4620
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15660
                    Top =4380
                    Width =720
                    FontSize =8
                    TabIndex =27
                    Name ="LCityID"
                    ControlSource ="LCityID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =15660
                    LayoutCachedTop =4380
                    LayoutCachedWidth =16380
                    LayoutCachedHeight =4620
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =6720
                    Top =2520
                    Width =3570
                    Height =270
                    FontSize =9
                    TabIndex =8
                    Name ="LCity"
                    ControlSource ="LCity"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT City, CityID, County, CountyID From vw_CCSZ_Lookup WHERE StateID =34 GROU"
                        "P BY City, CityID, County, CountyID ORDER BY City, County;"
                    ColumnWidths ="1800;0;1800;0"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =2520
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =2790
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
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =6345
                            Top =2520
                            Width =315
                            Height =270
                            FontSize =8
                            Name ="Label42"
                            Caption ="City"
                            LayoutCachedLeft =6345
                            LayoutCachedTop =2520
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =2790
                            RowStart =5
                            RowEnd =5
                            ThemeFontIndex =1
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
                    ColumnCount =2
                    ListRows =20
                    ListWidth =3600
                    Left =6720
                    Top =2820
                    Width =3570
                    Height =270
                    FontSize =9
                    TabIndex =9
                    Name ="LCounty"
                    ControlSource ="LCounty"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT County, CountyID From vw_CCSZ_Lookup WHERE StateID =34 GROUP BY County, C"
                        "ountyID ORDER BY County;"
                    ColumnWidths ="3600;0"
                    StatusBarText ="link to counties table"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =2820
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =3090
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
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =6120
                            Top =2820
                            Width =540
                            Height =270
                            FontSize =8
                            Name ="Label43"
                            Caption ="County"
                            LayoutCachedLeft =6120
                            LayoutCachedTop =2820
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =3090
                            RowStart =6
                            RowEnd =6
                            ThemeFontIndex =1
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
                    Left =6720
                    Top =2203
                    Width =900
                    Height =270
                    FontSize =9
                    TabIndex =7
                    Name ="LState"
                    ControlSource ="LState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TaxStates.PostalAbbreviation, IIf([RentReceiptAvailable]=True And [Upfron"
                        "tAvailable]=True,\"RR/Upfront\",IIf([RentReceiptAvailable]=True,\"RR\",IIf([Upfr"
                        "ontAvailable]=True,\"Upfront\",\"None\"))) AS [Option], TaxStates.ISOCountryCode"
                        "2, TaxStates.ID FROM TaxStates GROUP BY TaxStates.PostalAbbreviation, IIf([RentR"
                        "eceiptAvailable]=True And [UpfrontAvailable]=True,\"RR/Upfront\",IIf([RentReceip"
                        "tAvailable]=True,\"RR\",IIf([UpfrontAvailable]=True,\"Upfront\",\"None\"))), Tax"
                        "States.ISOCountryCode2, TaxStates.ID, TaxStates.PostalAbbreviation HAVING (((Tax"
                        "States.ISOCountryCode2)=\"USA\" Or (TaxStates.ISOCountryCode2)=\"CAN\" Or (TaxSt"
                        "ates.ISOCountryCode2)=\"MEX\")) ORDER BY TaxStates.ISOCountryCode2 DESC , TaxSta"
                        "tes.PostalAbbreviation;"
                    ColumnWidths ="864;0;432;0"
                    StatusBarText ="lookup in lookup_state table"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =6720
                    LayoutCachedTop =2203
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =2473
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
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =120
                    Top =5280
                    Width =1560
                    Height =240
                    FontSize =8
                    Name ="Label776"
                    Caption ="Old Tax  system"
                    LayoutCachedLeft =120
                    LayoutCachedTop =5280
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =5520
                    RowStart =11
                    RowEnd =11
                    ColumnEnd =1
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =3
                    Left =6225
                    Top =2190
                    Width =435
                    Height =270
                    FontSize =8
                    Name ="Label777"
                    Caption ="State"
                    LayoutCachedLeft =6225
                    LayoutCachedTop =2190
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =2460
                    RowStart =5
                    RowEnd =5
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =3
                    Left =5520
                    Top =3150
                    Width =1140
                    Height =270
                    FontSize =8
                    Name ="Label778"
                    Caption ="Zip"
                    LayoutCachedLeft =5520
                    LayoutCachedTop =3150
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =3420
                    RowStart =6
                    RowEnd =6
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8565
                    Top =3150
                    Width =1725
                    Height =240
                    FontSize =8
                    TabIndex =28
                    ForeColor =8210719
                    Name ="LookUpTaxZip"
                    Caption ="Look Up Tax (Zip)"
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

                    LayoutCachedLeft =8565
                    LayoutCachedTop =3150
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =3390
                    PictureCaptionArrangement =5
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderWidth =1
                    BorderColor =8210719
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =8210719
                    PressedForeColor =8210719
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Width =5310
                    Height =435
                    FontSize =16
                    TabIndex =29
                    BorderColor =0
                    Name ="Label65"
                    ControlSource ="=\"Locations\""
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedWidth =5310
                    LayoutCachedHeight =435
                    ForeThemeColorIndex =2
                End
                Begin ToggleButton
                    OverlapFlags =93
                    Left =1099
                    Top =495
                    Width =2448
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =30
                    ForeColor =8210719
                    Name ="ToggleSortName"
                    DefaultValue ="0"
                    Caption ="↓↑ Location"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"

                    LayoutCachedLeft =1099
                    LayoutCachedTop =495
                    LayoutCachedWidth =3547
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =95
                    Left =90
                    Top =495
                    Width =1009
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =31
                    ForeColor =8210719
                    Name ="ToggleSortType"
                    DefaultValue ="0"
                    Caption ="↓↑ Type"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"

                    LayoutCachedLeft =90
                    LayoutCachedTop =495
                    LayoutCachedWidth =1099
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =95
                    Left =3547
                    Top =495
                    Width =432
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =32
                    ForeColor =8210719
                    Name ="ToggleSortState"
                    DefaultValue ="0"
                    Caption ="St"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"

                    LayoutCachedLeft =3547
                    LayoutCachedTop =495
                    LayoutCachedWidth =3979
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =4980
                    Top =495
                    Width =390
                    Height =240
                    FontSize =8
                    TabIndex =33
                    ForeColor =8210719
                    Name ="ClearSorts"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Edit Address"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000003255d6ab3255d65a0000000000000000000000000000000000000000 ,
                        0x3255d62d3255d693000000000000000000000000000000000000000000000000 ,
                        0x000000003255d6ae3255d6f93255d6360000000000000000000000003255d62d ,
                        0x3255d6db3255d61e000000000000000000000000000000000000000000000000 ,
                        0x000000003255d6153255d6db3255d6f03255d630000000003255d6303255d6ea ,
                        0x3255d66300000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000003255d61b3255d6c33255d6ed3255d66f3255d6ea3255d6ae ,
                        0x00000000000000000000000000000000000000000000000000000000727272ff ,
                        0x727272ff727272ff000000003255d6033255d6c63255d6ff3255d6de3255d60c ,
                        0x00000000000000000000000000000000000000000000000000000000727272ff ,
                        0x7272727e000000003255d6153255d6ab3255d6ff3255d6cf3255d6bd3255d696 ,
                        0x3255d609000000000000000000000000000000000000000000000000727272ff ,
                        0x000000003255d64e3255d6ed3255d6ff3255d6b73255d60c000000003255d645 ,
                        0x3255d6a53255d6420000000000000000000000000000000000000000727272ff ,
                        0x000000003255d6753255d6de3255d65a00000000000000000000000000000000 ,
                        0x000000003255d6270000000000000000000000000000000000000000727272ff ,
                        0x7272728100000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000007272726c727272ff ,
                        0x727272ff727272ff727272780000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000007272724e727272fc727272ff ,
                        0x727272ff727272ff727272ff7272725a00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000072727236727272f6727272ff727272ff ,
                        0x727272ff727272ff727272ff727272f972727242000000000000000000000000 ,
                        0x00000000000000000000000072727224727272ea727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272f07272722d0000000000000000 ,
                        0x000000000000000000000000727272d2727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272e40000000000000000 ,
                        0x000000000000000000000000727272f0727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4980
                    LayoutCachedTop =495
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =735
                    PictureCaptionArrangement =5
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderWidth =1
                    BorderColor =8210719
                    HoverColor =13017476
                    PressedColor =10319446
                    HoverForeColor =8210719
                    PressedForeColor =8210719
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    Left =4411
                    Top =495
                    Width =552
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =34
                    ForeColor =8210719
                    Name ="ToggleSortTax"
                    DefaultValue ="0"
                    Caption ="Tax%"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"

                    LayoutCachedLeft =4411
                    LayoutCachedTop =495
                    LayoutCachedWidth =4963
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =95
                    Left =3979
                    Top =495
                    Width =432
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =35
                    ForeColor =8210719
                    Name ="ToggleSortUnits"
                    DefaultValue ="0"
                    Caption ="↓↑ Units@"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"

                    LayoutCachedLeft =3979
                    LayoutCachedTop =495
                    LayoutCachedWidth =4411
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
' See "Bank_Locations.cls"
