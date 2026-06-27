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
    ItemSuffix =802
    Left =4875
    Top =7335
    Right =24675
    Bottom =14985
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x6673c231d27de640
    End
    RecordSource ="SELECT Locations.*, lookup_States.StTaxType FROM Locations LEFT JOIN lookup_Stat"
        "es ON Locations.LState = lookup_States.StateAbbrev WHERE Locations.LocationID = "
        "8251;"
    OnCurrent ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
            Height =7560
            Name ="Detail"
            AlternateBackColor =14545355
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =5460
                    Top =4860
                    Width =900
                    Height =2400
                    BorderColor =8210719
                    Name ="Box800"
                    LayoutCachedLeft =5460
                    LayoutCachedTop =4860
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =7260
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =223
                    Left =5700
                    Top =6900
                    Width =780
                    Height =300
                    Name ="Box801"
                    LayoutCachedLeft =5700
                    LayoutCachedTop =6900
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =7200
                    BackThemeColorIndex =1
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =95
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
                    ListRows =10
                    ListWidth =1440
                    Left =6720
                    Top =720
                    Width =3570
                    ColumnWidth =1650
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =2
                    Name ="Location Type"
                    ControlSource ="LocationType"
                    RowSourceType ="Value List"
                    RowSource ="\"Billing\";\"Garage\";\"Headquarters\";\"Registration\";\"Repair\";\"Title\";\""
                        "Other\""
                    ColumnWidths ="1440"
                    EventProcPrefix ="Location_Type"

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
                    TabIndex =3
                    Name ="LocationName"
                    ControlSource ="LocationName"
                    OnClick ="[Event Procedure]"

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
                    TabIndex =4
                    Name ="Locationsn"
                    ControlSource ="Locationsn"
                    StatusBarText ="short name for locaition or Lessee number"
                    AfterUpdate ="[Event Procedure]"

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
                    TabIndex =5
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
                    TabIndex =11
                    Name ="LZip"
                    ControlSource ="LZip"
                    AfterUpdate ="[Event Procedure]"

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
                    TabIndex =12
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
                    TabIndex =13
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
                    Top =4395
                    Width =210
                    Height =255
                    ColumnOrder =11
                    TabIndex =21
                    Name ="LocObsolete"
                    ControlSource ="LocObsolete"

                    LayoutCachedLeft =7380
                    LayoutCachedTop =4395
                    LayoutCachedWidth =7590
                    LayoutCachedHeight =4650
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
                            Top =4395
                            Width =1800
                            Height =255
                            FontSize =8
                            FontWeight =700
                            Name ="Obsolete"
                            Caption ="Obsolete /Date / Reason "
                            LayoutCachedLeft =5520
                            LayoutCachedTop =4395
                            LayoutCachedWidth =7320
                            LayoutCachedHeight =4650
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
                    Top =4380
                    Width =1170
                    Height =255
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =22
                    Name ="LocOBsoletDt"
                    ControlSource ="LocObsdate"
                    Format ="Short Date"

                    LayoutCachedLeft =7590
                    LayoutCachedTop =4380
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =4635
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
                    TabIndex =6
                    Name ="LAddrLine1"
                    ControlSource ="LAddrLine1"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

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
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10740
                    Top =3240
                    Width =1365
                    ColumnWidth =2340
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =25
                    Name ="Sub Lessee"
                    ControlSource ="LocSubLessee"
                    StatusBarText ="allowable sublesee if not Lessee"
                    EventProcPrefix ="Sub_Lessee"

                    LayoutCachedLeft =10740
                    LayoutCachedTop =3240
                    LayoutCachedWidth =12105
                    LayoutCachedHeight =3480
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
                            OverlapFlags =85
                            TextAlign =3
                            Left =10680
                            Top =2940
                            Width =1500
                            Height =240
                            FontSize =8
                            Name ="Label62"
                            Caption ="Allowable SubLessee"
                            LayoutCachedLeft =10680
                            LayoutCachedTop =2940
                            LayoutCachedWidth =12180
                            LayoutCachedHeight =3180
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
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =60
                    Top =780
                    Width =5280
                    Height =6300
                    FontSize =9
                    TabIndex =16
                    ForeColor =0
                    Name ="LocationList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT L.LocationID, L.LocationType AS Type, L.LocationName AS Name, L.LState AS"
                        " State, Count(U.UnitID) AS Units, Format([LTaxRate],'Percent') AS TaxRate, L.Loc"
                        "Obsolete, L.ClientREF FROM Locations AS L LEFT JOIN Units AS U ON L.LocationID ="
                        " U.UnitGarageLocationRef WHERE L.ClientGroupID =226 GROUP BY L.LocationID, L.Loc"
                        "ationType, L.LocationName, L.LState, L.LTaxRate, L.LocObsolete, L.ClientREF ORDE"
                        "R BY L.LocationType, L.LocationName, L.LocObsolete;"
                    ColumnWidths ="0;1008;2448;432;432;547"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    AllowValueListEdits =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =780
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =7080
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
                    TabIndex =15
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
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9735
                    Top =7140
                    Width =675
                    FontSize =9
                    TabIndex =19
                    Name ="StTaxType"

                    LayoutCachedLeft =9735
                    LayoutCachedTop =7140
                    LayoutCachedWidth =10410
                    LayoutCachedHeight =7380
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
                            Left =9060
                            Top =7140
                            Width =675
                            Height =240
                            FontSize =8
                            Name ="Label69"
                            Caption ="Tax Type "
                            LayoutCachedLeft =9060
                            LayoutCachedTop =7140
                            LayoutCachedWidth =9735
                            LayoutCachedHeight =7380
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
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ListRows =2
                    Left =8040
                    Top =7140
                    Width =960
                    FontSize =9
                    TabIndex =20
                    Name ="LTaxOption"
                    ControlSource ="LTaxOption"
                    RowSourceType ="Value List"
                    RowSource ="\"RR\";\"Upfront\""
                    StatusBarText ="For Optional states, customer selection of Upfront or Rent Receipts"
                    ControlTipText ="For Optional states, customer selection of Upfront or Rent Receipts"

                    LayoutCachedLeft =8040
                    LayoutCachedTop =7140
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =7380
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
                            OverlapFlags =93
                            TextAlign =1
                            Left =7500
                            Top =7140
                            Width =540
                            Height =240
                            FontSize =8
                            Name ="Label80"
                            Caption ="Option"
                            ControlTipText ="For Optional states, customer selection of Upfront or Rent Receipts"
                            LayoutCachedLeft =7500
                            LayoutCachedTop =7140
                            LayoutCachedWidth =8040
                            LayoutCachedHeight =7380
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
                    Top =4380
                    Width =3525
                    Height =255
                    FontSize =9
                    TabIndex =23
                    Name ="LocObsCmmt"
                    ControlSource ="LocObsCmt"

                    LayoutCachedLeft =8880
                    LayoutCachedTop =4380
                    LayoutCachedWidth =12405
                    LayoutCachedHeight =4635
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =8940
                    Top =180
                    Width =1095
                    Height =270
                    FontSize =8
                    TabIndex =1
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
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =8
                    ListWidth =2880
                    Left =6720
                    Top =4080
                    Width =3570
                    FontSize =9
                    TabIndex =18
                    Name ="DivisionID"
                    ControlSource ="DivisionID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DivisionID, DivisionName, ClientID From ClientDivisions WHERE Deleted=Fal"
                        "se and ClientGroupID=226 Order by DivisionName"
                    ColumnWidths ="0;2160;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =6720
                    LayoutCachedTop =4080
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =4320
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
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =5520
                            Top =4080
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label566"
                            Caption ="Division"
                            LayoutCachedLeft =5520
                            LayoutCachedTop =4080
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =4320
                            RowStart =9
                            RowEnd =9
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
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6720
                    Top =1920
                    Width =3570
                    FontSize =9
                    TabIndex =7
                    Name ="LAddrLine2"
                    ControlSource ="LAddrLine2"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

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
                    TabIndex =14
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
                Begin Tab
                    OverlapFlags =85
                    Left =12540
                    Top =60
                    Width =7155
                    Height =7080
                    FontSize =9
                    FontWeight =700
                    Name ="TabCtl128"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =12540
                    LayoutCachedTop =60
                    LayoutCachedWidth =19695
                    LayoutCachedHeight =7140
                    ThemeFontIndex =-1
                    BackColor =16777215
                    BackShade =100.0
                    BorderColor =6108695
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverThemeColorIndex =4
                    PressedColor =16645629
                    PressedThemeColorIndex =-1
                    PressedForeThemeColorIndex =2
                    PressedForeTint =100.0
                    ForeColor =0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =12615
                            Top =495
                            Width =7007
                            Height =6575
                            Name ="Units"
                            Caption ="65 Units @ Location"
                            LayoutCachedLeft =12615
                            LayoutCachedTop =495
                            LayoutCachedWidth =19622
                            LayoutCachedHeight =7070
                            Begin
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =12770
                                    Top =575
                                    Width =6450
                                    Height =225
                                    FontSize =8
                                    Name ="Label88"
                                    Caption =" ┌─ Search   ─┬────────┬────────┬──────────┬─────────────┐"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =12770
                                    LayoutCachedTop =575
                                    LayoutCachedWidth =19220
                                    LayoutCachedHeight =800
                                    ForeThemeColorIndex =0
                                    ForeTint =50.0
                                End
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    IMESentenceMode =3
                                    ColumnCount =6
                                    Left =12710
                                    Top =1025
                                    Width =6912
                                    Height =5055
                                    FontSize =8
                                    ForeColor =0
                                    Name ="UnitsList"
                                    RowSourceType ="Table/Query"
                                    RowSource ="Select AssetID, MLNO, Schedule, UnitNum, UnitDescShortVer as Descp, Status From "
                                        "vw_SixKeys Where LocationID=8251 Order by MLOrig, AssetID;"
                                    ColumnWidths ="648;1152;1152;720;2520;432"
                                    OnDblClick ="[Event Procedure]"
                                    LeftPadding =60
                                    RightPadding =75
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =12710
                                    LayoutCachedTop =1025
                                    LayoutCachedWidth =19622
                                    LayoutCachedHeight =6080
                                    ColumnEnd =3
                                    ThemeFontIndex =-1
                                    BorderThemeColorIndex =2
                                    BorderShade =100.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =12710
                                    Top =770
                                    Width =648
                                    FontSize =9
                                    TabIndex =1
                                    BackColor =13434879
                                    BorderColor =967423
                                    Name ="SearchAssetId"
                                    OnChange ="[Event Procedure]"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =12710
                                    LayoutCachedTop =770
                                    LayoutCachedWidth =13358
                                    LayoutCachedHeight =1010
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =15662
                                    Top =770
                                    Width =720
                                    FontSize =9
                                    TabIndex =2
                                    BackColor =13434879
                                    BorderColor =967423
                                    Name ="SearchUnit"
                                    OnChange ="[Event Procedure]"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =15662
                                    LayoutCachedTop =770
                                    LayoutCachedWidth =16382
                                    LayoutCachedHeight =1010
                                    ColumnStart =2
                                    ColumnEnd =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =13370
                                    Top =770
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

                                    LayoutCachedLeft =13370
                                    LayoutCachedTop =770
                                    LayoutCachedWidth =14522
                                    LayoutCachedHeight =1010
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =14510
                                    Top =770
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

                                    LayoutCachedLeft =14510
                                    LayoutCachedTop =770
                                    LayoutCachedWidth =15662
                                    LayoutCachedHeight =1010
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =16382
                                    Top =770
                                    Width =2520
                                    FontSize =9
                                    TabIndex =5
                                    BackColor =13434879
                                    BorderColor =967423
                                    Name ="SearchDesc"
                                    OnChange ="[Event Procedure]"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =16382
                                    LayoutCachedTop =770
                                    LayoutCachedWidth =18902
                                    LayoutCachedHeight =1010
                                    ColumnStart =2
                                    ColumnEnd =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =18902
                                    Top =770
                                    Width =432
                                    FontSize =9
                                    TabIndex =6
                                    BackColor =13434879
                                    BorderColor =967423
                                    Name ="SearchStatus"
                                    OnChange ="[Event Procedure]"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =18902
                                    LayoutCachedTop =770
                                    LayoutCachedWidth =19334
                                    LayoutCachedHeight =1010
                                    ColumnStart =2
                                    ColumnEnd =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =12615
                            Top =495
                            Width =7005
                            Height =6570
                            Name ="ContactsLinked"
                            Caption ="1 Contact Linked to Location"
                            LayoutCachedLeft =12615
                            LayoutCachedTop =495
                            LayoutCachedWidth =19620
                            LayoutCachedHeight =7065
                            Begin
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =6
                                    Left =12660
                                    Top =605
                                    Width =6891
                                    Height =3720
                                    FontSize =9
                                    ForeColor =0
                                    BorderColor =12632256
                                    Name ="LinkedContactsList"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT ContactID, LocationID, ContactType, [FirstName] & ' ' & [LastName] AS Nam"
                                        "e, Email, Phone From vw_LocationContacts Where LocationID= 8251;"
                                    ColumnWidths ="0;0;1152;2592;2592;0"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =0

                                    LayoutCachedLeft =12660
                                    LayoutCachedTop =605
                                    LayoutCachedWidth =19551
                                    LayoutCachedHeight =4325
                                    ThemeFontIndex =-1
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =12735
                                    Top =4695
                                    Width =3720
                                    Height =1305
                                    FontSize =8
                                    TabIndex =1
                                    BorderColor =0
                                    ForeColor =-2147483617
                                    Name ="ContactDetail"

                                    LayoutCachedLeft =12735
                                    LayoutCachedTop =4695
                                    LayoutCachedWidth =16455
                                    LayoutCachedHeight =6000
                                    ThemeFontIndex =1
                                End
                                Begin Rectangle
                                    SpecialEffect =0
                                    OverlapFlags =255
                                    Left =12660
                                    Top =4500
                                    Width =3900
                                    Height =1560
                                    Name ="Box656"
                                    LayoutCachedLeft =12660
                                    LayoutCachedTop =4500
                                    LayoutCachedWidth =16560
                                    LayoutCachedHeight =6060
                                    BorderThemeColorIndex =2
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =12705
                                    Top =4380
                                    Width =1485
                                    Height =240
                                    FontSize =8
                                    FontWeight =700
                                    Name ="Label657"
                                    Caption ="Contact Detail"
                                    LayoutCachedLeft =12705
                                    LayoutCachedTop =4380
                                    LayoutCachedWidth =14190
                                    LayoutCachedHeight =4620
                                    ThemeFontIndex =1
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =2
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin ComboBox
                                    ColumnHeads = NotDefault
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =4
                                    ListWidth =5760
                                    Left =17460
                                    Top =5580
                                    Width =1620
                                    FontSize =9
                                    TabIndex =2
                                    BorderColor =967423
                                    Name ="ContactsCBO"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Location_contacts.contactID, Location_contacts.clientRef, [contactFirst] "
                                        "& \" \" & [contactLast] & \" | \" & [contactClassification] AS Contact, Location"
                                        "_contacts.contactEmail AS Email, Location_contacts.contactFirst FROM Location_co"
                                        "ntacts WHERE (((Location_contacts.clientRef)=[Forms]![ClientsView]![clientID]));"
                                    ColumnWidths ="0;0;3312;2376"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =17460
                                    LayoutCachedTop =5580
                                    LayoutCachedWidth =19080
                                    LayoutCachedHeight =5820
                                    ThemeFontIndex =1
                                    BackThemeColorIndex =4
                                    BackTint =20.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OldBorderStyle =1
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =17460
                                            Top =5340
                                            Width =1575
                                            Height =240
                                            FontSize =8
                                            FontWeight =700
                                            BorderColor =967423
                                            Name ="Label730"
                                            Caption ="Select Contact to Link"
                                            LayoutCachedLeft =17460
                                            LayoutCachedTop =5340
                                            LayoutCachedWidth =19035
                                            LayoutCachedHeight =5580
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
                                Begin CommandButton
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    PictureType =1
                                    Left =17040
                                    Top =4380
                                    Width =2460
                                    Height =270
                                    FontSize =8
                                    TabIndex =3
                                    Name ="UnLinkedContact"
                                    Caption ="Un-Link Selected Location"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120

                                    LayoutCachedLeft =17040
                                    LayoutCachedTop =4380
                                    LayoutCachedWidth =19500
                                    LayoutCachedHeight =4650
                                    PictureCaptionArrangement =5
                                    ForeThemeColorIndex =1
                                    ForeShade =65.0
                                    UseTheme =1
                                    BorderColor =16777215
                                    HoverColor =1643706
                                    PressedThemeColorIndex =1
                                    PressedShade =95.0
                                    HoverForeThemeColorIndex =1
                                    HoverForeShade =65.0
                                    PressedForeThemeColorIndex =1
                                    PressedForeShade =65.0
                                    QuickStyle =22
                                    QuickStyleMask =-629
                                    Overlaps =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =12615
                            Top =495
                            Width =7005
                            Height =6570
                            Name ="Divisions"
                            LayoutCachedLeft =12615
                            LayoutCachedTop =495
                            LayoutCachedWidth =19620
                            LayoutCachedHeight =7065
                            Begin
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =16380
                                    Top =5520
                                    Width =2700
                                    Height =270
                                    FontSize =10
                                    FontWeight =700
                                    Name ="DeleteDivision"
                                    Caption ="Delete Selected Division"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =16380
                                    LayoutCachedTop =5520
                                    LayoutCachedWidth =19080
                                    LayoutCachedHeight =5790
                                    Alignment =1
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =255
                                    BackThemeColorIndex =9
                                    BackShade =75.0
                                    BorderColor =16777215
                                    HoverColor =2366701
                                    PressedColor =14211288
                                    HoverForeColor =967423
                                    Shadow =-1
                                    QuickStyle =30
                                    QuickStyleMask =-881
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =16380
                                    Top =1140
                                    Width =2700
                                    Height =270
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =1
                                    ForeColor =6108695
                                    Name ="AddNewDivision"
                                    Caption ="Add A New Division"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =16380
                                    LayoutCachedTop =1140
                                    LayoutCachedWidth =19080
                                    LayoutCachedHeight =1410
                                    Alignment =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =255
                                    BackColor =14270637
                                    BorderColor =16777215
                                    HoverColor =11525325
                                    PressedColor =14211288
                                    HoverForeColor =6108695
                                    Shadow =-1
                                    QuickStyle =30
                                    QuickStyleMask =-881
                                End
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =6
                                    Left =12720
                                    Top =1080
                                    Width =3306
                                    Height =4740
                                    FontSize =9
                                    TabIndex =2
                                    ForeColor =0
                                    BorderColor =12632256
                                    Name ="DivisionList"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT DivisionID, DivisionName, ClientID, ClientGroupID From ClientDivisions WH"
                                        "ERE Deleted=False and ClientID=4096 Order by DivisionName"
                                    ColumnWidths ="0;3240;0"
                                    GridlineColor =0

                                    LayoutCachedLeft =12720
                                    LayoutCachedTop =1080
                                    LayoutCachedWidth =16026
                                    LayoutCachedHeight =5820
                                    ThemeFontIndex =-1
                                    BackThemeColorIndex =-1
                                    BorderThemeColorIndex =-1
                                    BorderShade =100.0
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                    GridlineThemeColorIndex =-1
                                    GridlineShade =100.0
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =16380
                                    Top =1680
                                    Width =2700
                                    Height =1260
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =3
                                    Name ="SyncUnitsLocationDivisions"
                                    Caption ="Sync Units Locations with Locations Assigned to Divisons"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =16380
                                    LayoutCachedTop =1680
                                    LayoutCachedWidth =19080
                                    LayoutCachedHeight =2940
                                    Alignment =1
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =255
                                    BackThemeColorIndex =2
                                    BackShade =75.0
                                    BorderColor =16777215
                                    HoverColor =11525325
                                    PressedColor =14211288
                                    HoverForeThemeColorIndex =1
                                    PressedForeThemeColorIndex =1
                                    Shadow =-1
                                    QuickStyle =30
                                    QuickStyleMask =-881
                                End
                            End
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =5745
                    Top =4680
                    Width =6555
                    Height =2280
                    TabIndex =24
                    Name ="LocationTax"
                    SourceObject ="Form.LocationTax"
                    LinkChildFields ="LocationID"
                    LinkMasterFields ="LocationID"

                    LayoutCachedLeft =5745
                    LayoutCachedTop =4680
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =6960
                End
                Begin TextBox
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12540
                    Top =7320
                    Width =660
                    FontSize =8
                    TabIndex =26
                    Name ="ClientGroupID"
                    ControlSource ="ClientGroupID"
                    StatusBarText ="link to Clients table"
                    FontName ="Tahoma"

                    LayoutCachedLeft =12540
                    LayoutCachedTop =7320
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =7560
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13200
                    Top =7320
                    Width =720
                    FontSize =8
                    TabIndex =27
                    Name ="Locationid"
                    ControlSource ="LocationID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =13200
                    LayoutCachedTop =7320
                    LayoutCachedWidth =13920
                    LayoutCachedHeight =7560
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13920
                    Top =7320
                    Width =720
                    FontSize =8
                    TabIndex =28
                    Name ="LStateID"
                    ControlSource ="LStateID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =13920
                    LayoutCachedTop =7320
                    LayoutCachedWidth =14640
                    LayoutCachedHeight =7560
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14640
                    Top =7320
                    Width =720
                    FontSize =8
                    TabIndex =29
                    Name ="LCountyID"
                    ControlSource ="LCountyID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =14640
                    LayoutCachedTop =7320
                    LayoutCachedWidth =15360
                    LayoutCachedHeight =7560
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15360
                    Top =7320
                    Width =720
                    FontSize =8
                    TabIndex =30
                    Name ="LCityID"
                    ControlSource ="LCityID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =15360
                    LayoutCachedTop =7320
                    LayoutCachedWidth =16080
                    LayoutCachedHeight =7560
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
                    TabIndex =9
                    Name ="LCity"
                    ControlSource ="LCity"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT City, CityID, County, CountyID From vw_CCSZ_Lookup WHERE StateID =18 GROU"
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
                    TabIndex =10
                    Name ="LCounty"
                    ControlSource ="LCounty"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT County, CountyID From vw_CCSZ_Lookup WHERE StateID =18 GROUP BY County, C"
                        "ountyID ORDER BY County;"
                    ColumnWidths ="3600;0"
                    StatusBarText ="link to counties table"
                    OnClick ="[Event Procedure]"

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
                    TabIndex =8
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
                    Left =11865
                    Top =180
                    Width =480
                    Height =270
                    FontSize =8
                    TabIndex =31
                    Name ="LocationsReportView"
                    Caption ="ReportView"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272b7 ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffd7d7d7ff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffdbdbdbff727272ff ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272bd ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =11865
                    LayoutCachedTop =180
                    LayoutCachedWidth =12345
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
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11340
                    Top =180
                    Width =480
                    Height =270
                    FontSize =8
                    TabIndex =32
                    Name ="LocationsReportPrintView"
                    Caption ="PrintPreview"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff0000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff727272ff0000000000000000 ,
                        0x000000000000000000000000000000000000000000000000727272ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff727272ff0000000000000000 ,
                        0x00000000000000007272726f727272ea727272ff727272ff727272ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff727272ff727272ff727272ff ,
                        0x727272ea7272726c727272fc727272ff727272ff727272ff727272ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff727272ff727272ff00000000 ,
                        0x727272ff727272f6727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272de7272726672727218000000007272721872727266 ,
                        0x727272de727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ba7272720c90a4686c90a468c390a468ed90a468c390a4686c ,
                        0x7272720c727272ba727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272de787c700f90a468bd90a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468bd787c700f727272fc727272ff727272ff000000000000000000000000 ,
                        0x0000000090a4687590a468ff93a76df8e2e7d8ff9daf7aff90a468ff90a468ff ,
                        0x90a468ff90a4687572727284727272fc727272ff00000000727272ff727272ff ,
                        0x7272721b94a86ee792a66bfee2e7d8ffeef1e8ffe7ecdfff93a66cff90a468ff ,
                        0x90a468ff90a468de00000000000000000000000000000000727272f0ffffffff ,
                        0xffffff0691a569fbb5c29affdee4d3ff93a66cffdde3d1ffd1dac1ff90a468ff ,
                        0x90a468ff90a468f900000000000000000000000000000000727272f0ffffffff ,
                        0xffffff1b93a76de590a468ff90a468ff90a468ff98aa73fff2f4edffbac6a1ff ,
                        0x90a468ff90a468de00000000000000000000000000000000727272f0ffffffff ,
                        0xffffff6c90a4688490a468ff90a468ff90a468ff90a468ffa8b788ffe7ecdfff ,
                        0x90a468ff90a4688400000000000000000000000000000000727272ff727272ff ,
                        0x727272e17e866e1790a468cc90a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x93a66cd190a46809000000000000000000000000000000000000000000000000 ,
                        0x000000000000000090a4680690a4688190a468d590a468ff90a468d590a46881 ,
                        0x90a4680600000000
                    End

                    LayoutCachedLeft =11340
                    LayoutCachedTop =180
                    LayoutCachedWidth =11820
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
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8565
                    Top =3150
                    Width =1725
                    Height =240
                    FontSize =8
                    TabIndex =33
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
                    Left =60
                    Width =5310
                    Height =435
                    FontSize =16
                    TabIndex =34
                    BorderColor =0
                    Name ="Label65"
                    ControlSource ="=\"Locations\""
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedLeft =60
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =435
                    ForeThemeColorIndex =2
                End
                Begin ToggleButton
                    OverlapFlags =93
                    Left =1039
                    Top =495
                    Width =2448
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =35
                    ForeColor =8210719
                    Name ="ToggleSortName"
                    DefaultValue ="0"
                    Caption ="↓↑ Location"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"

                    LayoutCachedLeft =1039
                    LayoutCachedTop =495
                    LayoutCachedWidth =3487
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
                    OverlapFlags =247
                    Left =90
                    Top =495
                    Width =1009
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =36
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
                    Left =3487
                    Top =495
                    Width =432
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =37
                    ForeColor =8210719
                    Name ="ToggleSortState"
                    DefaultValue ="0"
                    Caption ="St"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"

                    LayoutCachedLeft =3487
                    LayoutCachedTop =495
                    LayoutCachedWidth =3919
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
                    OverlapFlags =85
                    Left =4980
                    Top =495
                    Width =390
                    Height =240
                    FontSize =8
                    TabIndex =38
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
                    Left =4351
                    Top =495
                    Width =552
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =39
                    ForeColor =8210719
                    Name ="ToggleSortTax"
                    DefaultValue ="0"
                    Caption ="Tax%"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"

                    LayoutCachedLeft =4351
                    LayoutCachedTop =495
                    LayoutCachedWidth =4903
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
                    OverlapFlags =87
                    Left =3919
                    Top =495
                    Width =432
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =40
                    ForeColor =8210719
                    Name ="ToggleSortUnits"
                    DefaultValue ="0"
                    Caption ="↓↑ Units@"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"

                    LayoutCachedLeft =3919
                    LayoutCachedTop =495
                    LayoutCachedWidth =4351
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
                    OverlapFlags =85
                    Left =8400
                    Top =180
                    Width =480
                    Height =270
                    FontSize =8
                    TabIndex =41
                    Name ="ExportLocations"
                    Caption ="Export"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000375c18ef375c18ff375c18ff ,
                        0x375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ef000000000000000000000000375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff0000000000000000000000001b2e0cff1b2e0cff1b2e0cff ,
                        0x1b2e0cff1b2e0cff1b2e0cff1b2e0cff294512ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff417c10ef417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff3f7710ff1b2e0cff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff417c10ff417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ffacc696ffffffffff598c2eff719d4cff ,
                        0xffffffffa0be88ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff4d841fffe7efe1ffd0dec3ffd0dec3ff ,
                        0xe7efe1ff4d841fff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff417c10ff7ca55bffffffffffffffffff ,
                        0x7ca55bff417c10ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff417c10ff719d4cffffffffffffffffff ,
                        0x719d4cff417c10ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff417c10ffe7efe1ffe7efe1ffd0dec3ff ,
                        0xe7efe1ff417c10ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff94b579ffffffffff88ad6aff65953dff ,
                        0xffffffffa0be88ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff365611ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ef417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff437e11ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff00000000000000000000000066a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff000000000000000066a3211066a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff000000000000000066a3211066a321ef66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ef
                    End

                    LayoutCachedLeft =8400
                    LayoutCachedTop =180
                    LayoutCachedWidth =8880
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
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =7260
                    Width =5220
                    Height =240
                    FontSize =9
                    TabIndex =42
                    Name ="DeleteLocation"
                    Caption ="Delete  Selected Location"
                    OnClick ="[Event Procedure]"
                    RightPadding =45
                    BottomPadding =45

                    LayoutCachedLeft =60
                    LayoutCachedTop =7260
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =7500
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1643706
                    BorderThemeColorIndex =5
                    ThemeFontIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-53
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11340
                    Top =7140
                    Width =960
                    FontSize =8
                    TabIndex =43
                    Name ="LastTaxUpdate"
                    ControlSource ="LastTaxUpdate"

                    LayoutCachedLeft =11340
                    LayoutCachedTop =7140
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =7380
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =10440
                            Top =7140
                            Width =900
                            Height =240
                            FontSize =8
                            Name ="Label799"
                            Caption ="Last Update"
                            LayoutCachedLeft =10440
                            LayoutCachedTop =7140
                            LayoutCachedWidth =11340
                            LayoutCachedHeight =7380
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =3
                    Left =5655
                    Top =7110
                    Width =705
                    Height =285
                    FontSize =10
                    Name ="Label68"
                    Caption ="Tax Rate"
                    LayoutCachedLeft =5655
                    LayoutCachedTop =7110
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =7395
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
                Begin TextBox
                    DecimalPlaces =4
                    SpecialEffect =3
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6420
                    Top =7080
                    Width =1020
                    Height =360
                    FontWeight =700
                    TabIndex =17
                    BackColor =12444887
                    Name ="LTaxRate"
                    ControlSource ="LTaxRate"
                    Format ="Percent"
                    ValidationRule =">=0"
                    ValidationText ="Tax rate must be greater than or equal to 0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"

                    LayoutCachedLeft =6420
                    LayoutCachedTop =7080
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =7440
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
            End
        End
    End
End
CodeBehindForm
' See "Locations_sub.cls"
