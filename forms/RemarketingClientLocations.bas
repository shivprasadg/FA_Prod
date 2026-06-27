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
    ItemSuffix =679
    Left =7260
    Top =7860
    Right =28665
    Bottom =16605
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x96fc62c91c07e540
    End
    RecordSource ="SELECT Locations.*, lookup_States.StTaxType FROM Locations LEFT JOIN lookup_Stat"
        "es ON Locations.LState = lookup_States.StateAbbrev ORDER BY Locations.LocationTy"
        "pe, Locations.LState, Locations.LCity;"
    OnCurrent ="[Event Procedure]"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =5790
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
                    Top =720
                    Width =3570
                    ColumnWidth =1650
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =8
                    Name ="Location Type"
                    ControlSource ="LocationType"
                    RowSourceType ="Value List"
                    RowSource ="\"Billing\";\"Garage\";\"Headquarters\";\"Registration\";\"Title\";\"Other\";Con"
                        "signment"
                    ColumnWidths ="1440"
                    EventProcPrefix ="Location_Type"
                    GroupTable =3
                    AllowValueListEdits =0

                    LayoutCachedLeft =5835
                    LayoutCachedTop =720
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =960
                    ColumnStart =1
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =720
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label3"
                            Caption ="Classification"
                            GroupTable =3
                            LayoutCachedLeft =4635
                            LayoutCachedTop =720
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =960
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5835
                    Top =1020
                    Width =3570
                    ColumnWidth =1800
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =9
                    Name ="LocationName"
                    ControlSource ="LocationName"
                    GroupTable =3

                    LayoutCachedLeft =5835
                    LayoutCachedTop =1020
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =1260
                    DatasheetCaption ="Location Name"
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =1020
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label39"
                            Caption ="Location Name"
                            GroupTable =3
                            LayoutCachedLeft =4635
                            LayoutCachedTop =1020
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =1260
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10425
                    Top =1020
                    Width =2070
                    ColumnWidth =1920
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =10
                    Name ="Locationsn"
                    ControlSource ="Locationsn"
                    StatusBarText ="short name for locaition or Lessee number"
                    GroupTable =3

                    LayoutCachedLeft =10425
                    LayoutCachedTop =1020
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =1260
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =9465
                            Top =1020
                            Width =900
                            Height =240
                            FontSize =8
                            Name ="Label40"
                            Caption ="Short Name"
                            GroupTable =3
                            LayoutCachedLeft =9465
                            LayoutCachedTop =1020
                            LayoutCachedWidth =10365
                            LayoutCachedHeight =1260
                            RowStart =1
                            RowEnd =1
                            ColumnStart =5
                            ColumnEnd =5
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
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
                    Top =1320
                    Width =3570
                    ColumnWidth =2985
                    ColumnOrder =15
                    FontSize =9
                    TabIndex =12
                    Name ="LocationStreet"
                    ControlSource ="LocationStreet"
                    GroupTable =3

                    LayoutCachedLeft =5835
                    LayoutCachedTop =1320
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =1560
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =1320
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label41"
                            Caption ="Address 1"
                            GroupTable =3
                            LayoutCachedLeft =4635
                            LayoutCachedTop =1320
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =1560
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
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
                    Top =2220
                    Width =3570
                    ColumnWidth =1380
                    ColumnOrder =6
                    FontSize =9
                    TabIndex =17
                    Name ="LCity"
                    ControlSource ="LCity"
                    GroupTable =3

                    LayoutCachedLeft =5835
                    LayoutCachedTop =2220
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =2460
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =2220
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label42"
                            Caption ="City, State, Zip"
                            GroupTable =3
                            LayoutCachedLeft =4635
                            LayoutCachedTop =2220
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =2460
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
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
                    Top =2520
                    Width =3570
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =20
                    Name ="countyRef"
                    ControlSource ="LCounty"
                    StatusBarText ="link to counties table"
                    GroupTable =3

                    LayoutCachedLeft =5835
                    LayoutCachedTop =2520
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =2760
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =2520
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label43"
                            Caption ="County"
                            GroupTable =3
                            LayoutCachedLeft =4635
                            LayoutCachedTop =2520
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =2760
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ListRows =8
                    ListWidth =1440
                    Left =9465
                    Top =2220
                    Width =900
                    ColumnWidth =855
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =18
                    Name ="LState"
                    ControlSource ="LState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lookup_States.StateAbbrev FROM lookup_States ORDER BY [StateAbbrev]; "
                    ColumnWidths ="1440"
                    StatusBarText ="lookup in lookup_state table"
                    GroupTable =3
                    AllowValueListEdits =0

                    LayoutCachedLeft =9465
                    LayoutCachedTop =2220
                    LayoutCachedWidth =10365
                    LayoutCachedHeight =2460
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10425
                    Top =2220
                    Width =1035
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =19
                    Name ="LZip"
                    ControlSource ="LZip"
                    InputMask ="00000\\-9999;;_"
                    GroupTable =3

                    LayoutCachedLeft =10425
                    LayoutCachedTop =2220
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =2460
                    RowStart =5
                    RowEnd =5
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5835
                    Top =2820
                    Width =1785
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =21
                    Name ="LPhone"
                    ControlSource ="LPhone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =3

                    LayoutCachedLeft =5835
                    LayoutCachedTop =2820
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =3060
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =2820
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label46"
                            Caption ="Phone "
                            GroupTable =3
                            LayoutCachedLeft =4635
                            LayoutCachedTop =2820
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =3060
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5835
                    Top =3120
                    Width =1785
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =23
                    Name ="LFax"
                    ControlSource ="LFax"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =3

                    LayoutCachedLeft =5835
                    LayoutCachedTop =3120
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =3360
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =3120
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label50"
                            Caption ="Fax "
                            GroupTable =3
                            LayoutCachedLeft =4635
                            LayoutCachedTop =3120
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =3360
                            RowStart =8
                            RowEnd =8
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10350
                    Top =4320
                    Width =210
                    ColumnOrder =11
                    TabIndex =11
                    Name ="LocObsolete"
                    ControlSource ="LocObsolete"

                    LayoutCachedLeft =10350
                    LayoutCachedTop =4320
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =4560
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
                            Top =4320
                            Width =810
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Obsolete"
                            Caption ="Obsolete "
                            LayoutCachedLeft =9480
                            LayoutCachedTop =4320
                            LayoutCachedWidth =10290
                            LayoutCachedHeight =4560
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
                    Top =4320
                    Width =1170
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =13
                    Name ="LocOBsoletDt"
                    ControlSource ="LocObsdate"
                    Format ="Short Date"

                    LayoutCachedLeft =11430
                    LayoutCachedTop =4320
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =4560
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
                            Top =4320
                            Width =810
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label55"
                            Caption ="As of "
                            LayoutCachedLeft =10620
                            LayoutCachedTop =4320
                            LayoutCachedWidth =11430
                            LayoutCachedHeight =4560
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
                    Top =1620
                    Width =3570
                    ColumnWidth =4605
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =15
                    Name ="Location Add2"
                    ControlSource ="LAddrLine1"
                    EventProcPrefix ="Location_Add2"
                    GroupTable =3

                    LayoutCachedLeft =5835
                    LayoutCachedTop =1620
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =1860
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =1620
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label60"
                            Caption ="Address 2"
                            GroupTable =3
                            LayoutCachedLeft =4635
                            LayoutCachedTop =1620
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =1860
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
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
                    Top =2820
                    Width =3030
                    ColumnWidth =2340
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =22
                    Name ="Sub Lessee"
                    ControlSource ="LocSubLessee"
                    StatusBarText ="allowable sublesee if not Lessee"
                    EventProcPrefix ="Sub_Lessee"
                    GroupTable =3

                    LayoutCachedLeft =9465
                    LayoutCachedTop =2820
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =3060
                    RowStart =7
                    RowEnd =7
                    ColumnStart =5
                    ColumnEnd =7
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7680
                            Top =2820
                            Width =1725
                            Height =240
                            FontSize =8
                            Name ="Label62"
                            Caption ="Allowable SubLessee"
                            GroupTable =3
                            LayoutCachedLeft =7680
                            LayoutCachedTop =2820
                            LayoutCachedWidth =9405
                            LayoutCachedHeight =3060
                            RowStart =7
                            RowEnd =7
                            ColumnStart =3
                            ColumnEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7740
                    Top =4980
                    Width =660
                    ColumnWidth =1260
                    ColumnOrder =13
                    FontSize =8
                    TabIndex =28
                    Name ="ClientID"
                    ControlSource ="ClientREF"
                    StatusBarText ="link to Clients table"
                    FontName ="Tahoma"

                    LayoutCachedLeft =7740
                    LayoutCachedTop =4980
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =5220
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8460
                    Top =4980
                    Width =720
                    ColumnWidth =1290
                    ColumnOrder =14
                    FontSize =8
                    TabIndex =29
                    Name ="Locationid"
                    ControlSource ="LocationID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =8460
                    LayoutCachedTop =4980
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =5220
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin EmptyCell
                    Left =9465
                    Top =720
                    Width =900
                    Name ="EmptyCell83"
                    GroupTable =3
                    LayoutCachedLeft =9465
                    LayoutCachedTop =720
                    LayoutCachedWidth =10365
                    LayoutCachedHeight =960
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =10425
                    Top =720
                    Width =2070
                    Name ="EmptyCell84"
                    GroupTable =3
                    LayoutCachedLeft =10425
                    LayoutCachedTop =720
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =960
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =9465
                    Top =1320
                    Width =900
                    Name ="EmptyCell91"
                    GroupTable =3
                    LayoutCachedLeft =9465
                    LayoutCachedTop =1320
                    LayoutCachedWidth =10365
                    LayoutCachedHeight =1560
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =10425
                    Top =1320
                    Width =2070
                    Name ="EmptyCell92"
                    GroupTable =3
                    LayoutCachedLeft =10425
                    LayoutCachedTop =1320
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =1560
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =9465
                    Top =1620
                    Width =900
                    Name ="EmptyCell99"
                    GroupTable =3
                    LayoutCachedLeft =9465
                    LayoutCachedTop =1620
                    LayoutCachedWidth =10365
                    LayoutCachedHeight =1860
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =10425
                    Top =1620
                    Width =2070
                    Name ="EmptyCell100"
                    GroupTable =3
                    LayoutCachedLeft =10425
                    LayoutCachedTop =1620
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =1860
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =9465
                    Top =2520
                    Width =900
                    Name ="EmptyCell119"
                    GroupTable =3
                    LayoutCachedLeft =9465
                    LayoutCachedTop =2520
                    LayoutCachedWidth =10365
                    LayoutCachedHeight =2760
                    RowStart =6
                    RowEnd =6
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =10425
                    Top =2520
                    Width =2070
                    Name ="EmptyCell120"
                    GroupTable =3
                    LayoutCachedLeft =10425
                    LayoutCachedTop =2520
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =2760
                    RowStart =6
                    RowEnd =6
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =90
                    Top =735
                    Width =4470
                    Height =5055
                    FontSize =9
                    TabIndex =6
                    ForeColor =0
                    Name ="LocationList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ParkedLocationsUnitCount.LocationID, ParkedLocationsUnitCount.LocationTyp"
                        "e AS Type, ParkedLocationsUnitCount.LocationName AS Name, ParkedLocationsUnitCou"
                        "nt.LState AS State, ParkedLocationsUnitCount.Units, ParkedLocationsUnitCount.Loc"
                        "Obsolete, ParkedLocationsUnitCount.ClientREF FROM ParkedLocationsUnitCount WHERE"
                        " (((ParkedLocationsUnitCount.ClientREF)=[Forms]![BuyersForm]![BuyerID])) ORDER B"
                        "Y ParkedLocationsUnitCount.LocationType, ParkedLocationsUnitCount.LocObsolete, P"
                        "arkedLocationsUnitCount.LocationName;"
                    ColumnWidths ="0;864;2448;0;576;432"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    AllowValueListEdits =0

                    LayoutCachedLeft =90
                    LayoutCachedTop =735
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =5790
                    ColumnEnd =3
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10680
                    Top =180
                    Width =1095
                    Height =270
                    FontSize =8
                    TabIndex =1
                    Name ="SaveMe"
                    Caption ="Save"
                    OnClick ="Refresh.Refresh"
                    FontName ="Tahoma"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10680
                    LayoutCachedTop =180
                    LayoutCachedWidth =11775
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
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5295
                    Top =4665
                    Width =1575
                    FontSize =9
                    TabIndex =30
                    Name ="LocABA"
                    ControlSource ="LocABA"
                    StatusBarText ="ABA ROUTING NUMBER location type billing only"
                    GroupTable =18

                    LayoutCachedLeft =5295
                    LayoutCachedTop =4665
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =4905
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =8
                    BackTint =40.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =1
                            Left =4635
                            Top =4665
                            Width =600
                            Height =240
                            FontSize =8
                            Name ="Label83"
                            Caption ="ABA No "
                            GroupTable =18
                            LayoutCachedLeft =4635
                            LayoutCachedTop =4665
                            LayoutCachedWidth =5235
                            LayoutCachedHeight =4905
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =4635
                    Top =4320
                    Width =4545
                    Height =285
                    FontSize =10
                    Name ="Label84"
                    Caption ="Wire Instructions for Customer"
                    FontName ="Segoe UI"
                    GroupTable =18
                    LayoutCachedLeft =4635
                    LayoutCachedTop =4320
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =4605
                    ColumnEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7605
                    Top =4665
                    Width =1575
                    FontSize =9
                    TabIndex =31
                    Name ="LocACH"
                    ControlSource ="LocACH"
                    StatusBarText ="ACH ROUTING NUMBER location type billing only"
                    GroupTable =18

                    LayoutCachedLeft =7605
                    LayoutCachedTop =4665
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =4905
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BackThemeColorIndex =8
                    BackTint =40.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =1
                            Left =6930
                            Top =4665
                            Width =615
                            Height =240
                            FontSize =8
                            Name ="Label85"
                            Caption ="ACH NO "
                            GroupTable =18
                            LayoutCachedLeft =6930
                            LayoutCachedTop =4665
                            LayoutCachedWidth =7545
                            LayoutCachedHeight =4905
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =4
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7680
                    Top =3720
                    Width =870
                    FontSize =9
                    FontWeight =700
                    TabIndex =26
                    Name ="LTaxRate"
                    ControlSource ="LTaxRate"
                    Format ="Percent"
                    ValidationRule =">=0"
                    ValidationText ="Tax rate must be reater than or equal to 0"
                    DefaultValue ="0"
                    GroupTable =3

                    LayoutCachedLeft =7680
                    LayoutCachedTop =3720
                    LayoutCachedWidth =8550
                    LayoutCachedHeight =3960
                    RowStart =10
                    RowEnd =10
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6795
                            Top =3720
                            Width =825
                            Height =240
                            FontSize =8
                            Name ="Label68"
                            Caption ="Tax Rate"
                            GroupTable =3
                            LayoutCachedLeft =6795
                            LayoutCachedTop =3720
                            LayoutCachedWidth =7620
                            LayoutCachedHeight =3960
                            RowStart =10
                            RowEnd =10
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
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
                    Left =5835
                    Top =3720
                    Width =900
                    FontSize =9
                    TabIndex =25
                    Name ="StTaxType"
                    ControlSource ="StTaxType"
                    GroupTable =3

                    LayoutCachedLeft =5835
                    LayoutCachedTop =3720
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =3960
                    RowStart =10
                    RowEnd =10
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =3720
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label69"
                            Caption ="Tax Type "
                            GroupTable =3
                            LayoutCachedLeft =4635
                            LayoutCachedTop =3720
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =3960
                            RowStart =10
                            RowEnd =10
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
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
                    Left =10425
                    Top =3720
                    Width =1035
                    FontSize =9
                    TabIndex =27
                    Name ="LTaxOption"
                    ControlSource ="LTaxOption"
                    RowSourceType ="Value List"
                    RowSource ="\"RR\";\"Upfront\""
                    StatusBarText ="For Optional states, customer selection of Upfront or Rent Receipts"
                    ControlTipText ="For Optional states, customer selection of Upfront or Rent Receipts"
                    GroupTable =3

                    LayoutCachedLeft =10425
                    LayoutCachedTop =3720
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =3960
                    RowStart =10
                    RowEnd =10
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8610
                            Top =3720
                            Width =1755
                            Height =240
                            FontSize =8
                            Name ="Label80"
                            Caption ="Option Selection "
                            ControlTipText ="For Optional states, customer selection of Upfront or Rent Receipts"
                            GroupTable =3
                            LayoutCachedLeft =8610
                            LayoutCachedTop =3720
                            LayoutCachedWidth =10365
                            LayoutCachedHeight =3960
                            RowStart =10
                            RowEnd =10
                            ColumnStart =4
                            ColumnEnd =5
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
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
                    Top =4620
                    Width =3120
                    Height =780
                    FontSize =9
                    TabIndex =14
                    Name ="LocObsCmmt"
                    ControlSource ="LocObsCmt"

                    LayoutCachedLeft =9480
                    LayoutCachedTop =4620
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =5400
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
                Begin EmptyCell
                    Left =11520
                    Top =3720
                    Width =975
                    Name ="EmptyCell537"
                    GroupTable =3
                    LayoutCachedLeft =11520
                    LayoutCachedTop =3720
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =3960
                    RowStart =10
                    RowEnd =10
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =11520
                    Top =2220
                    Width =975
                    Name ="EmptyCell542"
                    GroupTable =3
                    LayoutCachedLeft =11520
                    LayoutCachedTop =2220
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =2460
                    RowStart =5
                    RowEnd =5
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =7680
                    Top =3120
                    Width =1725
                    Name ="EmptyCell546"
                    GroupTable =3
                    LayoutCachedLeft =7680
                    LayoutCachedTop =3120
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =3360
                    RowStart =8
                    RowEnd =8
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =9465
                    Top =3120
                    Width =3030
                    Name ="EmptyCell547"
                    GroupTable =3
                    LayoutCachedLeft =9465
                    LayoutCachedTop =3120
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =3360
                    RowStart =8
                    RowEnd =8
                    ColumnStart =5
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1800
                    Top =105
                    Width =6210
                    Height =300
                    FontWeight =700
                    TabIndex =3
                    BorderColor =0
                    ForeColor =1643706
                    Name ="ObsoleteWarning"
                    ControlSource ="=IIf([LocObsolete]=-1,\"Note: This Location is Obsolete\",Null)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =105
                    LayoutCachedWidth =8010
                    LayoutCachedHeight =405
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9480
                    Top =180
                    Width =1095
                    Height =270
                    FontSize =8
                    Name ="AddNew"
                    Caption ="New"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9480
                    LayoutCachedTop =180
                    LayoutCachedWidth =10575
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
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =8
                    ListWidth =2880
                    Left =5835
                    Top =3420
                    Width =3570
                    FontSize =9
                    TabIndex =24
                    Name ="DivisionID"
                    ControlSource ="DivisionID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DivisionID, DivisionName, ClientID From ClientDivisions WHERE Deleted=Fal"
                        "se and ClientID= Order by DivisionName"
                    ColumnWidths ="0;2160;0"
                    GroupTable =3
                    AllowValueListEdits =0

                    LayoutCachedLeft =5835
                    LayoutCachedTop =3420
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =3660
                    RowStart =9
                    RowEnd =9
                    ColumnStart =1
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =3420
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label566"
                            Caption ="Division"
                            GroupTable =3
                            LayoutCachedLeft =4635
                            LayoutCachedTop =3420
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =3660
                            RowStart =9
                            RowEnd =9
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin EmptyCell
                    Left =9465
                    Top =3420
                    Width =900
                    Name ="EmptyCell574"
                    GroupTable =3
                    LayoutCachedLeft =9465
                    LayoutCachedTop =3420
                    LayoutCachedWidth =10365
                    LayoutCachedHeight =3660
                    RowStart =9
                    RowEnd =9
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =10425
                    Top =3420
                    Width =1035
                    Name ="EmptyCell575"
                    GroupTable =3
                    LayoutCachedLeft =10425
                    LayoutCachedTop =3420
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =3660
                    RowStart =9
                    RowEnd =9
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =11520
                    Top =3420
                    Width =975
                    Name ="EmptyCell576"
                    GroupTable =3
                    LayoutCachedLeft =11520
                    LayoutCachedTop =3420
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =3660
                    RowStart =9
                    RowEnd =9
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5835
                    Top =1920
                    Width =3570
                    FontSize =9
                    TabIndex =16
                    Name ="Text579"
                    ControlSource ="LAddrLine2"
                    GroupTable =3

                    LayoutCachedLeft =5835
                    LayoutCachedTop =1920
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =2160
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4635
                            Top =1920
                            Width =1140
                            Height =240
                            FontSize =8
                            Name ="Label580"
                            Caption ="Address 3"
                            GroupTable =3
                            LayoutCachedLeft =4635
                            LayoutCachedTop =1920
                            LayoutCachedWidth =5775
                            LayoutCachedHeight =2160
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =3
                        End
                    End
                End
                Begin EmptyCell
                    Left =9465
                    Top =1920
                    Width =900
                    Name ="EmptyCell598"
                    GroupTable =3
                    LayoutCachedLeft =9465
                    LayoutCachedTop =1920
                    LayoutCachedWidth =10365
                    LayoutCachedHeight =2160
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =10425
                    Top =1920
                    Width =1035
                    Name ="EmptyCell599"
                    GroupTable =3
                    LayoutCachedLeft =10425
                    LayoutCachedTop =1920
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =2160
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =11520
                    Top =1920
                    Width =975
                    Name ="EmptyCell600"
                    GroupTable =3
                    LayoutCachedLeft =11520
                    LayoutCachedTop =1920
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =2160
                    RowStart =4
                    RowEnd =4
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
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
                    TabIndex =7
                    ForeColor =0
                    Name ="UnitsList"
                    RowSourceType ="Table/Query"
                    RowSource ="Select AssetID, MLNO, Schedule, [Unit#], Desc From UnitsParkedLocations Where Ga"
                        "rageLocationID=745 Order by MLOrig, AssetID;"
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
                    TabIndex =5
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
                    TabIndex =4
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
                    TabIndex =2
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
                    TabIndex =32
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
                    TabIndex =33
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
            End
        End
    End
End
CodeBehindForm
' See "RemarketingClientLocations.cls"
