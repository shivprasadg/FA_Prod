Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =22176
    DatasheetFontHeight =11
    ItemSuffix =437
    Left =4845
    Top =4515
    Right =19290
    Bottom =11070
    Filter ="([Bulk_Update_SelectedUnits].[LState]=\"FL\")"
    ShortcutMenuBar ="SortClip"
    OrderBy ="[Bulk_Update_SelectedUnits].[LState] DESC"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0xa055b52e1528e640
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Bulk_Update_SelectedUnits"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1185
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =15607
                    Width =1020
                    Height =270
                    FontSize =9
                    Name ="Label431"
                    Caption ="┌─Taxes ──┐"
                    RightPadding =38
                    LayoutCachedLeft =15607
                    LayoutCachedWidth =16627
                    LayoutCachedHeight =270
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =5520
                    Width =6165
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label985"
                    Caption ="┌────────────────────────────  Locations  ────────────────────────────┐"
                    LayoutCachedLeft =5520
                    LayoutCachedWidth =11685
                    LayoutCachedHeight =300
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Top =990
                    Width =22176
                    Height =195
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =990
                    LayoutCachedWidth =22176
                    LayoutCachedHeight =1185
                    TabIndex =21
                    LayoutGroup =1
                    GroupTable =6
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =2
                    ListWidth =432
                    Left =12405
                    Top =555
                    Width =660
                    Height =255
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =13
                    BoundColumn =1
                    Name ="IRPCbo"
                    RowSourceType ="Value List"
                    RowSource ="\"Yes\";1;\"No\";0;\"P\""
                    ColumnWidths ="432;0"
                    StatusBarText ="Link to Contacts table"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =12405
                    LayoutCachedTop =555
                    LayoutCachedWidth =13065
                    LayoutCachedHeight =810
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =4335
                    Top =210
                    Width =1965
                    Height =240
                    FontSize =8
                    FontWeight =700
                    Name ="DeliverySubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4335
                    LayoutCachedTop =210
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =450
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =13365
                    Top =180
                    Width =765
                    Height =210
                    ColumnOrder =11
                    FontSize =8
                    TabIndex =18
                    Name ="txtCurrent"

                    LayoutCachedLeft =13365
                    LayoutCachedTop =180
                    LayoutCachedWidth =14130
                    LayoutCachedHeight =390
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =885
                    Top =870
                    Width =1305
                    Height =270
                    FontSize =9
                    Name ="Label182"
                    Caption ="Unit #"
                    RightPadding =38
                    LayoutCachedLeft =885
                    LayoutCachedTop =870
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin Label
                    OverlapFlags =247
                    Left =2265
                    Top =870
                    Width =1995
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label181"
                    Caption ="   VIN"
                    RightPadding =38
                    LayoutCachedLeft =2265
                    LayoutCachedTop =870
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =30
                    Top =870
                    Width =780
                    Height =270
                    FontSize =9
                    Name ="Label180"
                    Caption ="♦AssetID"
                    RightPadding =38
                    LayoutCachedLeft =30
                    LayoutCachedTop =870
                    LayoutCachedWidth =810
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =12405
                    Top =870
                    Width =660
                    Height =270
                    FontSize =9
                    Name ="LabelIRP"
                    Caption ="IRP"
                    RightPadding =38
                    LayoutCachedLeft =12405
                    LayoutCachedTop =870
                    LayoutCachedWidth =13065
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =13200
                    Top =870
                    Width =1125
                    Height =270
                    FontSize =9
                    Name ="Label176"
                    Caption ="Plate"
                    RightPadding =38
                    LayoutCachedLeft =13200
                    LayoutCachedTop =870
                    LayoutCachedWidth =14325
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =14370
                    Top =870
                    Width =870
                    Height =270
                    FontSize =9
                    Name ="Label177"
                    Caption ="RegExp"
                    RightPadding =38
                    LayoutCachedLeft =14370
                    LayoutCachedTop =870
                    LayoutCachedWidth =15240
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =2
                    Left =15420
                    Top =870
                    Width =675
                    Height =270
                    FontSize =9
                    Name ="Label178"
                    Caption ="Exempt"
                    RightPadding =38
                    LayoutCachedLeft =15420
                    LayoutCachedTop =870
                    LayoutCachedWidth =16095
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =17700
                    Top =870
                    Width =3945
                    Height =270
                    FontSize =9
                    Name ="Label179"
                    Caption ="Tax Notes"
                    RightPadding =38
                    LayoutCachedLeft =17700
                    LayoutCachedTop =870
                    LayoutCachedWidth =21645
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =7
                    ListWidth =10080
                    Left =4335
                    Top =555
                    Width =1965
                    Height =255
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =9
                    Name ="DeliveryCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LAddrLine1, Locat"
                        "ions.LCity, Locations.LState, Locations.LZip, Locations.LocationType FROM Locati"
                        "ons WHERE (((Locations.LocationType)=\"Garage\") AND ((Locations.ClientGroupID)="
                        "[Forms]![BulkChange_Form]![CLIENTgroupID]) AND ((Locations.LocObsolete)=No)) ORD"
                        "ER BY Locations.LocationName, Locations.LState;"
                    ColumnWidths ="0;3600;2160;2160;360;1008;720"
                    StatusBarText ="Link to Contacts table"
                    SmartTags ="\"Medit\""
                    RightPadding =38

                    LayoutCachedLeft =4335
                    LayoutCachedTop =555
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =810
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =4335
                    Top =870
                    Width =1965
                    Height =270
                    FontSize =9
                    Name ="Label242"
                    Caption ="Delivery"
                    RightPadding =38
                    LayoutCachedLeft =4335
                    LayoutCachedTop =870
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =6375
                    Top =870
                    Width =1965
                    Height =270
                    FontSize =9
                    Name ="Label112"
                    Caption ="Garage"
                    RightPadding =38
                    LayoutCachedLeft =6375
                    LayoutCachedTop =870
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =7
                    ListWidth =10080
                    Left =6375
                    Top =555
                    Width =1965
                    Height =255
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =10
                    Name ="GarageCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LAddrLine1, Locat"
                        "ions.LCity, Locations.LState, Locations.LZip, Locations.LocationType FROM Locati"
                        "ons WHERE (((Locations.LocationType)=\"Garage\") AND ((Locations.ClientGroupID)="
                        "[Forms]![BulkChange_Form]![CLIENTgroupID]) AND ((Locations.LocObsolete)=No)) ORD"
                        "ER BY Locations.LocationName, Locations.LState;"
                    ColumnWidths ="0;3600;2160;2160;360;1008;720"
                    StatusBarText ="Link to Contacts table"
                    SmartTags ="\"Medit\""
                    RightPadding =38

                    LayoutCachedLeft =6375
                    LayoutCachedTop =555
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =810
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =6375
                    Top =210
                    Width =1965
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="GarageSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6375
                    LayoutCachedTop =210
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =450
                    ColumnStart =4
                    ColumnEnd =4
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =8415
                    Top =870
                    Width =1965
                    Height =270
                    FontSize =9
                    Name ="Label113"
                    Caption ="Title"
                    RightPadding =38
                    LayoutCachedLeft =8415
                    LayoutCachedTop =870
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =7
                    ListWidth =10080
                    Left =8415
                    Top =555
                    Width =1965
                    Height =255
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =11
                    Name ="TitleCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT      Locations.LocationID,     Locations.LocationSN,     Locations.Locati"
                        "onType,     Locations.LAddrLine1,     Locations.LCity,     Locations.LState,    "
                        " Locations.LZip FROM Locations WHERE        Locations.LocationType <> \"Billing\""
                        "   AND Locations.ClientGroupID = [Forms]![BulkChange_Form]![CLIENTgroupID]   AND"
                        " Locations.LocObsolete = No  UNION ALL SELECT     0 AS LocationID,     \"Not App"
                        "licable\" AS LocationSN,     Null AS LocationType,     Null AS LAddrLine1,     N"
                        "ull AS LCity,     \"N/A\" AS LState,     Null AS LZip FROM ClientGroups WHERE Cl"
                        "ientGroups.ClientGroupID =[Forms]![BulkChange_Form]![CLIENTgroupID] ORDER BY Loc"
                        "ationSN;"
                    ColumnWidths ="0;3600;2160;2160;360;1008;720"
                    StatusBarText ="Link to Contacts table"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =8415
                    LayoutCachedTop =555
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =810
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =8415
                    Top =210
                    Width =1965
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="TitleSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8415
                    LayoutCachedTop =210
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =450
                    ColumnStart =5
                    ColumnEnd =5
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =10455
                    Top =870
                    Width =1875
                    Height =270
                    FontSize =9
                    Name ="Label195"
                    Caption ="Registration"
                    RightPadding =38
                    LayoutCachedLeft =10455
                    LayoutCachedTop =870
                    LayoutCachedWidth =12330
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =7
                    ListWidth =10080
                    Left =10455
                    Top =555
                    Width =1875
                    Height =255
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =12
                    Name ="RegistrationCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT      Locations.LocationID,     Locations.LocationSN,     Locations.Locati"
                        "onType,     Locations.LAddrLine1,     Locations.LCity,     Locations.LState,    "
                        " Locations.LZip FROM Locations WHERE        Locations.LocationType <> \"Billing\""
                        "   AND Locations.ClientGroupID = [Forms]![BulkChange_Form]![CLIENTgroupID]   AND"
                        " Locations.LocObsolete = No  UNION ALL SELECT     0 AS LocationID,     \"Not App"
                        "licable\" AS LocationSN,     Null AS LocationType,     Null AS LAddrLine1,     N"
                        "ull AS LCity,     \"N/A\" AS LState,     Null AS LZip FROM ClientGroups WHERE Cl"
                        "ientGroups.ClientGroupID =[Forms]![BulkChange_Form]![CLIENTgroupID]\015\012ORDER"
                        " BY LocationSN;\015\012"
                    ColumnWidths ="0;3600;2160;2160;360;1008;720"
                    StatusBarText ="Link to Contacts table"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =10455
                    LayoutCachedTop =555
                    LayoutCachedWidth =12330
                    LayoutCachedHeight =810
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =10455
                    Top =210
                    Width =1875
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="RegistrationSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10455
                    LayoutCachedTop =210
                    LayoutCachedWidth =12330
                    LayoutCachedHeight =450
                    ColumnStart =6
                    ColumnEnd =6
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =12405
                    Top =210
                    Width =660
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="IRPSubmit"
                    Caption ="▼"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12405
                    LayoutCachedTop =210
                    LayoutCachedWidth =13065
                    LayoutCachedHeight =450
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =14370
                    Top =210
                    Width =990
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="RegExpireSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14370
                    LayoutCachedTop =210
                    LayoutCachedWidth =15360
                    LayoutCachedHeight =450
                    ColumnStart =10
                    ColumnEnd =10
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14370
                    Top =555
                    Width =990
                    Height =255
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =14
                    Name ="RegExpire"
                    Format ="Short Date"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =14370
                    LayoutCachedTop =555
                    LayoutCachedWidth =15360
                    LayoutCachedHeight =810
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =2
                    ListWidth =432
                    Left =15435
                    Top =555
                    Width =645
                    Height =255
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =15
                    BoundColumn =1
                    Name ="ExemptCbo"
                    RowSourceType ="Value List"
                    RowSource ="\"Yes\";-1;\"No\";0"
                    ColumnWidths ="432;0"
                    StatusBarText ="Link to Contacts table"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =15435
                    LayoutCachedTop =555
                    LayoutCachedWidth =16080
                    LayoutCachedHeight =810
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =15435
                    Top =210
                    Width =645
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="ExemptSubmit"
                    Caption ="▼Send"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15435
                    LayoutCachedTop =210
                    LayoutCachedWidth =16080
                    LayoutCachedHeight =450
                    ColumnStart =11
                    ColumnEnd =11
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =16800
                    Top =870
                    Width =840
                    Height =270
                    FontSize =9
                    Name ="Label206"
                    Caption ="Weight"
                    RightPadding =38
                    LayoutCachedLeft =16800
                    LayoutCachedTop =870
                    LayoutCachedWidth =17640
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =16800
                    Top =210
                    Width =840
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="WeightSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =16800
                    LayoutCachedTop =210
                    LayoutCachedWidth =17640
                    LayoutCachedHeight =450
                    ColumnStart =12
                    ColumnEnd =12
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =16800
                    Top =555
                    Width =840
                    Height =255
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =16
                    Name ="Weight"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =16800
                    LayoutCachedTop =555
                    LayoutCachedWidth =17640
                    LayoutCachedHeight =810
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =17700
                    Top =210
                    Width =2430
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="NoteSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17700
                    LayoutCachedTop =210
                    LayoutCachedWidth =20130
                    LayoutCachedHeight =450
                    ColumnStart =13
                    ColumnEnd =13
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =17700
                    Top =555
                    Width =3945
                    Height =255
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =17
                    Name ="Notes"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =17700
                    LayoutCachedTop =555
                    LayoutCachedWidth =21645
                    LayoutCachedHeight =810
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Width =4215
                    Height =510
                    FontSize =20
                    FontWeight =700
                    Name ="Label115"
                    Caption ="Locations|Plate|Weight"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =4275
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =0
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =16125
                    Top =210
                    Width =645
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =19
                    Name ="IsUpFrontSubmit"
                    Caption ="▼Send"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =16125
                    LayoutCachedTop =210
                    LayoutCachedWidth =16770
                    LayoutCachedHeight =450
                    ColumnStart =11
                    ColumnEnd =11
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =2
                    ListWidth =432
                    Left =16125
                    Top =555
                    Width =645
                    Height =255
                    FontSize =9
                    TabIndex =20
                    BoundColumn =1
                    Name ="IsUpFrontCbo"
                    RowSourceType ="Value List"
                    RowSource ="\"Yes\";1;\"No\";0"
                    ColumnWidths ="432;0"
                    StatusBarText ="Link to Contacts table"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =16125
                    LayoutCachedTop =555
                    LayoutCachedWidth =16770
                    LayoutCachedHeight =810
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =16072
                    Top =870
                    Width =720
                    Height =270
                    FontSize =9
                    Name ="Label434"
                    Caption ="UpFront"
                    RightPadding =38
                    LayoutCachedLeft =16072
                    LayoutCachedTop =870
                    LayoutCachedWidth =16792
                    LayoutCachedHeight =1140
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                    BackThemeColorIndex =2
                    ForeThemeColorIndex =1
                End
            End
        End
        Begin Section
            Height =386
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12405
                    Top =45
                    Width =660
                    Height =255
                    FontSize =9
                    TabIndex =7
                    Name ="UnitIRP"
                    ControlSource ="UnitIRP"
                    Format ="Yes/No"
                    StatusBarText ="yes unit is IRP'ed"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =12405
                    LayoutCachedTop =45
                    LayoutCachedWidth =13065
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =7
                    ColumnEnd =7
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =13200
                    Top =45
                    Width =1125
                    Height =255
                    FontSize =9
                    TabIndex =8
                    Name ="UnitPlateNum"
                    ControlSource ="UnitPlateNum"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =13200
                    LayoutCachedTop =45
                    LayoutCachedWidth =14325
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =9
                    ColumnEnd =9
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14370
                    Top =45
                    Width =990
                    Height =255
                    FontSize =9
                    TabIndex =9
                    Name ="UnitRegExp"
                    ControlSource ="UnitRegExp"
                    Format ="Short Date"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =14370
                    LayoutCachedTop =45
                    LayoutCachedWidth =15360
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =10
                    ColumnEnd =10
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15435
                    Top =45
                    Width =645
                    Height =255
                    FontSize =9
                    TabIndex =10
                    Name ="UnitTaxExempt"
                    ControlSource ="UnitTaxExempt"
                    Format ="Yes/No"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =15435
                    LayoutCachedTop =45
                    LayoutCachedWidth =16080
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =11
                    ColumnEnd =11
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =17700
                    Top =45
                    Width =3945
                    Height =255
                    FontSize =9
                    TabIndex =12
                    Name ="UnitTaxNotes"
                    ControlSource ="UnitTaxNotes"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =17700
                    LayoutCachedTop =45
                    LayoutCachedWidth =21645
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =13
                    ColumnEnd =13
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =30
                    Top =45
                    Width =780
                    Height =255
                    FontSize =9
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =30
                    LayoutCachedTop =45
                    LayoutCachedWidth =810
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =2265
                    Top =45
                    Width =1995
                    Height =255
                    FontSize =9
                    TabIndex =2
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    StatusBarText ="VIN number must be unique in table"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2265
                    LayoutCachedTop =45
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =885
                    Top =45
                    Width =1305
                    Height =255
                    FontSize =9
                    TabIndex =1
                    Name ="UnitUnitNum"
                    ControlSource ="UnitUnitNum"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =885
                    LayoutCachedTop =45
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =7
                    ListWidth =8640
                    Left =4335
                    Top =45
                    Width =1965
                    Height =255
                    FontSize =9
                    TabIndex =3
                    Name ="UnitDeliveryLocationRef"
                    ControlSource ="UnitDeliveryLocationRef"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LAddrLine1, Locat"
                        "ions.LCity, Locations.LState, Locations.LZip, Locations.LocationType FROM Locati"
                        "ons WHERE (((Locations.LocationType)=\"Garage\") AND ((Locations.ClientGroupID)="
                        "[Forms]![BulkChange_Form]![CLIENTgroupID]) AND ((Locations.LocObsolete)=No)) ORD"
                        "ER BY Locations.LocationName, Locations.LState;"
                    ColumnWidths ="0;2160;2160;2160;360;1008;720"
                    StatusBarText ="Link to Contacts table"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =4335
                    LayoutCachedTop =45
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =7
                    ListWidth =8640
                    Left =6375
                    Top =45
                    Width =1965
                    Height =255
                    FontSize =9
                    TabIndex =4
                    Name ="UnitGarageLocationRef"
                    ControlSource ="UnitGarageLocationRef"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LAddrLine1, Locat"
                        "ions.LCity, Locations.LState, Locations.LZip, Locations.LocationType FROM Locati"
                        "ons WHERE (((Locations.LocationType)=\"Garage\") AND ((Locations.ClientGroupID)="
                        "[Forms]![BulkChange_Form]![CLIENTgroupID]) AND ((Locations.LocObsolete)=No)) OR "
                        "(((Locations.LocationType)=\"Garage\") AND ((Locations.ClientGroupID)=771) AND ("
                        "(Locations.LocObsolete)=No)) ORDER BY Locations.LocationName, Locations.LState;"
                    ColumnWidths ="0;2160;2160;2160;360;1008;720"
                    StatusBarText ="Link to Contacts table"
                    ValidationRule ="Is Not Null And >0"
                    ValidationText ="Domicile required if unknown, use the 'TBD' Loc"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =6375
                    LayoutCachedTop =45
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =7
                    ListWidth =8640
                    Left =8415
                    Top =45
                    Width =1965
                    Height =255
                    FontSize =9
                    TabIndex =5
                    Name ="UnitTitleLocationRef"
                    ControlSource ="UnitTitleLocationRef"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT      Locations.LocationID,     Locations.LocationSN,     Locations.Locati"
                        "onType,     Locations.LAddrLine1,     Locations.LCity,     Locations.LState,    "
                        " Locations.LZip FROM Locations WHERE        Locations.LocationType <> \"Billing\""
                        "   AND Locations.ClientGroupID = [Forms]![BulkChange_Form]![CLIENTgroupID]   AND"
                        " Locations.LocObsolete = No  UNION ALL SELECT     0 AS LocationID,     \"Not App"
                        "licable\" AS LocationSN,     Null AS LocationType,     Null AS LAddrLine1,     N"
                        "ull AS LCity,     \"N/A\" AS LState,     Null AS LZip FROM ClientGroups WHERE Cl"
                        "ientGroups.ClientGroupID =[Forms]![BulkChange_Form]![CLIENTgroupID]\015\012ORDER"
                        " BY LocationSN;\015\012"
                    ColumnWidths ="0;2160;2160;2160;360;1008;720"
                    StatusBarText ="Link to Contacts table"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =8415
                    LayoutCachedTop =45
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =7
                    ListWidth =8640
                    Left =10455
                    Top =45
                    Width =1875
                    Height =255
                    FontSize =9
                    TabIndex =6
                    Name ="UnitRegLocationRef"
                    ControlSource ="UnitRegLocationRef"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT      Locations.LocationID,     Locations.LocationSN,     Locations.Locati"
                        "onType,     Locations.LAddrLine1,     Locations.LCity,     Locations.LState,    "
                        " Locations.LZip FROM Locations WHERE        Locations.LocationType <> \"Billing\""
                        "   AND Locations.ClientGroupID = [Forms]![BulkChange_Form]![CLIENTgroupID]   AND"
                        " Locations.LocObsolete = No  UNION ALL SELECT     0 AS LocationID,     \"Not App"
                        "licable\" AS LocationSN,     Null AS LocationType,     Null AS LAddrLine1,     N"
                        "ull AS LCity,     \"N/A\" AS LState,     Null AS LZip FROM ClientGroups WHERE Cl"
                        "ientGroups.ClientGroupID =[Forms]![BulkChange_Form]![CLIENTgroupID]\015\012ORDER"
                        " BY LocationSN;\015\012"
                    ColumnWidths ="0;2160;2160;2160;360;1008;720"
                    StatusBarText ="Link to Contacts table"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =10455
                    LayoutCachedTop =45
                    LayoutCachedWidth =12330
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16800
                    Top =45
                    Width =840
                    Height =255
                    FontSize =9
                    TabIndex =11
                    Name ="UnitRegW"
                    ControlSource ="UnitRegW"
                    StatusBarText ="VIN number must be unique in table"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =16800
                    LayoutCachedTop =45
                    LayoutCachedWidth =17640
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =12
                    ColumnEnd =12
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =16147
                    Top =45
                    Width =645
                    Height =255
                    FontSize =9
                    TabIndex =13
                    BoundColumn =1
                    Name ="IsUpFront"
                    ControlSource ="IsUpFront"
                    RowSourceType ="Value List"
                    RowSource ="\"Yes\";1;\"No\";0"
                    ColumnWidths ="432;0"
                    Format ="Yes/No"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =16147
                    LayoutCachedTop =45
                    LayoutCachedWidth =16792
                    LayoutCachedHeight =300
                    RowStart =4
                    RowEnd =4
                    ColumnStart =11
                    ColumnEnd =11
                    BackThemeColorIndex =2
                    BackTint =20.0
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =21720
                    Top =60
                    Width =285
                    Height =255
                    FontSize =9
                    TabIndex =15
                    Name ="Lstate"
                    ControlSource ="LState"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =21720
                    LayoutCachedTop =60
                    LayoutCachedWidth =22005
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =11
                    ColumnEnd =11
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =314
                    Width =22176
                    Height =72
                    TabIndex =14
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End
                    HorizontalAnchor =2

                    LayoutCachedTop =314
                    LayoutCachedWidth =22176
                    LayoutCachedHeight =386
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
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
' See "BulkChange_Locations.cls"
