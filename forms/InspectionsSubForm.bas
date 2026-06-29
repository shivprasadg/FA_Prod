Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4320
    DatasheetFontHeight =11
    ItemSuffix =142
    Left =13800
    Top =10410
    Right =17955
    Bottom =17190
    RecSrcDt = Begin
        0xbb221f1e7156e640
    End
    RecordSource ="Select * From UnitInspections WHERE AssetID=363218 ORDER BY iDate DESC"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =570
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =1
                    PictureAlignment =0
                    PictureType =2
                    Top =360
                    Width =4320
                    Height =210
                    Name ="Image324"
                    Picture ="1_WhiteShadow"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =360
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =570
                    TabIndex =1
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2520
                    Top =60
                    Width =1290
                    Height =255
                    FontSize =9
                    FontWeight =700
                    Name ="AddInspection"
                    Caption ="Add Inspction"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2520
                    LayoutCachedTop =60
                    LayoutCachedWidth =3810
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =8
                    BackTint =100.0
                    BackShade =75.0
                    BorderColor =12566463
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedThemeColorIndex =8
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Width =1680
                    Height =495
                    FontSize =16
                    BorderColor =0
                    Name ="Label1053"
                    Caption ="Inspections"
                    FontName ="Segoe UI"
                    GridlineColor =0
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =495
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
            End
        End
        Begin Section
            Height =6600
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =852
                    Top =2232
                    Width =1236
                    Height =276
                    TabIndex =5
                    BackColor =16183539
                    Name ="iDate"
                    ControlSource ="iDate"
                    Format ="Short Date"
                    GroupTable =2
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =852
                    LayoutCachedTop =2232
                    LayoutCachedWidth =2088
                    LayoutCachedHeight =2508
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =7
                    ForeTint =100.0
                    ForeShade =75.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =96
                            Top =2232
                            Width =756
                            Height =276
                            FontSize =10
                            BorderColor =0
                            Name ="Label43"
                            Caption ="Date"
                            GroupTable =2
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =96
                            LayoutCachedTop =2232
                            LayoutCachedWidth =852
                            LayoutCachedHeight =2508
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =7
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =852
                    Top =1560
                    Width =1932
                    Height =270
                    TabIndex =2
                    BackColor =16183539
                    Name ="iVendorID"
                    ControlSource ="iVendorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorSNm, Vendors.VendorName, Vendors.VendorOb"
                        "solete FROM Vendors WHERE (((Vendors.VendorObsolete)=No)) ORDER BY Vendors.Vendo"
                        "rSNm;"
                    ColumnWidths ="0;10800;2880"
                    GroupTable =2
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =852
                    LayoutCachedTop =1560
                    LayoutCachedWidth =2784
                    LayoutCachedHeight =1830
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =96
                            Top =1560
                            Width =756
                            Height =270
                            FontSize =10
                            BorderColor =0
                            Name ="Label76"
                            Caption ="Co."
                            GroupTable =2
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =96
                            LayoutCachedTop =1560
                            LayoutCachedWidth =852
                            LayoutCachedHeight =1830
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =7
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =852
                    Top =1896
                    Width =1236
                    Height =276
                    TabIndex =3
                    BackColor =16183539
                    Name ="InvoiceNumber"
                    ControlSource ="InvoiceNumber"
                    DefaultValue ="0"
                    GroupTable =2
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =852
                    LayoutCachedTop =1896
                    LayoutCachedWidth =2088
                    LayoutCachedHeight =2172
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =7
                    ForeTint =100.0
                    ForeShade =75.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =96
                            Top =1896
                            Width =756
                            Height =276
                            FontSize =10
                            BorderColor =0
                            Name ="Label17"
                            Caption ="Inv#"
                            GroupTable =2
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =96
                            LayoutCachedTop =1896
                            LayoutCachedWidth =852
                            LayoutCachedHeight =2172
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =7
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =6
                    Left =2784
                    Top =2232
                    Width =1236
                    Height =276
                    TabIndex =6
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000002000000000000000200000001000000 ,
                        0x00000000b5cb8800000000000200000003000000050000000101000000000000 ,
                        0xf7b5800000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000320000000000
                    End
                    Name ="iResultID"
                    ControlSource ="iResultID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitInspectionResults.InpsectionResultId, UnitInspectionResults.Inspectio"
                        "nResult FROM UnitInspectionResults;"
                    ColumnWidths ="0;1008"
                    OnChange ="[Event Procedure]"
                    GroupTable =2
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =2784
                    LayoutCachedTop =2232
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =2508
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    ConditionalFormat14 = Begin
                        0x01000200000000000000020000000100000000000000b5cb8800010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010100 ,
                        0x0000000000f7b580000100000032000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2088
                            Top =2232
                            Width =696
                            Height =276
                            FontSize =10
                            BorderColor =0
                            Name ="Label13"
                            Caption ="Results"
                            GroupTable =2
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2088
                            LayoutCachedTop =2232
                            LayoutCachedWidth =2784
                            LayoutCachedHeight =2508
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =7
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =852
                    Top =3204
                    Width =1236
                    Height =276
                    TabIndex =11
                    BackColor =16183539
                    Name ="RequestedDate"
                    ControlSource ="RequestedDate"
                    GroupTable =2
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =852
                    LayoutCachedTop =3204
                    LayoutCachedWidth =2088
                    LayoutCachedHeight =3480
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =7
                    ForeTint =100.0
                    ForeShade =75.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =96
                            Top =3204
                            Width =756
                            Height =276
                            FontSize =10
                            BorderColor =0
                            Name ="Label44"
                            Caption ="Request"
                            GroupTable =2
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =96
                            LayoutCachedTop =3204
                            LayoutCachedWidth =852
                            LayoutCachedHeight =3480
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =7
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =852
                    Top =2544
                    Width =1236
                    Height =276
                    TabIndex =7
                    BackColor =16183539
                    Name ="iOdoRead"
                    ControlSource ="iOdoRead"
                    Format ="Standard"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =852
                    LayoutCachedTop =2544
                    LayoutCachedWidth =2088
                    LayoutCachedHeight =2820
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =7
                    ForeTint =100.0
                    ForeShade =75.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =96
                            Top =2544
                            Width =756
                            Height =276
                            FontSize =10
                            BorderColor =0
                            Name ="Label78"
                            Caption ="Mileage"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =96
                            LayoutCachedTop =2544
                            LayoutCachedWidth =852
                            LayoutCachedHeight =2820
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =7
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =852
                    Top =2856
                    Width =1236
                    Height =288
                    TabIndex =9
                    BackColor =16183539
                    Name ="iHourRead"
                    ControlSource ="iHourRead"
                    Format ="Standard"
                    GroupTable =2
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =852
                    LayoutCachedTop =2856
                    LayoutCachedWidth =2088
                    LayoutCachedHeight =3144
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =7
                    ForeTint =100.0
                    ForeShade =75.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =96
                            Top =2856
                            Width =756
                            Height =288
                            FontSize =10
                            BorderColor =0
                            Name ="Label1597"
                            Caption ="Hours"
                            GroupTable =2
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =96
                            LayoutCachedTop =2856
                            LayoutCachedWidth =852
                            LayoutCachedHeight =3144
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =7
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2784
                    Top =2544
                    Width =1236
                    Height =276
                    TabIndex =8
                    Name ="iPassedDate"
                    ControlSource ="iPassedDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2784
                    LayoutCachedTop =2544
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =2820
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =7
                    ForeTint =100.0
                    ForeShade =75.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2088
                            Top =2544
                            Width =696
                            Height =276
                            FontSize =10
                            BorderColor =0
                            Name ="Label1608"
                            Caption ="Passed"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =2088
                            LayoutCachedTop =2544
                            LayoutCachedWidth =2784
                            LayoutCachedHeight =2820
                            RowStart =3
                            RowEnd =3
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =7
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =6
                    Left =2784
                    Top =1896
                    Width =1236
                    Height =276
                    TabIndex =4
                    Name ="iTypeID"
                    ControlSource ="iTypeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitInspectionTypes.InspectionTypeId, UnitInspectionTypes.InspectionType "
                        "FROM UnitInspectionTypes;"
                    ColumnWidths ="0;1008"
                    GroupTable =2
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =2784
                    LayoutCachedTop =1896
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =2172
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2088
                            Top =1896
                            Width =696
                            Height =276
                            FontSize =10
                            BorderColor =0
                            Name ="Label2"
                            Caption ="Type"
                            GroupTable =2
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2088
                            LayoutCachedTop =1896
                            LayoutCachedWidth =2784
                            LayoutCachedHeight =2172
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =7
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
                Begin EmptyCell
                    Left =2784
                    Top =1560
                    Width =1236
                    Height =270
                    Name ="EmptyCell44"
                    GroupTable =2
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2784
                    LayoutCachedTop =1560
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =1830
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2784
                    Top =3204
                    Width =1236
                    Height =276
                    TabIndex =12
                    BackColor =16183539
                    Name ="iWebLink"
                    ControlSource ="iWebLink"
                    GroupTable =2
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =2784
                    LayoutCachedTop =3204
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =3480
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =7
                    ForeTint =100.0
                    ForeShade =75.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2088
                            Top =3204
                            Width =696
                            Height =276
                            FontSize =10
                            BorderColor =0
                            Name ="Label71"
                            Caption ="Link"
                            GroupTable =2
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2088
                            LayoutCachedTop =3204
                            LayoutCachedWidth =2784
                            LayoutCachedHeight =3480
                            RowStart =5
                            RowEnd =5
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =7
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =750
                    Top =165
                    Width =3279
                    Height =1260
                    FontSize =9
                    Name ="InspectionList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT I.InspnID, I.InspNum As Insp, I.iDate AS [Date], R.InspectionResult AS Re"
                        "sult, T.InspectionType AS Type  FROM (UnitInspections as I INNER JOIN UnitInspec"
                        "tionResults AS R ON I.iResultID = R.InpsectionResultId)  INNER JOIN UnitInspecti"
                        "onTypes as T ON I.iTypeID = T.InspectionTypeId  WHERE AssetID=363218 ORDER BY I."
                        "iDate DESC;"
                    ColumnWidths ="0;432;936;720;720"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =750
                    LayoutCachedTop =165
                    LayoutCachedWidth =4029
                    LayoutCachedHeight =1425
                    BorderThemeColorIndex =9
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2865
                    Top =1560
                    Width =510
                    Height =225
                    FontSize =8
                    TabIndex =13
                    Name ="InspnID"
                    ControlSource ="InspnID"

                    LayoutCachedLeft =2865
                    LayoutCachedTop =1560
                    LayoutCachedWidth =3375
                    LayoutCachedHeight =1785
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3420
                    Top =1560
                    Width =615
                    Height =225
                    FontSize =8
                    TabIndex =14
                    Name ="AssetID"
                    ControlSource ="AssetID"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =1560
                    LayoutCachedWidth =4035
                    LayoutCachedHeight =1785
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =135
                    Top =420
                    Width =435
                    Height =255
                    FontSize =8
                    Name ="GroupsLabel"
                    Caption ="Num"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =135
                    LayoutCachedTop =420
                    LayoutCachedWidth =570
                    LayoutCachedHeight =675
                    ThemeFontIndex =-1
                    BackThemeColorIndex =8
                    BackShade =50.0
                    BorderThemeColorIndex =7
                    BorderTint =100.0
                    BorderShade =75.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =135
                    Top =615
                    Width =435
                    Height =360
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BorderColor =9211020
                    ForeColor =7500402
                    Name ="InspNum"
                    ControlSource ="InspNum"
                    ControlTipText ="Number on Groups that will be inserted into ATLAAS Core"
                    GridlineColor =0

                    LayoutCachedLeft =135
                    LayoutCachedTop =615
                    LayoutCachedWidth =570
                    LayoutCachedHeight =975
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =156
                    Top =4536
                    Width =3996
                    Height =912
                    FontSize =12
                    BackColor =14865828
                    BorderColor =0
                    Name ="Label99"
                    Caption ="Fleet"
                    GroupTable =3
                    LayoutCachedLeft =156
                    LayoutCachedTop =4536
                    LayoutCachedWidth =4152
                    LayoutCachedHeight =5448
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =7
                    ForeTint =100.0
                    ForeShade =75.0
                    GroupTable =3
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =132
                    Top =3540
                    Width =3996
                    Height =915
                    FontSize =12
                    BackColor =8965045
                    BorderColor =0
                    Name ="Label100"
                    Caption ="Atlaas"
                    GroupTable =4
                    LayoutCachedLeft =132
                    LayoutCachedTop =3540
                    LayoutCachedWidth =4128
                    LayoutCachedHeight =4455
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =7
                    ForeTint =100.0
                    ForeShade =75.0
                    GroupTable =4
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =240
                    Top =4560
                    Width =3512
                    Height =842
                    FontSize =9
                    TabIndex =16
                    BackColor =16250085
                    BorderColor =12632256
                    ForeColor =0
                    Name ="iNotes2"
                    ControlSource ="iNotes2"

                    LayoutCachedLeft =240
                    LayoutCachedTop =4560
                    LayoutCachedWidth =3752
                    LayoutCachedHeight =5402
                    RowStart =7
                    RowEnd =7
                    ColumnEnd =7
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =225
                    Top =3570
                    Width =3525
                    Height =842
                    FontSize =9
                    TabIndex =15
                    BackColor =15398642
                    BorderColor =15858167
                    ForeColor =0
                    Name ="iNotes"
                    ControlSource ="iNotes"

                    LayoutCachedLeft =225
                    LayoutCachedTop =3570
                    LayoutCachedWidth =3750
                    LayoutCachedHeight =4412
                    RowStart =7
                    RowEnd =7
                    ColumnEnd =7
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2460
                    Top =5580
                    Width =1170
                    Height =210
                    FontSize =8
                    FontWeight =700
                    TabIndex =17
                    Name ="DeleteInsp"
                    Caption ="Delete Inspction"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2460
                    LayoutCachedTop =5580
                    LayoutCachedWidth =3630
                    LayoutCachedHeight =5790
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =1643706
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12566463
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =0
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedThemeColorIndex =8
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =5520
                    Width =1260
                    Height =225
                    FontSize =8
                    TabIndex =18
                    Name ="RequestedBy"
                    ControlSource ="RequestedBy"

                    LayoutCachedLeft =120
                    LayoutCachedTop =5520
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =5745
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2784
                    Top =2856
                    Width =1236
                    Height =288
                    TabIndex =10
                    BackColor =16183539
                    Name ="Text138"
                    ControlSource ="vInspNum"
                    DefaultValue ="0"
                    GroupTable =2
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =2784
                    LayoutCachedTop =2856
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =3144
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =7
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =7
                    ForeTint =100.0
                    ForeShade =75.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2088
                            Top =2856
                            Width =696
                            Height =288
                            FontSize =10
                            BorderColor =0
                            Name ="Label139"
                            Caption ="vNum"
                            GroupTable =2
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2088
                            LayoutCachedTop =2856
                            LayoutCachedWidth =2784
                            LayoutCachedHeight =3144
                            RowStart =4
                            RowEnd =4
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =7
                            ForeTint =100.0
                            ForeShade =75.0
                            GroupTable =2
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "InspectionsSubForm.cls"
