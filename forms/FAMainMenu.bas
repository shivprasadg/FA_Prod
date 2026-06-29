Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3660
    DatasheetFontHeight =10
    ItemSuffix =236
    Left =20055
    Top =2715
    Right =23715
    Bottom =11070
    TimerInterval =18000000
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x36e6ebbe4c45e440
    End
    Caption ="Main Menu"
    DatasheetFontName ="Arial"
    OnTimer ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    FitToScreen =255
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1290
            BackColor =0
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1065
                    Top =1050
                    Width =2340
                    ColumnOrder =0
                    FontSize =8
                    BackColor =0
                    BorderColor =-2147483617
                    Name ="UserName"

                    LayoutCachedLeft =1065
                    LayoutCachedTop =1050
                    LayoutCachedWidth =3405
                    LayoutCachedHeight =1290
                    ForeThemeColorIndex =1
                    ForeShade =95.0
                End
                Begin Image
                    SizeMode =1
                    PictureType =2
                    Left =120
                    Top =180
                    Width =3300
                    Height =900
                    Name ="ATLAASlogo"
                    OnDblClick ="[Event Procedure]"
                    Picture ="ATLAASLogoNew"

                    LayoutCachedLeft =120
                    LayoutCachedTop =180
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =1080
                    TabIndex =3
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =285
                    Top =1140
                    Width =465
                    Height =45
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =1
                    BackColor =0
                    BorderColor =-2147483617
                    ForeColor =-2147483617
                    Name ="UserGroup"

                    LayoutCachedLeft =285
                    LayoutCachedTop =1140
                    LayoutCachedWidth =750
                    LayoutCachedHeight =1185
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =127
                    BackStyle =0
                    IMESentenceMode =3
                    Width =2952
                    Height =180
                    ColumnOrder =3
                    TabIndex =2
                    BackColor =11830108
                    Name ="MinMax"
                    OnClick ="[Event Procedure]"

                    LayoutCachedWidth =2952
                    LayoutCachedHeight =180
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    Left =960
                    Top =840
                    Width =2430
                    Height =210
                    Name ="Label199"
                    Caption ="A FLEET ADVANTAGE PRODUCT"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =960
                    LayoutCachedTop =840
                    LayoutCachedWidth =3390
                    LayoutCachedHeight =1050
                    BackThemeColorIndex =0
                    ForeThemeColorIndex =2
                    ForeTint =40.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =2655
                    Width =945
                    Height =210
                    FontWeight =700
                    ForeColor =62207
                    Name ="DevLabel"
                    Caption ="Office"
                    LayoutCachedLeft =2655
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =210
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5340
            BackColor =14211288
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =216
                    Top =2580
                    Width =3195
                    Height =300
                    FontSize =11
                    TabIndex =8
                    Name ="Operations"
                    Caption ="Operations"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =216
                    LayoutCachedTop =2580
                    LayoutCachedWidth =3411
                    LayoutCachedHeight =2880
                    RowStart =8
                    RowEnd =8
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =231
                    Top =4185
                    Width =3195
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =13
                    Name ="ExitMe"
                    Caption ="Exit ATLAAS Core"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =231
                    LayoutCachedTop =4185
                    LayoutCachedWidth =3426
                    LayoutCachedHeight =4485
                    PictureCaptionArrangement =5
                    RowStart =13
                    RowEnd =13
                    ColumnEnd =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =20
                    BackColor =2366701
                    BorderColor =16777215
                    HoverColor =2366701
                    PressedColor =2366701
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =216
                    Top =2940
                    Width =3195
                    Height =300
                    FontSize =11
                    TabIndex =9
                    Name ="ReMarket"
                    Caption ="REMARKETING "
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =216
                    LayoutCachedTop =2940
                    LayoutCachedWidth =3411
                    LayoutCachedHeight =3240
                    RowStart =9
                    RowEnd =9
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =216
                    Top =1500
                    Width =3195
                    Height =300
                    FontSize =11
                    TabIndex =5
                    Name ="Accounting"
                    Caption ="Accounting"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =216
                    LayoutCachedTop =1500
                    LayoutCachedWidth =3411
                    LayoutCachedHeight =1800
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =216
                    Top =1140
                    Width =3195
                    Height =300
                    FontSize =11
                    TabIndex =4
                    Name ="LORPTM"
                    Caption ="LO Queries && Spreadsheets"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =216
                    LayoutCachedTop =1140
                    LayoutCachedWidth =3411
                    LayoutCachedHeight =1440
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =216
                    Top =780
                    Width =3195
                    Height =300
                    FontSize =11
                    TabIndex =3
                    Name ="LinkedDocManager"
                    Caption ="Linked Doc Manager"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =216
                    LayoutCachedTop =780
                    LayoutCachedWidth =3411
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =216
                    Top =60
                    Width =3195
                    Height =300
                    FontSize =11
                    TabIndex =2
                    Name ="Clients_Prospects"
                    Caption ="Open Clients Form"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =216
                    LayoutCachedTop =60
                    LayoutCachedWidth =3411
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =216
                    Top =420
                    Width =3195
                    Height =302
                    FontSize =11
                    TabIndex =1
                    Name ="Search_Clients"
                    Caption ="UnFunded Group Dashboard"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =216
                    LayoutCachedTop =420
                    LayoutCachedWidth =3411
                    LayoutCachedHeight =722
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =4595
                    Width =3450
                    FontSize =9
                    TabIndex =14
                    BorderColor =-2147483617
                    Name ="datebox"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                    HorizontalAnchor =2

                    LayoutCachedLeft =120
                    LayoutCachedTop =4595
                    LayoutCachedWidth =3570
                    LayoutCachedHeight =4835
                    RowStart =14
                    RowEnd =14
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =216
                    Top =3300
                    Width =3195
                    Height =300
                    FontSize =11
                    TabIndex =10
                    Name ="UnitLookUp"
                    Caption ="Unit Look Up"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =216
                    LayoutCachedTop =3300
                    LayoutCachedWidth =3411
                    LayoutCachedHeight =3600
                    RowStart =10
                    RowEnd =10
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =216
                    Top =2220
                    Width =3195
                    Height =300
                    FontSize =11
                    TabIndex =7
                    Name ="Management"
                    Caption ="Management "
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =216
                    LayoutCachedTop =2220
                    LayoutCachedWidth =3411
                    LayoutCachedHeight =2520
                    RowStart =7
                    RowEnd =7
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =216
                    Top =3660
                    Width =1920
                    Height =300
                    FontSize =11
                    TabIndex =11
                    Name ="ScheduleControl"
                    Caption ="Schedule Control"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =216
                    LayoutCachedTop =3660
                    LayoutCachedWidth =2136
                    LayoutCachedHeight =3960
                    RowStart =11
                    RowEnd =11
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =216
                    Top =1860
                    Width =3195
                    Height =300
                    FontSize =11
                    TabIndex =6
                    Name ="DocsMenu"
                    Caption ="Documentation"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =216
                    LayoutCachedTop =1860
                    LayoutCachedWidth =3411
                    LayoutCachedHeight =2160
                    RowStart =6
                    RowEnd =6
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =2259
                    Top =3660
                    Width =1152
                    Height =300
                    FontSize =11
                    TabIndex =12
                    Name ="TreeCommand"
                    Caption ="Tree"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2259
                    LayoutCachedTop =3660
                    LayoutCachedWidth =3411
                    LayoutCachedHeight =3960
                    RowStart =11
                    RowEnd =11
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin ListBox
                    SpecialEffect =5
                    OverlapFlags =85
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =7
                    Left =60
                    Top =5210
                    Width =3600
                    Height =75
                    FontSize =8
                    TabIndex =15
                    BackColor =14211288
                    ForeColor =5855577
                    Name ="EventList"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="3600;0;0;0;0;0;0"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =60
                    LayoutCachedTop =5210
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =5285
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeTint =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3624
                    Top =3735
                    Width =0
                    Height =0
                    FontSize =9
                    BorderColor =-2147483617
                    Name ="USERID"

                    LayoutCachedLeft =3624
                    LayoutCachedTop =3735
                    LayoutCachedWidth =3624
                    LayoutCachedHeight =3735
                    RowStart =1
                    RowEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1500
                    Top =4880
                    Width =2055
                    FontSize =9
                    TabIndex =16
                    BorderColor =-2147483617
                    Name ="Version"
                    ControlTipText ="C:\\Users\\ljohnson\\OneDrive - Fleet Advantage, LLC\\Desktop\\ATLAAS_UAT.accdb"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =4880
                    LayoutCachedWidth =3555
                    LayoutCachedHeight =5120
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =4895
                    Width =510
                    FontSize =9
                    TabIndex =17
                    BorderColor =-2147483617
                    ForeColor =8872517
                    Name ="HideShow"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Hide / Show the event list"

                    LayoutCachedLeft =60
                    LayoutCachedTop =4895
                    LayoutCachedWidth =570
                    LayoutCachedHeight =5135
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =570
                    Top =4895
                    Width =285
                    FontSize =9
                    TabIndex =18
                    BorderColor =-2147483617
                    ForeColor =8872517
                    Name ="S10"
                    ControlSource ="=10"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Show the last 10 events"

                    LayoutCachedLeft =570
                    LayoutCachedTop =4895
                    LayoutCachedWidth =855
                    LayoutCachedHeight =5135
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =855
                    Top =4895
                    Width =285
                    FontSize =9
                    TabIndex =19
                    BorderColor =-2147483617
                    ForeColor =8872517
                    Name ="S25"
                    ControlSource ="=25"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Show the last 25 events"

                    LayoutCachedLeft =855
                    LayoutCachedTop =4895
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =5135
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1140
                    Top =4895
                    Width =285
                    FontSize =9
                    TabIndex =20
                    BorderColor =-2147483617
                    ForeColor =8872517
                    Name ="S50"
                    ControlSource ="=50"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Show the last 50 events"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =4895
                    LayoutCachedWidth =1425
                    LayoutCachedHeight =5135
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin FormFooter
            Visible = NotDefault
            Height =0
            BackColor =15527148
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "FAMainMenu.cls"
