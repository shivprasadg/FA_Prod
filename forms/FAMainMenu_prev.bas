Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3360
    DatasheetFontHeight =10
    ItemSuffix =218
    Left =17385
    Top =3720
    Right =21570
    Bottom =10470
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
                    Left =960
                    Top =1050
                    Width =2340
                    ColumnOrder =0
                    FontSize =8
                    BackColor =0
                    BorderColor =-2147483617
                    Name ="UserName"

                    LayoutCachedLeft =960
                    LayoutCachedTop =1050
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =1290
                    ForeThemeColorIndex =1
                    ForeShade =95.0
                End
                Begin Image
                    SizeMode =1
                    PictureType =2
                    Left =15
                    Top =180
                    Width =3300
                    Height =900
                    Name ="ATLAASlogo"
                    OnDblClick ="[Event Procedure]"
                    Picture ="ATLAASLogoNew"

                    LayoutCachedLeft =15
                    LayoutCachedTop =180
                    LayoutCachedWidth =3315
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
                    Left =180
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

                    LayoutCachedLeft =180
                    LayoutCachedTop =1140
                    LayoutCachedWidth =645
                    LayoutCachedHeight =1185
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Width =2592
                    Height =165
                    ColumnOrder =3
                    TabIndex =2
                    BackColor =11830108
                    Name ="MinMax"
                    OnClick ="[Event Procedure]"

                    LayoutCachedWidth =2592
                    LayoutCachedHeight =165
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    Left =855
                    Top =840
                    Width =2430
                    Height =210
                    Name ="Label199"
                    Caption ="A FLEET ADVANTAGE PRODUCT"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =855
                    LayoutCachedTop =840
                    LayoutCachedWidth =3285
                    LayoutCachedHeight =1050
                    BackThemeColorIndex =0
                    ForeThemeColorIndex =2
                    ForeTint =40.0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =2940
                    Width =360
                    Height =210
                    FontWeight =700
                    ForeColor =62207
                    Name ="DevLabel"
                    Caption ="DEV"
                    LayoutCachedLeft =2940
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =210
                End
            End
        End
        Begin Section
            Height =5467
            BackColor =14211288
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =150
                    Top =2940
                    Width =3060
                    Height =300
                    FontSize =11
                    TabIndex =8
                    Name ="Operations"
                    Caption ="Operations"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =150
                    LayoutCachedTop =2940
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =3240
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
                    OverlapFlags =93
                    Left =150
                    Top =4500
                    Width =3060
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    Name ="ExitMe"
                    Caption ="Exit System"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =150
                    LayoutCachedTop =4500
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =4800
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
                    Left =150
                    Top =3300
                    Width =3060
                    Height =300
                    FontSize =11
                    TabIndex =9
                    Name ="ReMarket"
                    Caption ="REMARKETING "
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =150
                    LayoutCachedTop =3300
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =3600
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
                    Left =150
                    Top =1860
                    Width =3060
                    Height =300
                    FontSize =11
                    TabIndex =5
                    Name ="Accounting"
                    Caption ="Accounting"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =150
                    LayoutCachedTop =1860
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =2160
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
                    Left =150
                    Top =1500
                    Width =3060
                    Height =300
                    FontSize =11
                    TabIndex =4
                    Name ="LORPTM"
                    Caption ="LO Queries && Spreadsheets"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =150
                    LayoutCachedTop =1500
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =1800
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
                    Left =150
                    Top =1140
                    Width =3060
                    Height =300
                    FontSize =11
                    TabIndex =3
                    Name ="Search_Transactions"
                    Caption ="Search Transactions"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =150
                    LayoutCachedTop =1140
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =1440
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
                    OverlapFlags =95
                    Left =150
                    Top =780
                    Width =3060
                    Height =300
                    FontSize =11
                    TabIndex =2
                    Name ="View_Transactions"
                    Caption ="View Transactions"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =150
                    LayoutCachedTop =780
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
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
                    Left =150
                    Top =420
                    Width =3060
                    Height =300
                    FontSize =11
                    TabIndex =1
                    Name ="Clients_Prospects"
                    Caption ="Open Clients Form"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =150
                    LayoutCachedTop =420
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =720
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
                    OverlapFlags =87
                    Left =150
                    Top =60
                    Width =3060
                    Height =302
                    FontSize =11
                    Name ="Search_Clients"
                    Caption ="Search Clients"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =150
                    LayoutCachedTop =60
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =362
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
                    FontItalic = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =150
                    Top =4860
                    Width =3060
                    FontSize =9
                    TabIndex =13
                    BorderColor =-2147483617
                    Name ="datebox"
                    ControlSource ="=Date()"
                    Format ="Long Date"

                    LayoutCachedLeft =150
                    LayoutCachedTop =4860
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =5100
                    RowStart =14
                    RowEnd =14
                    ColumnEnd =1
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
                    Left =2040
                    Top =5160
                    Width =1170
                    FontSize =9
                    TabIndex =15
                    BorderColor =-2147483617
                    Name ="Version"

                    LayoutCachedLeft =2040
                    LayoutCachedTop =5160
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =5400
                    RowStart =15
                    RowEnd =15
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
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =150
                    Top =5160
                    Width =1830
                    FontSize =9
                    TabIndex =14
                    BorderColor =-2147483617
                    Name ="USERID"

                    LayoutCachedLeft =150
                    LayoutCachedTop =5160
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =5400
                    RowStart =15
                    RowEnd =15
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =150
                    Top =3660
                    Width =3060
                    Height =300
                    FontSize =11
                    TabIndex =10
                    Name ="UnitLookUp"
                    Caption ="Unit Look Up"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =150
                    LayoutCachedTop =3660
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =3960
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
                    Left =150
                    Top =2580
                    Width =3060
                    Height =300
                    FontSize =11
                    TabIndex =7
                    Name ="Management"
                    Caption ="Management "
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =150
                    LayoutCachedTop =2580
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =2880
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
                    Left =150
                    Top =4020
                    Width =1950
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

                    LayoutCachedLeft =150
                    LayoutCachedTop =4020
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =4320
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
                    Left =150
                    Top =2220
                    Width =3060
                    Height =300
                    FontSize =11
                    TabIndex =6
                    Name ="DocsMenu"
                    Caption ="Documentation"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =150
                    LayoutCachedTop =2220
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =2520
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
                    Left =2160
                    Top =4020
                    Width =1050
                    Height =300
                    FontSize =11
                    TabIndex =16
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

                    LayoutCachedLeft =2160
                    LayoutCachedTop =4020
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =4320
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
            End
        End
        Begin FormFooter
            Visible = NotDefault
            Height =0
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "FAMainMenu_prev.cls"
