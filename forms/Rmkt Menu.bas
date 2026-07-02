Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11004
    DatasheetFontHeight =11
    ItemSuffix =604
    Left =2580
    Top =3024
    Right =19548
    Bottom =13104
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x90faff9076ebe340
    End
    Caption ="Remarketing Menu"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
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
        Begin Line
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
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            Height =720
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Left =45
                    Top =540
                    Width =9648
                    Height =180
                    Name ="Image142"
                    Picture ="3_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedLeft =45
                    LayoutCachedTop =540
                    LayoutCachedWidth =9693
                    LayoutCachedHeight =720
                    TabIndex =2
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =8115
                    Top =180
                    Width =750
                    Height =228
                    FontSize =8
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =8115
                    LayoutCachedTop =180
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =408
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =9
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =7560
                    Top =180
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =1
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="FAMainMenu"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"GoHome\" xmlns=\"http://schemas.microsoft.com/office/accesss"
                                "ervices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument Na"
                                "me=\"FormName\">FAMainMenu</Argum"
                        End
                        Begin
                            Comment ="_AXL:ent></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7560
                    LayoutCachedTop =180
                    LayoutCachedWidth =8115
                    LayoutCachedHeight =408
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =3
                    Width =2460
                    Height =690
                    FontSize =18
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="Remarketing"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =690
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =435
                    Top =375
                    Width =1815
                    Height =240
                    FontSize =8
                    Name ="Label227"
                    Caption =" MENU"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =435
                    LayoutCachedTop =375
                    LayoutCachedWidth =2250
                    LayoutCachedHeight =615
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            Height =6840
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =120
                    Width =2700
                    Height =330
                    Name ="MaintInvoiceRequest"
                    Caption ="Buyers Form && Orders"
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =450
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2964
                    Top =2256
                    Width =2652
                    Height =336
                    TabIndex =13
                    ForeColor =0
                    Name ="ScheduleControl"
                    Caption ="  Schedule Control"
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2964
                    LayoutCachedTop =2256
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =2592
                    Alignment =1
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =85
                    Left =240
                    Top =2760
                    Width =8640
                    BorderColor =8872517
                    Name ="Line48"
                    LayoutCachedLeft =240
                    LayoutCachedTop =2760
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =2760
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5760
                    Top =816
                    Width =2700
                    Height =336
                    TabIndex =5
                    ForeColor =0
                    Name ="MakeBankPayOffEmail"
                    Caption ="  Make Bank Pay Off Email"
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffb3b3b3ffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffb17d4affb17d4affffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffb17d4affb17d4affffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5760
                    LayoutCachedTop =816
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =1152
                    PictureCaptionArrangement =5
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =1296
                    Width =2700
                    Height =336
                    TabIndex =6
                    ForeColor =0
                    Name ="OpenDashBoardSales"
                    Caption ="  Dashboard - Units Sold"
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =120
                    LayoutCachedTop =1296
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =1632
                    Alignment =1
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =180
                    Top =4740
                    Width =420
                    Height =330
                    TabIndex =34
                    Name ="ClientId"

                    LayoutCachedLeft =180
                    LayoutCachedTop =4740
                    LayoutCachedWidth =600
                    LayoutCachedHeight =5070
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =816
                    Width =2700
                    Height =336
                    TabIndex =3
                    ForeColor =0
                    Name ="OpenDashBoardInv"
                    Caption ="  Dashboard- Inv (Sales)"
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =120
                    LayoutCachedTop =816
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =1152
                    Alignment =1
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5760
                    Top =120
                    Width =2700
                    Height =330
                    TabIndex =2
                    ForeColor =0
                    Name ="OpenRMKTMassMod"
                    Caption ="Mass Mod (RMKT)"
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5760
                    LayoutCachedTop =120
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =450
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    ListWidth =2880
                    Left =780
                    Top =5460
                    Width =975
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =35
                    BackColor =15788753
                    BorderColor =4754549
                    Name ="OverRideCbo"
                    RowSourceType ="Value List"
                    RowSource ="\"No Override\";\"Manager\";\"Sales\""
                    ColumnWidths ="2160"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    AllowValueListEdits =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =5460
                    LayoutCachedWidth =1755
                    LayoutCachedHeight =5730
                    RowStart =7
                    RowEnd =7
                    ColumnStart =4
                    ColumnEnd =4
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =780
                            Top =5040
                            Width =960
                            Height =420
                            FontSize =8
                            FontWeight =700
                            ForeColor =10040879
                            Name ="Label222"
                            Caption ="Access \015\012Over Ride:"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120
                            LayoutCachedLeft =780
                            LayoutCachedTop =5040
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =5460
                            RowStart =7
                            RowEnd =7
                            ColumnStart =2
                            ColumnEnd =3
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =1776
                    Width =2700
                    Height =336
                    TabIndex =9
                    ForeColor =0
                    Name ="OpenDashBoardAdmin"
                    Caption ="  Dashboard - Admin "
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =120
                    LayoutCachedTop =1776
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =2112
                    Alignment =1
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =13105149
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2964
                    Top =1296
                    Width =2652
                    Height =336
                    TabIndex =7
                    ForeColor =0
                    Name ="OpenDashBoardAcct"
                    Caption ="  Dashboard - Accounting"
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2964
                    LayoutCachedTop =1296
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =1632
                    Alignment =1
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2964
                    Top =1776
                    Width =2652
                    Height =336
                    TabIndex =10
                    ForeColor =0
                    Name ="SPDashboard"
                    Caption ="  SP Dashboard (CRM)"
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =2964
                    LayoutCachedTop =1776
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =2112
                    Alignment =1
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =2256
                    Width =2700
                    Height =336
                    TabIndex =12
                    ForeColor =0
                    Name ="OpendashBoardInvMgr"
                    Caption ="  Dashboard - Inv (MGR)"
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =120
                    LayoutCachedTop =2256
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =2592
                    Alignment =1
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =5760
                    Top =576
                    Width =2700
                    Height =216
                    Name ="EmptyCell327"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =0
                    RightPadding =75
                    BottomPadding =0
                    LayoutCachedLeft =5760
                    LayoutCachedTop =576
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =792
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =23
                End
                Begin EmptyCell
                    Left =2964
                    Top =576
                    Width =2652
                    Height =216
                    Name ="EmptyCell333"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =0
                    RightPadding =75
                    BottomPadding =0
                    LayoutCachedLeft =2964
                    LayoutCachedTop =576
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =792
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =23
                End
                Begin EmptyCell
                    Left =120
                    Top =576
                    Width =2700
                    Height =216
                    Name ="EmptyCell362"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =0
                    RightPadding =75
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =576
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =792
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =23
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =780
                    Top =4200
                    Width =975
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =33
                    BackColor =14151142
                    BorderColor =4754549
                    Name ="EndDate"
                    Format ="Short Date"
                    Tag ="Manager"

                    LayoutCachedLeft =780
                    LayoutCachedTop =4200
                    LayoutCachedWidth =1755
                    LayoutCachedHeight =4470
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =4200
                            Width =645
                            Height =240
                            FontSize =8
                            FontWeight =700
                            ForeColor =10040879
                            Name ="Label560"
                            Caption ="End  ►"
                            LayoutCachedLeft =60
                            LayoutCachedTop =4200
                            LayoutCachedWidth =705
                            LayoutCachedHeight =4440
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2160
                    Top =3720
                    Width =2880
                    Height =274
                    FontSize =10
                    TabIndex =19
                    ForeColor =0
                    Name ="MgtReport"
                    Caption ="  Quick Inventory Report"
                    OnClick ="[Event Procedure]"
                    Tag ="Manager"
                    GroupTable =26
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =2160
                    LayoutCachedTop =3720
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =3994
                    PictureCaptionArrangement =5
                    Alignment =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =26
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5760
                    Top =2256
                    Width =2700
                    Height =336
                    TabIndex =14
                    ForeColor =0
                    Name =" UpdateRMKTData"
                    Caption =" Update RMKT From Server"
                    OnClick ="[Event Procedure]"
                    EventProcPrefix ="Ctl_UpdateRMKTData"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5760
                    LayoutCachedTop =2256
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =2592
                    Alignment =1
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =24
                    ListWidth =4320
                    Left =2820
                    Top =3225
                    Width =1605
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =16
                    BorderColor =4754549
                    Name ="Vendor"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName FROM Vendors INNER JOIN RemarketExpe"
                        "nse ON Vendors.VendorID = RemarketExpense.VendorId GROUP BY Vendors.VendorID, Ve"
                        "ndors.VendorName ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;4320"
                    Tag ="Manager"

                    LayoutCachedLeft =2820
                    LayoutCachedTop =3225
                    LayoutCachedWidth =4425
                    LayoutCachedHeight =3495
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2820
                            Top =3000
                            Width =1380
                            Height =240
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label551"
                            Caption ="┌─── Vendor ───┐ "
                            Tag ="Manager"
                            LayoutCachedLeft =2820
                            LayoutCachedTop =3000
                            LayoutCachedWidth =4200
                            LayoutCachedHeight =3240
                            BackThemeColorIndex =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2160
                    Top =4152
                    Width =2880
                    Height =276
                    FontSize =10
                    TabIndex =21
                    ForeColor =0
                    Name ="SurrenderPDReport"
                    Caption ="Surrender / Hold Over Per Diem"
                    OnClick ="[Event Procedure]"
                    GroupTable =26
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =2160
                    LayoutCachedTop =4152
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =4428
                    PictureCaptionArrangement =5
                    Alignment =1
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =26
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5760
                    Top =1776
                    Width =2700
                    Height =336
                    TabIndex =11
                    ForeColor =0
                    Name ="OutstandingPayOffs"
                    Caption ="  Outstanding Pay Offs"
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
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

                    LayoutCachedLeft =5760
                    LayoutCachedTop =1776
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =2112
                    PictureCaptionArrangement =5
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5412
                    Top =3720
                    Width =3036
                    Height =274
                    FontSize =10
                    TabIndex =20
                    ForeColor =0
                    Name ="LookUpRMKTInvoice"
                    Caption ="  Look Up RMKT Expense INVOICE"
                    OnClick ="[Event Procedure]"
                    Tag ="Manager"
                    GroupTable =26
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5412
                    LayoutCachedTop =3720
                    LayoutCachedWidth =8448
                    LayoutCachedHeight =3994
                    PictureCaptionArrangement =5
                    Alignment =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =26
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4485
                    Top =3225
                    Width =1560
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =17
                    BorderColor =4754549
                    Name ="InvoiceNum"
                    Tag ="Manager"

                    LayoutCachedLeft =4485
                    LayoutCachedTop =3225
                    LayoutCachedWidth =6045
                    LayoutCachedHeight =3495
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =4485
                            Top =3000
                            Width =1560
                            Height =240
                            FontSize =8
                            ForeColor =10040879
                            Name ="Label552"
                            Caption ="┌─── Invoice Num ──┐"
                            Tag ="Manager"
                            LayoutCachedLeft =4485
                            LayoutCachedTop =3000
                            LayoutCachedWidth =6045
                            LayoutCachedHeight =3240
                            BackThemeColorIndex =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5412
                    Top =4152
                    Width =3036
                    Height =276
                    FontSize =10
                    TabIndex =22
                    ForeColor =0
                    Name ="RMKTVendorReport"
                    Caption =" Look Up RMKT Expense VENDOR"
                    OnClick ="[Event Procedure]"
                    Tag ="Manager"
                    GroupTable =26
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5412
                    LayoutCachedTop =4152
                    LayoutCachedWidth =8448
                    LayoutCachedHeight =4428
                    PictureCaptionArrangement =5
                    Alignment =1
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =26
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5412
                    Top =4584
                    Width =3036
                    Height =276
                    FontSize =10
                    TabIndex =24
                    ForeColor =0
                    Name ="RMKTExpenseDate_Report"
                    Caption =" Look Up RMKT Expense DATE"
                    OnClick ="[Event Procedure]"
                    Tag ="Manager"
                    GroupTable =26
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5412
                    LayoutCachedTop =4584
                    LayoutCachedWidth =8448
                    LayoutCachedHeight =4860
                    PictureCaptionArrangement =5
                    Alignment =1
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =26
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =1140
                    Top =3180
                    Width =1080
                    Height =240
                    FontSize =8
                    FontWeight =700
                    ForeColor =10040879
                    Name ="ClearAll"
                    Caption ="▼ Clear All  ►"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =1140
                    LayoutCachedTop =3180
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =3420
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =24
                    Left =6105
                    Top =3225
                    Width =1410
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =18
                    BorderColor =4754549
                    Name ="Bank"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.BankID, vw_SixKeys.BankShortName FROM vw_SixKeys GROUP BY vw_S"
                        "ixKeys.BankID, vw_SixKeys.BankShortName HAVING (((vw_SixKeys.BankShortName)<>\"T"
                        "BD\")) ORDER BY vw_SixKeys.BankShortName;"
                    ColumnWidths ="0;4320"
                    Tag ="Manager"

                    LayoutCachedLeft =6105
                    LayoutCachedTop =3225
                    LayoutCachedWidth =7515
                    LayoutCachedHeight =3495
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =6120
                            Top =3000
                            Width =1380
                            Height =240
                            FontSize =8
                            ForeColor =10040879
                            Name ="RangeLabel"
                            Caption ="   ┌─── Bank ──┐"
                            Tag ="Manager"
                            LayoutCachedLeft =6120
                            LayoutCachedTop =3000
                            LayoutCachedWidth =7500
                            LayoutCachedHeight =3240
                            BackThemeColorIndex =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5412
                    Top =5016
                    Width =3036
                    Height =276
                    FontSize =10
                    TabIndex =26
                    ForeColor =0
                    Name ="RMKTBankReport"
                    Caption =" Look Up RMKT Expense BANK"
                    OnClick ="[Event Procedure]"
                    Tag ="Manager"
                    GroupTable =26
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5412
                    LayoutCachedTop =5016
                    LayoutCachedWidth =8448
                    LayoutCachedHeight =5292
                    PictureCaptionArrangement =5
                    Alignment =1
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =26
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2160
                    Top =4584
                    Width =2880
                    Height =276
                    TabIndex =23
                    ForeColor =0
                    Name ="OpenRmktExpenseManager"
                    Caption =" RMKT Expense Manager"
                    OnClick ="[Event Procedure]"
                    GroupTable =26
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff000000000000000090a4681b90a468ff90a468ff90a46848 ,
                        0x00000000ffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff000000000000000090a4682190a4682d90a468ff90a468ff ,
                        0x90a46848ffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff90a4688d90a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468ff00000000ffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff90a468f390a468840000000090a4684290a468ff90a468ff ,
                        0x90a4682dffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff90a468fc0000000090a4682d90a468ff90a468ff90a4682d ,
                        0x00000000ffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff000000000000000000000000000000000000000000000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff00000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffff0000000000000000000000000000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffff0000000090a4682d90a468ff90a468ff90a4682d ,
                        0x0000000090a468fc727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffff90a4682d90a468ff90a468ff90a4684200000000 ,
                        0x90a4688490a468f3727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xb3b3b3ffffffffff0000000090a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468ff90a4688a727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272cf7272720f90a4684890a468ff90a468ff90a4682d90a46818 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272c67272720990a4684890a468ff90a468ff90a4681b ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272bd00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =2160
                    LayoutCachedTop =4584
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =4860
                    PictureCaptionArrangement =5
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =26
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5412
                    Top =5448
                    Width =3036
                    Height =336
                    FontSize =10
                    TabIndex =28
                    ForeColor =0
                    Name ="Outstanding SUFs"
                    Caption ="   Outstanding SUFs"
                    OnClick ="[Event Procedure]"
                    Tag ="Manager"
                    EventProcPrefix ="Outstanding_SUFs"
                    GroupTable =26
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5412
                    LayoutCachedTop =5448
                    LayoutCachedWidth =8448
                    LayoutCachedHeight =5784
                    PictureCaptionArrangement =5
                    Alignment =1
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =26
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2160
                    Top =5016
                    Width =2880
                    Height =276
                    FontSize =10
                    TabIndex =25
                    ForeColor =0
                    Name ="Exit2Excel"
                    Caption =" EXIT Report"
                    OnClick ="[Event Procedure]"
                    GroupTable =26
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

                    LayoutCachedLeft =2160
                    LayoutCachedTop =5016
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =5292
                    PictureCaptionArrangement =5
                    Alignment =1
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =26
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2160
                    Top =5448
                    Width =2880
                    Height =336
                    FontSize =10
                    TabIndex =27
                    ForeColor =0
                    Name ="SUFReturned"
                    Caption =" SUF Returned  Report"
                    OnClick ="[Event Procedure]"
                    Tag ="Manager"
                    GroupTable =26
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =2160
                    LayoutCachedTop =5448
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =5784
                    PictureCaptionArrangement =5
                    Alignment =1
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =26
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =780
                    Top =3840
                    Width =975
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =32
                    BackColor =15788753
                    BorderColor =4754549
                    Name ="StartDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Manager"

                    LayoutCachedLeft =780
                    LayoutCachedTop =3840
                    LayoutCachedWidth =1755
                    LayoutCachedHeight =4110
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =3840
                            Width =645
                            Height =240
                            FontSize =8
                            FontWeight =700
                            ForeColor =10040879
                            Name ="Label559"
                            Caption ="Start ►"
                            LayoutCachedLeft =60
                            LayoutCachedTop =3840
                            LayoutCachedWidth =705
                            LayoutCachedHeight =4080
                        End
                    End
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =780
                    Top =3600
                    Width =960
                    Height =240
                    FontSize =8
                    ForeColor =10040879
                    Name ="Label558"
                    Caption ="Date Range"
                    Tag ="Manager"
                    LayoutCachedLeft =780
                    LayoutCachedTop =3600
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =3840
                    BackThemeColorIndex =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5760
                    Top =1296
                    Width =2700
                    Height =336
                    TabIndex =8
                    ForeColor =0
                    Name ="MakeSUFEmail"
                    Caption ="  Make SUF Email"
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffb3b3b3ffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffb17d4affb17d4affffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffb17d4affb17d4affffffffff ,
                        0x727272ff0000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff0000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5760
                    LayoutCachedTop =1296
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =1632
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =8580
                    Top =1112
                    Width =1085
                    Height =343
                    TabIndex =15
                    Name ="Emailto"
                    DefaultValue ="2"

                    LayoutCachedLeft =8580
                    LayoutCachedTop =1112
                    LayoutCachedWidth =9665
                    LayoutCachedHeight =1455
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =8700
                            Top =960
                            Width =1005
                            Height =240
                            FontSize =8
                            Name ="Label1435"
                            Caption ="Email goes to"
                            LayoutCachedLeft =8700
                            LayoutCachedTop =960
                            LayoutCachedWidth =9705
                            LayoutCachedHeight =1200
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =8640
                            Top =1215
                            OptionValue =1
                            Name ="Option1437"

                            LayoutCachedLeft =8640
                            LayoutCachedTop =1215
                            LayoutCachedWidth =8900
                            LayoutCachedHeight =1455
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =8870
                                    Top =1217
                                    Width =240
                                    Height =210
                                    FontSize =8
                                    Name ="Label1438"
                                    Caption ="All "
                                    LayoutCachedLeft =8870
                                    LayoutCachedTop =1217
                                    LayoutCachedWidth =9110
                                    LayoutCachedHeight =1427
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =9180
                            Top =1215
                            TabIndex =1
                            OptionValue =2
                            Name ="Option1439"

                            LayoutCachedLeft =9180
                            LayoutCachedTop =1215
                            LayoutCachedWidth =9440
                            LayoutCachedHeight =1455
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =9410
                                    Top =1217
                                    Width =255
                                    Height =210
                                    FontSize =8
                                    Name ="Label1440"
                                    Caption ="Me"
                                    LayoutCachedLeft =9410
                                    LayoutCachedTop =1217
                                    LayoutCachedWidth =9665
                                    LayoutCachedHeight =1427
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2964
                    Top =816
                    Width =2652
                    Height =336
                    TabIndex =4
                    ForeColor =0
                    Name ="OpenUnitMoverManager"
                    Caption =" Move / Driveaway Orders"
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffb17d4affffffffffffffffffffffffffffffffff ,
                        0xffffffffb17d4aff727272ffffffffffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffffffffffb17d4affffffffffffffffff727272ffffffffff ,
                        0xffffffffb17d4aff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffb17d4affffffffff727272ff727272ff727272ff ,
                        0xffffffffb17d4aff727272ffffffffffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffffffffffb17d4affffffffffffffffffffffffffffffffff ,
                        0xffffffffb17d4aff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff727272ffffffffffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffffffffffb17d4affffffffffffffffffffffffffffffffff ,
                        0xffffffffb17d4aff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffb17d4affffffffff727272ff727272ff727272ff ,
                        0xffffffffb17d4aff727272ffffffffffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffffffffffb17d4affffffffffffffffff727272ffffffffff ,
                        0xffffffffb17d4aff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffb17d4affffffffffffffffffffffffffffffffff ,
                        0xffffffffb17d4aff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =2964
                    LayoutCachedTop =816
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =1152
                    PictureCaptionArrangement =5
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2964
                    Top =120
                    Width =2652
                    Height =330
                    TabIndex =1
                    ForeColor =0
                    Name ="BuyersOrdersMassMod"
                    Caption ="Mass Mod BuyersOrders"
                    OnClick ="[Event Procedure]"
                    GroupTable =23
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2964
                    LayoutCachedTop =120
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =450
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =23
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =2160
                    Top =5940
                    Width =2880
                    Height =336
                    FontSize =10
                    TabIndex =29
                    ForeColor =0
                    Name ="InterimStatusReport"
                    Caption =" Interim Status Report"
                    OnClick ="[Event Procedure]"
                    Tag ="Manager"
                    GroupTable =26
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =2160
                    LayoutCachedTop =5940
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =6276
                    PictureCaptionArrangement =5
                    Alignment =1
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =26
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5412
                    Top =5940
                    Width =3036
                    Height =336
                    FontSize =10
                    TabIndex =30
                    ForeColor =0
                    Name ="MonthlySoldReport"
                    Caption =" Monthly Sold Report"
                    OnClick ="[Event Procedure]"
                    Tag ="Manager"
                    GroupTable =26
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5412
                    LayoutCachedTop =5940
                    LayoutCachedWidth =8448
                    LayoutCachedHeight =6276
                    PictureCaptionArrangement =5
                    Alignment =1
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =26
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =5148
                    Top =3720
                    Width =192
                    Height =274
                    Name ="EmptyCell583"
                    GroupTable =26
                    BottomPadding =120
                    LayoutCachedLeft =5148
                    LayoutCachedTop =3720
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =3994
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =26
                End
                Begin EmptyCell
                    Left =5148
                    Top =4152
                    Width =192
                    Height =276
                    Name ="EmptyCell584"
                    GroupTable =26
                    BottomPadding =120
                    LayoutCachedLeft =5148
                    LayoutCachedTop =4152
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =4428
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =26
                End
                Begin EmptyCell
                    Left =5148
                    Top =4584
                    Width =192
                    Height =276
                    Name ="EmptyCell585"
                    GroupTable =26
                    BottomPadding =120
                    LayoutCachedLeft =5148
                    LayoutCachedTop =4584
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =4860
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =26
                End
                Begin EmptyCell
                    Left =5148
                    Top =5016
                    Width =192
                    Height =276
                    Name ="EmptyCell586"
                    GroupTable =26
                    BottomPadding =120
                    LayoutCachedLeft =5148
                    LayoutCachedTop =5016
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =5292
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =26
                End
                Begin EmptyCell
                    Left =5148
                    Top =5448
                    Width =192
                    Height =336
                    Name ="EmptyCell587"
                    GroupTable =26
                    BottomPadding =120
                    LayoutCachedLeft =5148
                    LayoutCachedTop =5448
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =5784
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =26
                End
                Begin EmptyCell
                    Left =5148
                    Top =5940
                    Width =192
                    Height =336
                    Name ="EmptyCell588"
                    GroupTable =26
                    BottomPadding =120
                    LayoutCachedLeft =5148
                    LayoutCachedTop =5940
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =6276
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =26
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =2160
                    Top =6336
                    Width =2880
                    Height =336
                    FontSize =10
                    TabIndex =36
                    ForeColor =0
                    Name ="BankInventoryReport"
                    Caption =" Bank Inv and Sale Report"
                    OnClick ="[Event Procedure]"
                    Tag ="Manager"
                    GroupTable =27
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =2160
                    LayoutCachedTop =6336
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =6672
                    PictureCaptionArrangement =5
                    Alignment =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =27
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5412
                    Top =6432
                    Width =3036
                    Height =276
                    TabIndex =31
                    ForeColor =0
                    Name ="BTF_BuyersOrder"
                    Caption =" Buy Out Transaction"
                    OnClick ="[Event Procedure]"
                    GroupTable =26
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff000000000000000090a4681b90a468ff90a468ff90a46848 ,
                        0x00000000ffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff000000000000000090a4682190a4682d90a468ff90a468ff ,
                        0x90a46848ffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff90a4688d90a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468ff00000000ffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff90a468f390a468840000000090a4684290a468ff90a468ff ,
                        0x90a4682dffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff90a468fc0000000090a4682d90a468ff90a468ff90a4682d ,
                        0x00000000ffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff000000000000000000000000000000000000000000000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff00000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffff0000000000000000000000000000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffff0000000090a4682d90a468ff90a468ff90a4682d ,
                        0x0000000090a468fc727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffff90a4682d90a468ff90a468ff90a4684200000000 ,
                        0x90a4688490a468f3727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xb3b3b3ffffffffff0000000090a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468ff90a4688a727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272cf7272720f90a4684890a468ff90a468ff90a4682d90a46818 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272c67272720990a4684890a468ff90a468ff90a4681b ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272bd00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5412
                    LayoutCachedTop =6432
                    LayoutCachedWidth =8448
                    LayoutCachedHeight =6708
                    PictureCaptionArrangement =5
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    GroupTable =26
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                End
                Begin EmptyCell
                    Left =2160
                    Top =6432
                    Width =2880
                    Height =276
                    Name ="EmptyCell592"
                    GroupTable =26
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    LayoutCachedLeft =2160
                    LayoutCachedTop =6432
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =6708
                    RowStart =6
                    RowEnd =6
                    LayoutGroup =2
                    GroupTable =26
                End
                Begin EmptyCell
                    Left =5148
                    Top =6432
                    Width =192
                    Height =276
                    Name ="EmptyCell593"
                    GroupTable =26
                    BottomPadding =120
                    LayoutCachedLeft =5148
                    LayoutCachedTop =6432
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =6708
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =26
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8820
                    Top =180
                    TabIndex =37
                    Name ="Command595"
                    Caption ="Command595"

                    LayoutCachedLeft =8820
                    LayoutCachedTop =180
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =540
                End
                Begin Label
                    OverlapFlags =85
                    Left =9060
                    Top =2280
                    Width =1212
                    Height =300
                    Name ="Label600"
                    Caption ="git dev 1 test"
                    LayoutCachedLeft =9060
                    LayoutCachedTop =2280
                    LayoutCachedWidth =10272
                    LayoutCachedHeight =2580
                End
                Begin Label
                    OverlapFlags =85
                    Left =9000
                    Top =3180
                    Width =1260
                    Height =300
                    Name ="Label601"
                    Caption ="git dev 2  test"
                    LayoutCachedLeft =9000
                    LayoutCachedTop =3180
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =3480
                End
                Begin Label
                    OverlapFlags =85
                    Left =9060
                    Top =3960
                    Width =1920
                    Height =300
                    Name ="Label602"
                    Caption ="change from Gnstwo"
                    LayoutCachedLeft =9060
                    LayoutCachedTop =3960
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =4260
                End
                Begin Label
                    OverlapFlags =85
                    Left =9120
                    Top =4560
                    Width =1884
                    Height =300
                    Name ="Label603"
                    Caption ="change from gnsone"
                    LayoutCachedLeft =9120
                    LayoutCachedTop =4560
                    LayoutCachedWidth =11004
                    LayoutCachedHeight =4860
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "Rmkt Menu.cls"
