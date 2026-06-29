Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7488
    DatasheetFontHeight =11
    ItemSuffix =207
    Left =5670
    Top =3315
    Right =13410
    Bottom =11130
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x62fc97d5f1cfe540
    End
    Caption ="Draw Down Add/Edit"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin FormHeader
            Height =630
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =6240
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
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

                    LayoutCachedLeft =6240
                    LayoutCachedTop =120
                    LayoutCachedWidth =6795
                    LayoutCachedHeight =348
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
                Begin CommandButton
                    OverlapFlags =223
                    PictureType =1
                    Left =6807
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =1
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =6807
                    LayoutCachedTop =120
                    LayoutCachedWidth =7362
                    LayoutCachedHeight =348
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
                Begin Label
                    OverlapFlags =93
                    Left =60
                    Width =3660
                    Height =480
                    FontSize =16
                    FontWeight =700
                    Name ="Header"
                    Caption ="Draw Down Add/Edit"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =9
                    ForeShade =75.0
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =420
                    Width =7488
                    Height =210
                    BorderColor =16777215
                    Name ="Image767"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =420
                    LayoutCachedWidth =7488
                    LayoutCachedHeight =630
                    TabIndex =2
                    BorderShade =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7200
            Name ="Detail"
            AlternateBackColor =14545355
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5808
                    Top =5760
                    FontSize =10
                    TabIndex =14
                    BackColor =11271167
                    Name ="CreatedOn"
                    Format ="Short Date"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Note Date is required"
                    DefaultValue ="=Date()"
                    Tag ="Ditem"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5808
                    LayoutCachedTop =5760
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =6000
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =4
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3900
                    Top =6660
                    Width =1080
                    Height =270
                    FontSize =10
                    TabIndex =15
                    BorderColor =14211288
                    Name ="DrawDownId"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3900
                    LayoutCachedTop =6660
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =6930
                    RowStart =3
                    RowEnd =3
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =5628
                    Top =3288
                    Width =1416
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =9
                    Name ="AddDrawBtn"
                    Caption ="Add Draw"
                    OnClick ="[Event Procedure]"
                    GroupTable =2

                    LayoutCachedLeft =5628
                    LayoutCachedTop =3288
                    LayoutCachedWidth =7044
                    LayoutCachedHeight =3564
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =48840
                    BorderThemeColorIndex =1
                    HoverColor =48840
                    PressedColor =48840
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =1140
                    Top =2376
                    Width =315
                    Height =2556
                    FontSize =10
                    Name ="Label1744"
                    Caption =" └─  Draw Details  ─┘"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =1140
                    LayoutCachedTop =2376
                    LayoutCachedWidth =1455
                    LayoutCachedHeight =4932
                    RowEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =9
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =5628
                    Top =2340
                    Width =1416
                    Height =540
                    FontSize =9
                    FontWeight =700
                    TabIndex =8
                    Name ="SaveMe"
                    Caption ="Save This\015\012 Draw"
                    OnClick ="[Event Procedure]"
                    GroupTable =2

                    LayoutCachedLeft =5628
                    LayoutCachedTop =2340
                    LayoutCachedWidth =7044
                    LayoutCachedHeight =2880
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BackTint =60.0
                    BorderThemeColorIndex =1
                    HoverColor =48840
                    PressedColor =48840
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4368
                    Top =5760
                    FontSize =10
                    TabIndex =13
                    BackColor =11271167
                    Name ="CreatedBy"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Intials of person entering note are required"
                    Tag ="Ditem"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4368
                    LayoutCachedTop =5760
                    LayoutCachedWidth =5808
                    LayoutCachedHeight =6000
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5880
                    Top =6660
                    Width =600
                    Height =270
                    FontSize =10
                    TabIndex =17
                    BorderColor =14211288
                    Name ="SchID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5880
                    LayoutCachedTop =6660
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =6930
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6480
                    Top =6660
                    Width =660
                    Height =270
                    FontSize =10
                    TabIndex =18
                    BorderColor =14211288
                    Name ="GroupID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6480
                    LayoutCachedTop =6660
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =6930
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4980
                    Top =6660
                    Width =900
                    Height =270
                    FontSize =10
                    TabIndex =16
                    BorderColor =14211288
                    Name ="DrawTypeId"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4980
                    LayoutCachedTop =6660
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =6930
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =180
                    Top =1740
                    Width =1152
                    Height =450
                    FontSize =18
                    FontWeight =700
                    BorderColor =-2147483615
                    Name ="DrawNumber"
                    Tag ="Ditem"

                    LayoutCachedLeft =180
                    LayoutCachedTop =1740
                    LayoutCachedWidth =1332
                    LayoutCachedHeight =2190
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =180
                            Top =1500
                            Width =1152
                            Height =270
                            FontSize =10
                            ForeColor =16777215
                            Name ="Label13"
                            Caption ="Draw"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1500
                            LayoutCachedWidth =1332
                            LayoutCachedHeight =1770
                            BackThemeColorIndex =6
                            BackShade =50.0
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2820
                    Top =3552
                    Width =2232
                    Height =336
                    TabIndex =4
                    BorderColor =-2147483615
                    Name ="DrawBankAccountNumber"
                    Tag ="Ditem"
                    GroupTable =6

                    LayoutCachedLeft =2820
                    LayoutCachedTop =3552
                    LayoutCachedWidth =5052
                    LayoutCachedHeight =3888
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    ForeThemeColorIndex =0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1596
                            Top =3552
                            Width =1152
                            Height =336
                            FontSize =10
                            FontWeight =700
                            BorderColor =-2147483615
                            Name ="Label15"
                            Caption ="Account Num"
                            GroupTable =6
                            LayoutCachedLeft =1596
                            LayoutCachedTop =3552
                            LayoutCachedWidth =2748
                            LayoutCachedHeight =3888
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2820
                    Top =4368
                    Width =2232
                    Height =336
                    TabIndex =6
                    BorderColor =-2147483615
                    Name ="DateDrawn"
                    Format ="Short Date"
                    Tag ="Ditem"
                    GroupTable =6

                    LayoutCachedLeft =2820
                    LayoutCachedTop =4368
                    LayoutCachedWidth =5052
                    LayoutCachedHeight =4704
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    BackThemeColorIndex =2
                    BackTint =20.0
                    ForeThemeColorIndex =0
                    GroupTable =6
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =1596
                            Top =4368
                            Width =1152
                            Height =336
                            FontSize =10
                            FontWeight =700
                            BorderColor =-2147483615
                            Name ="Label16"
                            Caption ="Date Drawn"
                            GroupTable =6
                            LayoutCachedLeft =1596
                            LayoutCachedTop =4368
                            LayoutCachedWidth =2748
                            LayoutCachedHeight =4704
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =3
                            BackThemeColorIndex =2
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =5760
                    FontSize =10
                    TabIndex =10
                    BorderColor =-2147483615
                    Name ="UnitDrawCount"
                    Tag ="RequestDetails"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =240
                    LayoutCachedTop =5760
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =6000
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =4
                    BackThemeColorIndex =6
                    BackTint =40.0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1680
                    Top =5760
                    FontSize =10
                    TabIndex =11
                    BackColor =15858167
                    BorderColor =-2147483615
                    Name ="ChildDrawCount"
                    Tag ="RequestDetails"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1680
                    LayoutCachedTop =5760
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =6000
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3120
                    Top =5760
                    Width =1245
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    BorderColor =-2147483615
                    Name ="fStatus"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="RequestDetails"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3120
                    LayoutCachedTop =5760
                    LayoutCachedWidth =4365
                    LayoutCachedHeight =6000
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2820
                    Top =3960
                    Width =2232
                    Height =336
                    TabIndex =5
                    BorderColor =-2147483615
                    Name ="DrawCreditLimit"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Ditem"
                    GroupTable =6

                    LayoutCachedLeft =2820
                    LayoutCachedTop =3960
                    LayoutCachedWidth =5052
                    LayoutCachedHeight =4296
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1596
                            Top =3960
                            Width =1152
                            Height =336
                            FontSize =10
                            FontWeight =700
                            BorderColor =-2147483615
                            Name ="Label350"
                            Caption ="Credit Limit"
                            GroupTable =6
                            LayoutCachedLeft =1596
                            LayoutCachedTop =3960
                            LayoutCachedWidth =2748
                            LayoutCachedHeight =4296
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2820
                    Top =4776
                    Width =2232
                    Height =336
                    TabIndex =7
                    BorderColor =-2147483615
                    Name ="DrawPaidOff"
                    Format ="Short Date"
                    Tag ="Ditem"
                    GroupTable =6

                    LayoutCachedLeft =2820
                    LayoutCachedTop =4776
                    LayoutCachedWidth =5052
                    LayoutCachedHeight =5112
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    BackShade =95.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =1596
                            Top =4776
                            Width =1152
                            Height =336
                            FontSize =10
                            FontWeight =700
                            BorderColor =-2147483615
                            Name ="Label356"
                            Caption ="Paid Off"
                            GroupTable =6
                            LayoutCachedLeft =1596
                            LayoutCachedTop =4776
                            LayoutCachedWidth =2748
                            LayoutCachedHeight =5112
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =3
                            BackThemeColorIndex =0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2820
                    Top =3144
                    Width =2232
                    Height =336
                    TabIndex =3
                    BorderColor =-2147483615
                    Name ="DrawRequestNumber"
                    Tag ="Ditem"
                    GroupTable =6

                    LayoutCachedLeft =2820
                    LayoutCachedTop =3144
                    LayoutCachedWidth =5052
                    LayoutCachedHeight =3480
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1596
                            Top =3144
                            Width =1152
                            Height =336
                            FontSize =10
                            FontWeight =700
                            BorderColor =-2147483615
                            Name ="Label362"
                            Caption ="Request Num"
                            GroupTable =6
                            LayoutCachedLeft =1596
                            LayoutCachedTop =3144
                            LayoutCachedWidth =2748
                            LayoutCachedHeight =3480
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1560
                    Top =1740
                    Width =5577
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    BorderColor =-2147483615
                    Name ="DrawName"
                    Tag ="Ditem"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =1740
                    LayoutCachedWidth =7137
                    LayoutCachedHeight =2190
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =1560
                            Top =1500
                            Width =5577
                            Height =270
                            FontSize =10
                            BackColor =8210719
                            BorderColor =-2147483615
                            ForeColor =16777215
                            Name ="Label44"
                            Caption ="Draw Name"
                            LayoutCachedLeft =1560
                            LayoutCachedTop =1500
                            LayoutCachedWidth =7137
                            LayoutCachedHeight =1770
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =2820
                    Top =2340
                    Width =2232
                    Height =330
                    TabIndex =2
                    Name ="DrawBankId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankID, Banks.BankShortName, Banks.BankType FROM Banks ORDER BY Ban"
                        "ks.BankShortName;"
                    ColumnWidths ="0;1440;0"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Ditem"
                    GroupTable =6

                    LayoutCachedLeft =2820
                    LayoutCachedTop =2340
                    LayoutCachedWidth =5052
                    LayoutCachedHeight =2670
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    ThemeFontIndex =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =1596
                            Top =2340
                            Width =1152
                            Height =330
                            FontSize =10
                            FontWeight =700
                            Name ="Label53"
                            Caption ="Interim Bank"
                            GroupTable =6
                            LayoutCachedLeft =1596
                            LayoutCachedTop =2340
                            LayoutCachedWidth =2748
                            LayoutCachedHeight =2670
                            LayoutGroup =3
                            BackThemeColorIndex =6
                            BackShade =50.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin EmptyCell
                    Left =7104
                    Top =2940
                    Width =168
                    Height =270
                    Name ="EmptyCell64"
                    GroupTable =2
                    LayoutCachedLeft =7104
                    LayoutCachedTop =2940
                    LayoutCachedWidth =7272
                    LayoutCachedHeight =3210
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =7104
                    Top =3288
                    Width =168
                    Height =276
                    Name ="EmptyCell85"
                    GroupTable =2
                    LayoutCachedLeft =7104
                    LayoutCachedTop =3288
                    LayoutCachedWidth =7272
                    LayoutCachedHeight =3564
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =7104
                    Top =2340
                    Width =168
                    Height =540
                    Name ="EmptyCell86"
                    GroupTable =2
                    LayoutCachedLeft =7104
                    LayoutCachedTop =2340
                    LayoutCachedWidth =7272
                    LayoutCachedHeight =2880
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =6480
                    Top =6420
                    Width =660
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label150"
                    Caption ="GroupID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6480
                    LayoutCachedTop =6420
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =6690
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4368
                    Top =5520
                    Width =1440
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label156"
                    Caption ="Created By"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4368
                    LayoutCachedTop =5520
                    LayoutCachedWidth =5808
                    LayoutCachedHeight =5760
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin EmptyCell
                    Left =240
                    Top =5520
                    Name ="EmptyCell158"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =240
                    LayoutCachedTop =5520
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =5760
                    LayoutGroup =4
                    GroupTable =9
                End
                Begin EmptyCell
                    Left =1680
                    Top =5520
                    Name ="EmptyCell159"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1680
                    LayoutCachedTop =5520
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =5760
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    GroupTable =9
                End
                Begin EmptyCell
                    Left =3120
                    Top =5520
                    Width =1245
                    Name ="EmptyCell160"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3120
                    LayoutCachedTop =5520
                    LayoutCachedWidth =4365
                    LayoutCachedHeight =5760
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    GroupTable =9
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5808
                    Top =5520
                    Width =1440
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label163"
                    Caption ="Date Created"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5808
                    LayoutCachedTop =5520
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =5760
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =4980
                    Top =6420
                    Width =900
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label165"
                    Caption ="Draw Type"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4980
                    LayoutCachedTop =6420
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =6690
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =3900
                    Top =6420
                    Width =1080
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label166"
                    Caption ="DrawDownID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3900
                    LayoutCachedTop =6420
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =6690
                    RowStart =2
                    RowEnd =2
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =5880
                    Top =6420
                    Width =600
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label174"
                    Caption ="SchID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5880
                    LayoutCachedTop =6420
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =6690
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin EmptyCell
                    Left =5628
                    Top =2940
                    Width =1416
                    Height =270
                    Name ="EmptyCell180"
                    GroupTable =2
                    LayoutCachedLeft =5628
                    LayoutCachedTop =2940
                    LayoutCachedWidth =7044
                    LayoutCachedHeight =3210
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1500
                    Top =240
                    Width =5577
                    Height =300
                    FontSize =12
                    TabIndex =19
                    BorderColor =-2147483615
                    Name ="Client"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =240
                    LayoutCachedWidth =7077
                    LayoutCachedHeight =540
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =1500
                            Width =555
                            Height =285
                            FontSize =10
                            BackColor =8210719
                            BorderColor =-2147483615
                            Name ="Label193"
                            Caption ="Client"
                            LayoutCachedLeft =1500
                            LayoutCachedWidth =2055
                            LayoutCachedHeight =285
                            ForeThemeColorIndex =2
                            ForeShade =75.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1770
                    Top =6660
                    Width =660
                    Height =270
                    FontSize =10
                    TabIndex =20
                    BorderColor =14211288
                    Name ="ClientId"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1770
                    LayoutCachedTop =6660
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =6930
                    RowStart =3
                    RowEnd =3
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1770
                            Top =6420
                            Width =660
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label204"
                            Caption ="ClientID"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1770
                            LayoutCachedTop =6420
                            LayoutCachedWidth =2430
                            LayoutCachedHeight =6690
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =2
                            ForeThemeColorIndex =0
                            ForeTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2400
                    Top =6660
                    Width =600
                    Height =270
                    FontSize =10
                    TabIndex =21
                    BorderColor =14211288
                    Name ="MLID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2400
                    LayoutCachedTop =6660
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =6930
                    RowStart =3
                    RowEnd =3
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3000
                    Top =6660
                    Width =900
                    Height =270
                    FontSize =10
                    TabIndex =22
                    BorderColor =14211288
                    Name ="FAID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3000
                    LayoutCachedTop =6660
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =6930
                    RowStart =3
                    RowEnd =3
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =127
                            TextAlign =2
                            Left =3000
                            Top =6420
                            Width =900
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label202"
                            Caption ="FAID"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =3000
                            LayoutCachedTop =6420
                            LayoutCachedWidth =3900
                            LayoutCachedHeight =6690
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =2
                            ForeThemeColorIndex =0
                            ForeTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =180
                    Top =240
                    Width =1152
                    Height =300
                    FontSize =12
                    FontWeight =700
                    TabIndex =23
                    BorderColor =-2147483615
                    Name ="Assignee"
                    Tag ="RequestDetails"

                    LayoutCachedLeft =180
                    LayoutCachedTop =240
                    LayoutCachedWidth =1332
                    LayoutCachedHeight =540
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =180
                            Width =1152
                            Height =285
                            FontSize =10
                            ForeColor =16777215
                            Name ="Label198"
                            Caption ="Assignee"
                            LayoutCachedLeft =180
                            LayoutCachedWidth =1332
                            LayoutCachedHeight =285
                            BackThemeColorIndex =4
                            BackShade =50.0
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =828
                    Width =1152
                    FontSize =10
                    TabIndex =24
                    BackColor =11271167
                    Name ="FundDate"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Intials of person entering note are required"
                    GroupTable =10
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =180
                    LayoutCachedTop =828
                    LayoutCachedWidth =1332
                    LayoutCachedHeight =1068
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =5
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =10
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =180
                    Top =540
                    Width =1152
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label200"
                    Caption ="Syndicated"
                    GroupTable =10
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =180
                    LayoutCachedTop =540
                    LayoutCachedWidth =1332
                    LayoutCachedHeight =810
                    LayoutGroup =5
                    BorderThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =10
                End
                Begin Line
                    OverlapFlags =85
                    Left =60
                    Top =1200
                    Width =7200
                    Name ="Line201"
                    LayoutCachedLeft =60
                    LayoutCachedTop =1200
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =1200
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =2400
                    Top =6420
                    Width =600
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label203"
                    Caption ="MLID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2400
                    LayoutCachedTop =6420
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =6690
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin EmptyCell
                    Left =1596
                    Top =2736
                    Width =1152
                    Height =336
                    Name ="EmptyCell205"
                    GroupTable =6
                    GridlineColor =0
                    LayoutCachedLeft =1596
                    LayoutCachedTop =2736
                    LayoutCachedWidth =2748
                    LayoutCachedHeight =3072
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =2820
                    Top =2736
                    Width =2232
                    Height =336
                    Name ="EmptyCell206"
                    GroupTable =6
                    GridlineColor =0
                    LayoutCachedLeft =2820
                    LayoutCachedTop =2736
                    LayoutCachedWidth =5052
                    LayoutCachedHeight =3072
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =6
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
' See "DrawDownAddADraw.cls"
