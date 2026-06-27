Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18000
    DatasheetFontHeight =11
    ItemSuffix =218
    Left =-19680
    Top =5385
    Right =-1425
    Bottom =12255
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x22ba041ead89e440
    End
    RecordSource ="Select * From DrawDowns"
    OnCurrent ="[Event Procedure]"
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
        Begin Tab
            TextFontFamily =18
            FontSize =11
            FontName ="Cambria"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
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
        Begin FormHeader
            Height =540
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =330
                    Width =18000
                    Height =210
                    BorderColor =16777215
                    Name ="Image767"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =330
                    LayoutCachedWidth =18000
                    LayoutCachedHeight =540
                    TabIndex =2
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =247
                    Left =120
                    Width =3660
                    Height =480
                    FontSize =16
                    FontWeight =700
                    Name ="Label24"
                    Caption ="Bank Draw Downs"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =480
                End
                Begin CommandButton
                    OverlapFlags =127
                    Left =15630
                    Top =45
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

                    LayoutCachedLeft =15630
                    LayoutCachedTop =45
                    LayoutCachedWidth =16185
                    LayoutCachedHeight =273
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
                    OverlapFlags =247
                    PictureType =1
                    Left =16197
                    Top =45
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

                    LayoutCachedLeft =16197
                    LayoutCachedTop =45
                    LayoutCachedWidth =16752
                    LayoutCachedHeight =273
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
            End
        End
        Begin Section
            Height =6120
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14580
                    Top =1140
                    Width =1650
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =16
                    Name ="Request2"
                    ControlSource ="DrawRequestNumber"

                    LayoutCachedLeft =14580
                    LayoutCachedTop =1140
                    LayoutCachedWidth =16230
                    LayoutCachedHeight =1410
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ForeThemeColorIndex =0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =14580
                            Top =960
                            Width =1650
                            Height =210
                            FontSize =8
                            Name ="RequestLabel2"
                            Caption ="Request Number"
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =14580
                            LayoutCachedTop =960
                            LayoutCachedWidth =16230
                            LayoutCachedHeight =1170
                            BackThemeColorIndex =2
                            BackShade =75.0
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =11340
                    Top =1740
                    Width =2448
                    Height =270
                    TabIndex =6
                    Name ="MLIDCbo"
                    ControlSource ="MLID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MstrLease.MLID, MstrLease.MLNO FROM MstrLease ORDER BY MstrLease.MLID;"
                    ColumnWidths ="0;1008"
                    Tag ="EditMe"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =11340
                    LayoutCachedTop =1740
                    LayoutCachedWidth =13788
                    LayoutCachedHeight =2010
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =10035
                            Top =1740
                            Width =1305
                            Height =270
                            FontSize =10
                            Name ="Label12"
                            Caption ="Master Lease"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =10035
                            LayoutCachedTop =1740
                            LayoutCachedWidth =11340
                            LayoutCachedHeight =2010
                            RowStart =3
                            RowEnd =3
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15390
                    Top =1440
                    Width =840
                    Height =630
                    FontSize =22
                    FontWeight =700
                    TabIndex =18
                    Name ="DrawNumber"
                    ControlSource ="DrawNumber"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    Tag ="EditMe"

                    LayoutCachedLeft =15390
                    LayoutCachedTop =1440
                    LayoutCachedWidth =16230
                    LayoutCachedHeight =2070
                    ForeThemeColorIndex =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =14580
                            Top =1440
                            Width =780
                            Height =630
                            FontSize =10
                            FontWeight =700
                            TopMargin =72
                            Name ="Label13"
                            Caption ="Draw\015\012Number"
                            LayoutCachedLeft =14580
                            LayoutCachedTop =1440
                            LayoutCachedWidth =15360
                            LayoutCachedHeight =2070
                            BackThemeColorIndex =2
                            BackShade =50.0
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3600
                    Left =11340
                    Top =1200
                    Width =2448
                    Height =270
                    TabIndex =5
                    BackColor =16183539
                    Name ="DrawBankId"
                    ControlSource ="DrawBankId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.bankID, Banks.BankName FROM Banks Order By BankName;"
                    ColumnWidths ="0;3600"
                    Tag ="EditMe"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =11340
                    LayoutCachedTop =1200
                    LayoutCachedWidth =13788
                    LayoutCachedHeight =1470
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =10035
                            Top =1200
                            Width =1305
                            Height =270
                            FontSize =10
                            Name ="Label14"
                            Caption ="Bank"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =10035
                            LayoutCachedTop =1200
                            LayoutCachedWidth =11340
                            LayoutCachedHeight =1470
                            RowStart =1
                            RowEnd =1
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =11340
                    Top =2550
                    Width =2448
                    Height =270
                    TabIndex =8
                    Name ="DrawBankAccountNumber"
                    ControlSource ="DrawBankAccountNumber"
                    Tag ="EditMe"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11340
                    LayoutCachedTop =2550
                    LayoutCachedWidth =13788
                    LayoutCachedHeight =2820
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =10035
                            Top =2550
                            Width =1305
                            Height =270
                            FontSize =10
                            Name ="Label15"
                            Caption ="Account Num"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =10035
                            LayoutCachedTop =2550
                            LayoutCachedWidth =11340
                            LayoutCachedHeight =2820
                            RowStart =5
                            RowEnd =5
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11557
                    Top =5685
                    FontSize =9
                    TabIndex =22
                    Name ="CreatedOn"
                    ControlSource ="CreatedOn"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =11557
                    LayoutCachedTop =5685
                    LayoutCachedWidth =12997
                    LayoutCachedHeight =5925
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =10305
                            Top =5685
                            Width =1200
                            Height =240
                            FontSize =9
                            Name ="Label17"
                            Caption ="Added Date:"
                            BottomPadding =38
                            LayoutCachedLeft =10305
                            LayoutCachedTop =5685
                            LayoutCachedWidth =11505
                            LayoutCachedHeight =5925
                            BorderThemeColorIndex =1
                            BorderShade =65.0
                            ForeThemeColorIndex =1
                            ForeShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14040
                    Top =5685
                    FontSize =9
                    TabIndex =23
                    Name ="CreatedBy"
                    ControlSource ="CreatedBy"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =14040
                    LayoutCachedTop =5685
                    LayoutCachedWidth =15480
                    LayoutCachedHeight =5925
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =13013
                            Top =5685
                            Width =945
                            Height =240
                            FontSize =9
                            Name ="Label18"
                            Caption ="Added By:"
                            BottomPadding =38
                            LayoutCachedLeft =13013
                            LayoutCachedTop =5685
                            LayoutCachedWidth =13958
                            LayoutCachedHeight =5925
                            RowStart =1
                            RowEnd =1
                            BorderThemeColorIndex =1
                            BorderShade =65.0
                            ForeThemeColorIndex =1
                            ForeShade =65.0
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =9
                    Left =3375
                    Top =1080
                    Width =6336
                    Height =4305
                    FontSize =9
                    TabIndex =19
                    BoundColumn =1
                    Name ="DrawDownList"
                    RowSourceType ="Table/Query"
                    RowSource ="Select * From DrawDowns_Active Where BankId=4200 Order BY MLID, DrawRequestNumbe"
                        "r, DrawNumber DESC;"
                    ColumnWidths ="1152;0;864;1440;1152;576;576;576;0"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3375
                    LayoutCachedTop =1080
                    LayoutCachedWidth =9711
                    LayoutCachedHeight =5385
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11340
                    Top =2820
                    Width =2448
                    Height =270
                    TabIndex =10
                    Name ="DrawCreditLimit"
                    ControlSource ="DrawCreditLimit"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="EditMe"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11340
                    LayoutCachedTop =2820
                    LayoutCachedWidth =13788
                    LayoutCachedHeight =3090
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =10035
                            Top =2820
                            Width =1305
                            Height =270
                            FontSize =10
                            Name ="Label28"
                            Caption ="Credit Limit"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =10035
                            LayoutCachedTop =2820
                            LayoutCachedWidth =11340
                            LayoutCachedHeight =3090
                            RowStart =6
                            RowEnd =6
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15225
                    Top =2280
                    Width =2160
                    Height =270
                    TabIndex =12
                    Name ="DrawRequestNumber"
                    ControlSource ="DrawRequestNumber"
                    Tag ="EditMe"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15225
                    LayoutCachedTop =2280
                    LayoutCachedWidth =17385
                    LayoutCachedHeight =2550
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =13920
                            Top =2280
                            Width =1305
                            Height =270
                            FontSize =10
                            Name ="RequestLabel1"
                            Caption ="Request Num"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =13920
                            LayoutCachedTop =2280
                            LayoutCachedWidth =15225
                            LayoutCachedHeight =2550
                            RowStart =7
                            RowEnd =7
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =11340
                    Top =2280
                    Width =2448
                    Height =270
                    TabIndex =9
                    Name ="DrawName"
                    ControlSource ="DrawName"
                    Tag ="EditMe"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11340
                    LayoutCachedTop =2280
                    LayoutCachedWidth =13788
                    LayoutCachedHeight =2550
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =10035
                            Top =2280
                            Width =1305
                            Height =270
                            FontSize =10
                            Name ="Label32"
                            Caption ="Draw  'Name'"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =10035
                            LayoutCachedTop =2280
                            LayoutCachedWidth =11340
                            LayoutCachedHeight =2550
                            RowStart =5
                            RowEnd =5
                            ColumnStart =2
                            ColumnEnd =2
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =87
                    TextFontCharSet =238
                    TextAlign =1
                    Left =10035
                    Top =840
                    Width =4485
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="EditLabel"
                    Caption ="Draw Details (Locked)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10035
                    LayoutCachedTop =840
                    LayoutCachedWidth =14520
                    LayoutCachedHeight =1200
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11340
                    Top =2010
                    Width =2448
                    Height =270
                    TabIndex =7
                    Name ="DrawTypeId"
                    ControlSource ="DrawTypeId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DrawDownType.DrawDownTypeId, DrawDownType.DrawnDownType FROM DrawDownType"
                        ";"
                    ColumnWidths ="0;864"
                    OnClick ="[Event Procedure]"
                    Tag ="EditMe"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11340
                    LayoutCachedTop =2010
                    LayoutCachedWidth =13788
                    LayoutCachedHeight =2280
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =10035
                            Top =2010
                            Width =1305
                            Height =270
                            FontSize =10
                            Name ="Label38"
                            Caption ="Draw Type"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =10035
                            LayoutCachedTop =2010
                            LayoutCachedWidth =11340
                            LayoutCachedHeight =2280
                            RowStart =4
                            RowEnd =4
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15225
                    Top =2550
                    Width =1185
                    Height =270
                    TabIndex =11
                    BackColor =16183539
                    Name ="DateDrawn"
                    ControlSource ="DateDrawn"
                    Tag ="EditMe"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15225
                    LayoutCachedTop =2550
                    LayoutCachedWidth =16410
                    LayoutCachedHeight =2820
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =13920
                            Top =2550
                            Width =1305
                            Height =270
                            FontSize =10
                            Name ="Label40"
                            Caption ="Date Drawn"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =13920
                            LayoutCachedTop =2550
                            LayoutCachedWidth =15225
                            LayoutCachedHeight =2820
                            RowStart =6
                            RowEnd =6
                            ColumnStart =2
                            ColumnEnd =2
                            BackThemeColorIndex =4
                            BackTint =20.0
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7710
                    Top =5460
                    Width =1980
                    Height =270
                    FontSize =9
                    TabIndex =20
                    Name ="AddDrawDown"
                    Caption ="Add A NEW Draw"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add a New Draw"
                    LeftPadding =90
                    RightPadding =105
                    BottomPadding =165
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7710
                    LayoutCachedTop =5460
                    LayoutCachedWidth =9690
                    LayoutCachedHeight =5730
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11340
                    Top =4020
                    Width =3180
                    Height =270
                    FontSize =9
                    TabIndex =15
                    Name ="CloneADraw"
                    Caption ="'Clone' This Draw"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Make a New Draw From the one displayed"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11340
                    LayoutCachedTop =4020
                    LayoutCachedWidth =14520
                    LayoutCachedHeight =4290
                    RowStart =11
                    RowEnd =11
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =30
                    QuickStyleMask =-5
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3375
                    Top =690
                    Width =4245
                    Height =315
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    BorderColor =0
                    Name ="Text42"
                    ControlSource ="=\"Select or Add a Draw\""

                    LayoutCachedLeft =3375
                    LayoutCachedTop =690
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =1005
                    ForeThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11340
                    Top =3690
                    Width =3180
                    Height =270
                    FontSize =9
                    TabIndex =14
                    ForeColor =0
                    Name ="editme"
                    Caption ="Edit This Draw"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Edit the Draw Displayed"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11340
                    LayoutCachedTop =3690
                    LayoutCachedWidth =14520
                    LayoutCachedHeight =3960
                    RowStart =10
                    RowEnd =10
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =6
                    BorderThemeColorIndex =6
                    HoverThemeColorIndex =6
                    HoverTint =80.0
                    PressedThemeColorIndex =6
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeColor =0
                    QuickStyle =32
                    QuickStyleMask =-517
                    Overlaps =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5760
                    Left =11340
                    Top =1470
                    Width =2448
                    Height =270
                    TabIndex =4
                    Name ="ClientCbo"
                    ControlSource ="ClientID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT clientID, ClientShortName AS Client FROM vw_SixKeys WHERE MLOrig='FA' GRO"
                        "UP BY clientID, ClientShortName ORDER BY ClientShortName;"
                    ColumnWidths ="0;2880"
                    OnClick ="[Event Procedure]"
                    Tag ="EditMe"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =11340
                    LayoutCachedTop =1470
                    LayoutCachedWidth =13788
                    LayoutCachedHeight =1740
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =10035
                            Top =1470
                            Width =1305
                            Height =270
                            FontSize =10
                            Name ="Label46"
                            Caption ="Leasee"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =10035
                            LayoutCachedTop =1470
                            LayoutCachedWidth =11340
                            LayoutCachedHeight =1740
                            RowStart =2
                            RowEnd =2
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin Tab
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =90
                    Top =720
                    Width =3180
                    Height =4725
                    FontSize =10
                    TabIndex =3
                    Name ="BankClientTabs"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =90
                    LayoutCachedTop =720
                    LayoutCachedWidth =3270
                    LayoutCachedHeight =5445
                    ThemeFontIndex =1
                    UseTheme =1
                    Shape =0
                    BackColor =16777215
                    BackThemeColorIndex =-1
                    BackShade =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    HoverThemeColorIndex =1
                    HoverShade =85.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeColor =8210719
                    PressedForeColor =8210719
                    ForeColor =8355711
                    ForeTint =50.0
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =165
                            Top =1155
                            Width =3030
                            Height =4215
                            Name ="ByBank"
                            OnClick ="[Event Procedure]"
                            Caption ="By Bank"
                            LayoutCachedLeft =165
                            LayoutCachedTop =1155
                            LayoutCachedWidth =3195
                            LayoutCachedHeight =5370
                            Begin
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    OverlapFlags =215
                                    OldBorderStyle =0
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    Left =180
                                    Top =1155
                                    Width =3000
                                    Height =4185
                                    FontSize =8
                                    BorderColor =4754549
                                    Name ="BankList"
                                    RowSourceType ="Table/Query"
                                    RowSource ="DrawBanksList"
                                    ColumnWidths ="0;1728;1152"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =1155
                                    LayoutCachedWidth =3180
                                    LayoutCachedHeight =5340
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =165
                            Top =1155
                            Width =3030
                            Height =4215
                            Name ="ByClient"
                            OnClick ="[Event Procedure]"
                            Caption ="By Client"
                            LayoutCachedLeft =165
                            LayoutCachedTop =1155
                            LayoutCachedWidth =3195
                            LayoutCachedHeight =5370
                            Begin
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    Left =225
                                    Top =1155
                                    Width =2880
                                    Height =4079
                                    FontSize =8
                                    Name ="ClientsList"
                                    RowSourceType ="Table/Query"
                                    RowSource ="DrawClientsList"
                                    ColumnWidths ="0;1728;1152"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =225
                                    LayoutCachedTop =1155
                                    LayoutCachedWidth =3105
                                    LayoutCachedHeight =5234
                                    BorderThemeColorIndex =9
                                    BorderShade =50.0
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =165
                            Top =1155
                            Width =3030
                            Height =4215
                            Name ="By Invoice"
                            EventProcPrefix ="By_Invoice"
                            LayoutCachedLeft =165
                            LayoutCachedTop =1155
                            LayoutCachedWidth =3195
                            LayoutCachedHeight =5370
                            Begin
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    IMESentenceMode =3
                                    ColumnCount =6
                                    Left =260
                                    Top =1235
                                    Width =2880
                                    Height =4079
                                    FontSize =8
                                    Name ="InvoiceList"
                                    RowSourceType ="Table/Query"
                                    RowSource ="DrawnUnitInvoices"
                                    ColumnWidths ="0;1296;936;432;0;0"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =260
                                    LayoutCachedTop =1235
                                    LayoutCachedWidth =3140
                                    LayoutCachedHeight =5314
                                    BorderThemeColorIndex =9
                                    BorderShade =50.0
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =105
                    Top =345
                    Width =2775
                    Height =270
                    FontSize =9
                    BackColor =13434879
                    Name ="SearchClient"
                    OnChange ="[Event Procedure]"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =105
                    LayoutCachedTop =345
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =615
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =9
                    BorderShade =75.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11340
                    Top =4350
                    Width =3180
                    Height =270
                    FontSize =9
                    TabIndex =17
                    Name ="SaveUpdate"
                    Caption ="Update | Save"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save - Update the Draw displayed"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11340
                    LayoutCachedTop =4350
                    LayoutCachedWidth =14520
                    LayoutCachedHeight =4620
                    RowStart =12
                    RowEnd =12
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =12
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverThemeColorIndex =0
                    HoverTint =80.0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-5
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15900
                    Top =5685
                    Width =750
                    Height =225
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =24
                    Name ="DrawDownId"
                    ControlSource ="DrawDownId"

                    LayoutCachedLeft =15900
                    LayoutCachedTop =5685
                    LayoutCachedWidth =16650
                    LayoutCachedHeight =5910
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =15525
                            Top =5685
                            Width =315
                            Height =225
                            FontSize =9
                            Name ="Label9"
                            Caption ="ID:"
                            LayoutCachedLeft =15525
                            LayoutCachedTop =5685
                            LayoutCachedWidth =15840
                            LayoutCachedHeight =5910
                            BorderThemeColorIndex =1
                            BorderShade =65.0
                            ForeThemeColorIndex =1
                            ForeShade =65.0
                        End
                    End
                End
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =9810
                    Top =5100
                    Width =6675
                    Height =285
                    Name ="CoverBox"
                    LayoutCachedLeft =9810
                    LayoutCachedTop =5100
                    LayoutCachedWidth =16485
                    LayoutCachedHeight =5385
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1260
                    Top =5520
                    Width =1980
                    Height =270
                    FontSize =9
                    TabIndex =21
                    Name ="OpenDrawDownManager"
                    Caption ="Draw Down Manager"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add a New Draw"
                    LeftPadding =90
                    RightPadding =105
                    BottomPadding =165
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1260
                    LayoutCachedTop =5520
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =5790
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =2880
                    Top =345
                    Width =345
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="Clear Search"
                    Caption ="Clear"
                    OnClick ="[Event Procedure]"
                    EventProcPrefix ="Clear_Search"
                    ControlTipText ="Click to Search Group"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
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

                    LayoutCachedLeft =2880
                    LayoutCachedTop =345
                    LayoutCachedWidth =3225
                    LayoutCachedHeight =615
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =4
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =4
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =3
                    Left =105
                    Top =105
                    Width =3120
                    Height =240
                    FontSize =8
                    FontWeight =700
                    Name ="Label668"
                    Caption ="Type Search, Here ▼                                           "
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =0
                    LayoutCachedLeft =105
                    LayoutCachedTop =105
                    LayoutCachedWidth =3225
                    LayoutCachedHeight =345
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =9
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15225
                    Top =2820
                    Width =1185
                    Height =270
                    TabIndex =13
                    Name ="DrawPaidOff"
                    ControlSource ="DrawPaidOff"
                    Tag ="EditMe"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15225
                    LayoutCachedTop =2820
                    LayoutCachedWidth =16410
                    LayoutCachedHeight =3090
                    RowStart =7
                    RowEnd =7
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =6
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =13920
                            Top =2820
                            Width =1305
                            Height =270
                            FontSize =10
                            Name ="Label116"
                            Caption ="Paid Off Date"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =13920
                            LayoutCachedTop =2820
                            LayoutCachedWidth =15225
                            LayoutCachedHeight =3090
                            RowStart =7
                            RowEnd =7
                            ColumnStart =2
                            ColumnEnd =2
                            BackThemeColorIndex =6
                            BackTint =20.0
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =225
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12600
                    Width =1875
                    Height =195
                    FontSize =8
                    Name ="ModifiedBy"
                    ControlSource ="ModifiedBy"

                    LayoutCachedLeft =12600
                    LayoutCachedWidth =14475
                    LayoutCachedHeight =195
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14520
                    Width =2085
                    Height =195
                    FontSize =8
                    TabIndex =1
                    Name ="ModifiedOn"
                    ControlSource ="ModifiedOn"

                    LayoutCachedLeft =14520
                    LayoutCachedWidth =16605
                    LayoutCachedHeight =195
                End
            End
        End
    End
End
CodeBehindForm
' See "DrawDownMaker.cls"
