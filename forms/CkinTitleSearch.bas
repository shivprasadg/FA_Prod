Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7020
    DatasheetFontHeight =11
    ItemSuffix =57
    Left =10635
    Top =4680
    Right =18990
    Bottom =8865
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x6bc21eca1c07e540
    End
    RecordSource ="SELECT Units.UnitID, Units.UnitVIN, Locations.LState AS TitleState, Units.UnitYr"
        ", Units.UnitTitleNumber FROM Units INNER JOIN Locations ON Units.UnitTitleLocati"
        "onRef = Locations.LocationID WHERE (((Units.UnitVIN)=[forms]![titlecheckin]![see"
        "kvin]));"
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
            Height =720
            BackColor =15527148
            Name ="FormHeader"
            AutoHeight =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =420
                    Width =7020
                    Height =240
                    BackColor =8210719
                    BorderColor =-2147483617
                    Name ="Box7"
                    GridlineColor =0
                    LayoutCachedTop =420
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =660
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =223
                    Left =90
                    Top =15
                    Width =3915
                    Height =585
                    FontSize =24
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label0"
                    Caption ="Title Log Entry"
                    LayoutCachedLeft =90
                    LayoutCachedTop =15
                    LayoutCachedWidth =4005
                    LayoutCachedHeight =600
                End
                Begin Label
                    OverlapFlags =215
                    Left =75
                    Width =3375
                    Height =585
                    FontSize =24
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label8"
                    Caption ="Title Log Entry"
                    LayoutCachedLeft =75
                    LayoutCachedWidth =3450
                    LayoutCachedHeight =585
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =5520
                    Top =60
                    Width =555
                    Height =228
                    FontSize =8
                    FontWeight =700
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
                    BackStyle =0

                    LayoutCachedLeft =5520
                    LayoutCachedTop =60
                    LayoutCachedWidth =6075
                    LayoutCachedHeight =288
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =0
                    BorderColor =8210719
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    PictureType =1
                    Left =6120
                    Top =60
                    Width =750
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    BackStyle =0

                    LayoutCachedLeft =6120
                    LayoutCachedTop =60
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =288
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =0
                    BorderColor =8210719
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3960
                    Top =60
                    Height =315
                    TabIndex =2
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    BottomPadding =150

                    LayoutCachedLeft =3960
                    LayoutCachedTop =60
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =375
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =1680
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1410
                    Width =2625
                    Height =315
                    FontSize =12
                    FontWeight =700
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    BottomPadding =150

                    LayoutCachedLeft =1410
                    LayoutCachedWidth =4035
                    LayoutCachedHeight =315
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =660
                            Width =690
                            Height =315
                            Name ="Label1"
                            Caption ="VIN"
                            BottomPadding =150
                            LayoutCachedLeft =660
                            LayoutCachedWidth =1350
                            LayoutCachedHeight =315
                            ColumnStart =1
                            ColumnEnd =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5595
                    Top =495
                    Width =975
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    Name ="TitleState"
                    ControlSource ="TitleState"
                    BottomPadding =150

                    LayoutCachedLeft =5595
                    LayoutCachedTop =495
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4095
                            Top =495
                            Width =1440
                            Height =315
                            Name ="Label2"
                            Caption ="Title State"
                            BottomPadding =150
                            LayoutCachedLeft =4095
                            LayoutCachedTop =495
                            LayoutCachedWidth =5535
                            LayoutCachedHeight =810
                            RowStart =1
                            RowEnd =1
                            ColumnStart =3
                            ColumnEnd =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =420
                    Top =1560
                    Width =6600
                    Height =0
                    FontWeight =700
                    TabIndex =6
                    ForeColor =255
                    Name ="sterr"
                    ControlSource ="=IIf([Forms]![titlecheckin]![seekst]<>[Forms]![CkinTitleSearch]![Title ST],\"Tit"
                        "le  State entered does not match Database, correct entry or contact Operations\""
                        ",Null)"

                    LayoutCachedLeft =420
                    LayoutCachedTop =1560
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =1560
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5595
                    Width =975
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="UnitYr"
                    ControlSource ="UnitYr"
                    BottomPadding =150

                    LayoutCachedLeft =5595
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =315
                    ColumnStart =4
                    ColumnEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4095
                            Width =1440
                            Height =315
                            Name ="Label5"
                            Caption ="Model Year"
                            BottomPadding =150
                            LayoutCachedLeft =4095
                            LayoutCachedWidth =5535
                            LayoutCachedHeight =315
                            ColumnStart =3
                            ColumnEnd =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =420
                    Top =1680
                    Width =6600
                    Height =0
                    FontWeight =700
                    TabIndex =7
                    ForeColor =255
                    Name ="MYErr"
                    ControlSource ="=IIf([Forms]![titlecheckin]![Seekmy]<>[Forms]![CkinTitleSearch]![unityr],\"Model"
                        " Yr does not match Database, correct entry or contact Operations\",Null)"

                    LayoutCachedLeft =420
                    LayoutCachedTop =1680
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =1680
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4290
                    Top =990
                    Width =2205
                    Height =390
                    TabIndex =5
                    Name ="LogTitleIn"
                    Caption ="Log Title In or Out"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"
                    LeftPadding =120
                    TopPadding =105
                    RightPadding =135
                    BottomPadding =120

                    LayoutCachedLeft =4290
                    LayoutCachedTop =990
                    LayoutCachedWidth =6495
                    LayoutCachedHeight =1380
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =1
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    Shadow =5
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1410
                    Top =495
                    Width =2625
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="UnitTitleNumber"
                    ControlSource ="UnitTitleNumber"
                    BottomPadding =150

                    LayoutCachedLeft =1410
                    LayoutCachedTop =495
                    LayoutCachedWidth =4035
                    LayoutCachedHeight =810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =495
                            Width =1350
                            Height =315
                            Name ="Label12"
                            Caption ="TitleNumber "
                            BottomPadding =150
                            LayoutCachedTop =495
                            LayoutCachedWidth =1350
                            LayoutCachedHeight =810
                            RowStart =1
                            RowEnd =1
                            ColumnEnd =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1410
                    Top =990
                    Width =2625
                    Height =330
                    TabIndex =3
                    BackColor =16515069
                    BorderColor =12632256
                    ForeColor =0
                    Name ="EntryReasonId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT reasons.ReasonDescription, reasons.ReasonType FROM reasons WHERE (((reaso"
                        "ns.ReasonType)=1)) ORDER BY reasons.ReasonDescription;"
                    DefaultValue ="\"Received from State\""
                    BottomPadding =150

                    LayoutCachedLeft =1410
                    LayoutCachedTop =990
                    LayoutCachedWidth =4035
                    LayoutCachedHeight =1320
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =990
                            Width =1350
                            Height =330
                            Name ="Label4"
                            Caption ="Reason"
                            BottomPadding =150
                            LayoutCachedTop =990
                            LayoutCachedWidth =1350
                            LayoutCachedHeight =1320
                            RowStart =2
                            RowEnd =2
                            ColumnEnd =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
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
' See "CkinTitleSearch.cls"
