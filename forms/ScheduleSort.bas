Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7568
    DatasheetFontHeight =11
    ItemSuffix =17
    Left =9000
    Top =4005
    Right =20145
    Bottom =10350
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x2b96e6e7b68ae440
    End
    RecordSource ="ScheduleSortOrder"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =870
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =75
                    Top =435
                    Width =6720
                    Height =270
                    FontWeight =700
                    Name ="Text5"
                    ControlSource ="=[CLIENTSHNM] & \" \" & [MLNO]"

                    LayoutCachedLeft =75
                    LayoutCachedTop =435
                    LayoutCachedWidth =6795
                    LayoutCachedHeight =705
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =75
                    Width =5055
                    Height =345
                    FontSize =16
                    Name ="Label16"
                    Caption ="CHANGE SCHEDULE SORT ORDER"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =60
                    LayoutCachedTop =75
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =95
                    PictureType =1
                    Left =6705
                    Top =105
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
                    HorizontalAnchor =1
                    BackStyle =0

                    LayoutCachedLeft =6705
                    LayoutCachedTop =105
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =333
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    HoverForeThemeColorIndex =4
                    HoverForeShade =75.0
                    PressedForeThemeColorIndex =4
                    PressedForeShade =75.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6120
                    Top =105
                    Width =555
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1
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

                    LayoutCachedLeft =6120
                    LayoutCachedTop =105
                    LayoutCachedWidth =6675
                    LayoutCachedHeight =333
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    HoverForeThemeColorIndex =4
                    HoverForeShade =75.0
                    PressedForeThemeColorIndex =4
                    PressedForeShade =75.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =390
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1860
                    Top =48
                    Width =636
                    Height =312
                    Name ="SCHNo1"
                    ControlSource ="SCHNo1"
                    Format ="Fixed"
                    StatusBarText ="added 8/15/2013 schedule number for sequencing"
                    GroupTable =1

                    LayoutCachedLeft =1860
                    LayoutCachedTop =48
                    LayoutCachedWidth =2496
                    LayoutCachedHeight =360
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =348
                            Top =48
                            Width =1440
                            Height =312
                            Name ="Label2"
                            Caption ="ORDER NO"
                            GroupTable =1
                            LayoutCachedLeft =348
                            LayoutCachedTop =48
                            LayoutCachedWidth =1788
                            LayoutCachedHeight =360
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3468
                    Top =48
                    Width =3300
                    Height =312
                    TabIndex =1
                    Name ="SchNo"
                    ControlSource ="SchNo"
                    StatusBarText ="schedule number starts with 1"
                    GroupTable =1

                    LayoutCachedLeft =3468
                    LayoutCachedTop =48
                    LayoutCachedWidth =6768
                    LayoutCachedHeight =360
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2556
                            Top =48
                            Width =840
                            Height =312
                            Name ="Label3"
                            Caption ="Sch No:"
                            GroupTable =1
                            LayoutCachedLeft =2556
                            LayoutCachedTop =48
                            LayoutCachedWidth =3396
                            LayoutCachedHeight =360
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
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
' See "ScheduleSort.cls"
