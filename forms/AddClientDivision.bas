Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5640
    DatasheetFontHeight =11
    ItemSuffix =70
    Left =10560
    Top =6345
    Right =16200
    Bottom =8595
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x99d01076e91ee440
    End
    Caption ="Add a Division"
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            Height =405
            Name ="FormHeader"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =120
                    Width =5520
                    Height =405
                    FontSize =16
                    FontWeight =700
                    Name ="Label4"
                    Caption ="Add Client Division"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =405
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =4200
                    Top =60
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

                    LayoutCachedLeft =4200
                    LayoutCachedTop =60
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =288
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
                    OverlapFlags =215
                    PictureType =1
                    Left =4767
                    Top =60
                    Width =750
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

                    LayoutCachedLeft =4767
                    LayoutCachedTop =60
                    LayoutCachedWidth =5517
                    LayoutCachedHeight =288
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
            Height =1620
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =540
                    Top =855
                    Width =5025
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    ForeColor =6108695
                    Name ="AddClientDvision"
                    Caption ="Press Here,  to Add Division"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Find Next"
                    LeftPadding =165
                    TopPadding =105
                    RightPadding =180
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xb17d4a45b17d4af9b17d4affb17d4afcb17d4a5a000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000b17d4a4eb17d4af9b17d4affb17d4af9b17d4a4b0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a57b17d4afcb17d4affb17d4af6b17d4a3f00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a60b17d4affb17d4affb17d4af0b17d4a33 ,
                        0x000000000000000000000000b17d4a03b17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aea ,
                        0xb17d4a27000000000000000000000000b17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4ae7 ,
                        0xb17d4a2400000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a5ab17d4affb17d4affb17d4aeab17d4a2d ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a45b17d4af9b17d4affb17d4af0b17d4a3300000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000b17d4a36b17d4af6b17d4affb17d4af6b17d4a390000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xb17d4a24b17d4aeab17d4affb17d4af6b17d4a42000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =540
                    LayoutCachedTop =855
                    LayoutCachedWidth =5565
                    LayoutCachedHeight =1170
                    PictureCaptionArrangement =5
                    Alignment =1
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =255
                    BackColor =14270637
                    BorderColor =16777215
                    HoverColor =11525325
                    PressedColor =14211288
                    HoverForeColor =6108695
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2205
                    Top =300
                    Width =3405
                    Height =315
                    BorderColor =10040879
                    Name ="NewDivisionName"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2205
                    LayoutCachedTop =300
                    LayoutCachedWidth =5610
                    LayoutCachedHeight =615
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =480
                    Top =300
                    Width =1665
                    Height =315
                    FontSize =12
                    Name ="Label8"
                    Caption =" New Divison Name►"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =480
                    LayoutCachedTop =300
                    LayoutCachedWidth =2145
                    LayoutCachedHeight =615
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    Left =210
                    Top =270
                    Width =420
                    Height =315
                    FontSize =18
                    Name ="Label2467"
                    Caption =""
                    FontName ="Wingdings"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =210
                    LayoutCachedTop =270
                    LayoutCachedWidth =630
                    LayoutCachedHeight =585
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    Left =210
                    Top =825
                    Width =420
                    Height =315
                    FontSize =18
                    Name ="Label2468"
                    Caption =""
                    FontName ="Wingdings"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =210
                    LayoutCachedTop =825
                    LayoutCachedWidth =630
                    LayoutCachedHeight =1140
                    RowStart =4
                    RowEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin FormFooter
            Height =240
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Width =720
                    FontSize =8
                    Name ="ClientGroupID"
                    FontName ="Tahoma"

                    LayoutCachedWidth =720
                    LayoutCachedHeight =240
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
                Begin TextBox
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Width =720
                    FontSize =8
                    TabIndex =1
                    Name ="ClientID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =720
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =240
                    ForeThemeColorIndex =1
                    ForeShade =75.0
                End
            End
        End
    End
End
CodeBehindForm
' See "AddClientDivision.cls"
