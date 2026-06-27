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
    Width =5700
    DatasheetFontHeight =11
    ItemSuffix =70
    Left =7230
    Top =11280
    Right =16335
    Bottom =17595
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x99d01076e91ee440
    End
    Caption ="A Company Must Be Part Of  A Client Group"
    DatasheetFontName ="Calibri"
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
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Width =5700
                    Height =375
                    BackColor =8210719
                    Name ="Box6"
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =1
                    Left =120
                    Width =5520
                    Height =405
                    FontSize =16
                    FontWeight =700
                    Name ="Label4"
                    Caption ="Add or Select a Client Group"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =405
                    ForeThemeColorIndex =1
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
            Height =2340
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3450
                    Top =1185
                    Width =2175
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    BackColor =13434879
                    BorderColor =62207
                    ForeColor =-2147483617
                    Name ="Password"
                    InputMask ="Password"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3450
                    LayoutCachedTop =1185
                    LayoutCachedWidth =5625
                    LayoutCachedHeight =1500
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =163
                    Left =585
                    Top =1815
                    Width =5025
                    Height =315
                    FontWeight =700
                    TabIndex =3
                    Name ="AddClientButton"
                    Caption ="Press Here,  to Input the Company's info."
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find Next"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =585
                    LayoutCachedTop =1815
                    LayoutCachedWidth =5610
                    LayoutCachedHeight =2130
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    HoverThemeColorIndex =2
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3450
                    Top =60
                    Width =2175
                    Height =315
                    BorderColor =10040879
                    Name ="NewClientGroup"
                    OnChange ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3450
                    LayoutCachedTop =60
                    LayoutCachedWidth =5625
                    LayoutCachedHeight =375
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =495
                    Top =60
                    Width =2940
                    Height =315
                    FontSize =12
                    Name ="Label8"
                    Caption ="Create a new Client Group ►"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =495
                    LayoutCachedTop =60
                    LayoutCachedWidth =3435
                    LayoutCachedHeight =375
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =165
                    Top =615
                    Width =3270
                    Height =300
                    FontSize =12
                    Name ="Label9"
                    Caption ="Use an existing Client Group ►"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =165
                    LayoutCachedTop =615
                    LayoutCachedWidth =3435
                    LayoutCachedHeight =915
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =585
                    Top =345
                    Width =2850
                    Height =240
                    FontWeight =700
                    Name ="Label10"
                    Caption ="- or-"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =585
                    LayoutCachedTop =345
                    LayoutCachedWidth =3435
                    LayoutCachedHeight =585
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =3450
                    Top =615
                    Width =2175
                    Height =300
                    TabIndex =1
                    BorderColor =10040879
                    Name ="ExistingClientGroup"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ClientGroups.ClientGroupId, ClientGroups.ClientGroupName, ClientGroups.Is"
                        "Deleted FROM ClientGroups WHERE (((ClientGroups.IsDeleted)=False)) ORDER BY Clie"
                        "ntGroups.ClientGroupName;"
                    ColumnWidths ="0;2160"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Client Group must be entered - add or select short name if not part of a group"
                    OnClick ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =3450
                    LayoutCachedTop =615
                    LayoutCachedWidth =5625
                    LayoutCachedHeight =915
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =585
                    Top =1185
                    Width =2850
                    Height =315
                    FontSize =12
                    Name ="Label6"
                    Caption ="Enter Authorization Code ►"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =585
                    LayoutCachedTop =1185
                    LayoutCachedWidth =3435
                    LayoutCachedHeight =1500
                    RowStart =4
                    RowEnd =4
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
                    Left =225
                    Top =30
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
                    LayoutCachedLeft =225
                    LayoutCachedTop =30
                    LayoutCachedWidth =645
                    LayoutCachedHeight =345
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    Left =225
                    Top =1140
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
                    LayoutCachedLeft =225
                    LayoutCachedTop =1140
                    LayoutCachedWidth =645
                    LayoutCachedHeight =1455
                    RowStart =4
                    RowEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    Left =225
                    Top =1815
                    Width =420
                    Height =315
                    FontSize =18
                    Name ="Label2469"
                    Caption =""
                    FontName ="Wingdings"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =225
                    LayoutCachedTop =1815
                    LayoutCachedWidth =645
                    LayoutCachedHeight =2130
                    RowStart =6
                    RowEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
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
' See "AddClientGroup.cls"
