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
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13701
    DatasheetFontHeight =11
    ItemSuffix =85
    Left =6255
    Top =2775
    Right =21720
    Bottom =15615
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x39278026d95ee440
    End
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    OnCloseEmMacro = Begin
        Version =196611
        ColumnsShown =8
        Begin
            Action ="Close"
            Argument ="2"
            Argument ="bosubform"
            Argument ="0"
        End
    End
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =540
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =360
                    Width =13680
                    Height =180
                    Name ="Image142"
                    Picture ="3_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =360
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =540
                    TabIndex =2
                End
                Begin Label
                    OverlapFlags =247
                    Left =120
                    Width =2535
                    Height =495
                    FontSize =20
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label50"
                    Caption ="Move Group"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =2655
                    LayoutCachedHeight =495
                End
                Begin CommandButton
                    OverlapFlags =127
                    Left =12240
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

                    LayoutCachedLeft =12240
                    LayoutCachedTop =60
                    LayoutCachedWidth =12795
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
                    OverlapFlags =247
                    PictureType =1
                    Left =12807
                    Top =60
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
                    HorizontalAnchor =1

                    LayoutCachedLeft =12807
                    LayoutCachedTop =60
                    LayoutCachedWidth =13362
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
            CanGrow = NotDefault
            Height =12165
            Name ="Detail"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    Left =9435
                    Top =7455
                    Width =4176
                    Height =480
                    FontSize =8
                    FontWeight =700
                    BackColor =15527148
                    Name ="Label42"
                    Caption =" Asset               | VIN Search                                            |  "
                        " UnitNo "
                    VerticalAnchor =2
                    LayoutCachedLeft =9435
                    LayoutCachedTop =7455
                    LayoutCachedWidth =13611
                    LayoutCachedHeight =7935
                    ForeThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Top =7455
                    Width =6441
                    Height =480
                    FontSize =8
                    FontWeight =700
                    BackColor =15527148
                    Name ="Label58"
                    Caption =" Asset         | VIN Search                                           |   UnitNo"
                        "            | Year      | Make                 | Type"
                    LayoutCachedTop =7455
                    LayoutCachedWidth =6441
                    LayoutCachedHeight =7935
                    ForeThemeColorIndex =2
                End
                Begin ComboBox
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2160
                    Left =1380
                    Top =135
                    Width =4350
                    Height =270
                    Name ="SelectCompany"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Clients.clientID, Clients.ClientShNm FROM Clients INNER JOIN Mst"
                        "rLease ON Clients.clientID = MstrLease.Client WHERE (((Clients.ClientShNm) Not L"
                        "ike \"demo*\") AND ((Clients.clientType)=\"client\" Or (Clients.clientType)=\"gu"
                        "arantor\" Or (Clients.clientType) Like \"*target\")) ORDER BY Clients.ClientShNm"
                        ";"
                    ColumnWidths ="0;2160"
                    ValidationRule ="Is Not Null"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1380
                    LayoutCachedTop =135
                    LayoutCachedWidth =5730
                    LayoutCachedHeight =405
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =1
                            Left =255
                            Top =60
                            Width =5505
                            Height =375
                            FontSize =10
                            FontWeight =700
                            LeftMargin =45
                            TopMargin =45
                            RightMargin =45
                            BottomMargin =45
                            Name ="Label1"
                            Caption ="Company:"
                            LayoutCachedLeft =255
                            LayoutCachedTop =60
                            LayoutCachedWidth =5760
                            LayoutCachedHeight =435
                            BackThemeColorIndex =2
                            BackTint =60.0
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1380
                    Top =480
                    Width =4350
                    Height =270
                    TabIndex =3
                    Name ="LeaseList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MLID, MLNo, MLOrig FROM MstrLease WHERE Client=4971 ORDER By MLNo;"
                    ColumnWidths ="0;2160;720"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1380
                    LayoutCachedTop =480
                    LayoutCachedWidth =5730
                    LayoutCachedHeight =750
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextFontCharSet =238
                            TextAlign =1
                            Left =255
                            Top =405
                            Width =5505
                            Height =390
                            FontSize =10
                            FontWeight =700
                            LeftMargin =45
                            TopMargin =45
                            RightMargin =45
                            BottomMargin =45
                            Name ="Label3"
                            Caption ="Lease:"
                            LayoutCachedLeft =255
                            LayoutCachedTop =405
                            LayoutCachedWidth =5760
                            LayoutCachedHeight =795
                            BackThemeColorIndex =4
                            BackShade =75.0
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =7
                    Top =7965
                    Width =7506
                    Height =2550
                    FontSize =10
                    TabIndex =18
                    Name ="UnitsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitID AS AssetId, unitvin AS VIN, UnitUnitNum AS UnitNo, UnitYr As Year,"
                        " UnitMake As Make, UnitType As Type FROM Units WHERE GroupREF=1971 ORDER BY Unit"
                        "ID;"
                    ColumnWidths ="720;2160;1008;576;1008;1008;1008"
                    ControlTipText ="Unit to be moved"
                    VerticalAnchor =2
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedTop =7965
                    LayoutCachedWidth =7506
                    LayoutCachedHeight =10515
                    BackThemeColorIndex =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =9435
                    Top =7965
                    Width =4176
                    Height =2550
                    FontSize =10
                    TabIndex =17
                    BackColor =16249583
                    Name ="UnitsFormsList"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="864;2160;1152"
                    OnClick ="Requery"
                    ControlTipText ="Units in Target Group"
                    VerticalAnchor =2
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =9435
                    LayoutCachedTop =7965
                    LayoutCachedWidth =13611
                    LayoutCachedHeight =10515
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =7440
                    Top =960
                    Width =1785
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label34"
                    Caption ="◄►"
                    LayoutCachedLeft =7440
                    LayoutCachedTop =960
                    LayoutCachedWidth =9225
                    LayoutCachedHeight =1200
                    ForeThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =7740
                    Top =9075
                    FontSize =10
                    FontWeight =700
                    TabIndex =20
                    Name ="MoveUnits"
                    Caption ="►Move Units►"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7740
                    LayoutCachedTop =9075
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =9435
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =26316
                    BorderThemeColorIndex =1
                    HoverColor =5676533
                    PressedColor =26316
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =7740
                    Top =8475
                    Height =510
                    FontSize =10
                    FontWeight =700
                    TabIndex =19
                    Name ="GetAll"
                    Caption ="◄Select All         Units"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7740
                    LayoutCachedTop =8475
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =8985
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =26316
                    BorderThemeColorIndex =1
                    HoverColor =5676533
                    PressedColor =26316
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Top =7695
                    Width =720
                    FontSize =9
                    TabIndex =11
                    BackColor =13434879
                    Name ="SearchUnitID"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =7695
                    LayoutCachedWidth =720
                    LayoutCachedHeight =7935
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =720
                    Top =7695
                    Width =2159
                    FontSize =9
                    TabIndex =12
                    BackColor =13434879
                    Name ="SearchVIN"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =720
                    LayoutCachedTop =7695
                    LayoutCachedWidth =2879
                    LayoutCachedHeight =7935
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2880
                    Top =7695
                    Width =1005
                    FontSize =9
                    TabIndex =13
                    BackColor =13434879
                    Name ="SearchUnitNo"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2880
                    LayoutCachedTop =7695
                    LayoutCachedWidth =3885
                    LayoutCachedHeight =7935
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9420
                    Top =7695
                    Width =864
                    FontSize =9
                    TabIndex =8
                    BackColor =13434879
                    Name ="SearchUnitID2"
                    OnChange ="[Event Procedure]"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9420
                    LayoutCachedTop =7695
                    LayoutCachedWidth =10284
                    LayoutCachedHeight =7935
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10290
                    Top =7695
                    Width =2160
                    FontSize =9
                    TabIndex =9
                    BackColor =13434879
                    Name ="SearchVIN2"
                    OnChange ="[Event Procedure]"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10290
                    LayoutCachedTop =7695
                    LayoutCachedWidth =12450
                    LayoutCachedHeight =7935
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin TextBox
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12450
                    Top =7695
                    Width =1138
                    FontSize =9
                    TabIndex =10
                    BackColor =13434879
                    Name ="SearchUnitNo2"
                    OnChange ="[Event Procedure]"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12450
                    LayoutCachedTop =7695
                    LayoutCachedWidth =13588
                    LayoutCachedHeight =7935
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =9
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =7740
                    Top =10500
                    FontSize =10
                    FontWeight =700
                    TabIndex =22
                    Name ="MoveUnitsBack"
                    Caption ="◄Move Units◄"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7740
                    LayoutCachedTop =10500
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =10860
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =7740
                    Top =9915
                    Height =510
                    FontSize =10
                    FontWeight =700
                    TabIndex =21
                    Name ="GetAllBack"
                    Caption ="        Select All► Units   "
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7740
                    LayoutCachedTop =9915
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =10425
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9375
                    Top =7080
                    Width =1365
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    BackColor =16183539
                    ForeColor =-2147483617
                    Name ="Label33"
                    ControlSource ="=\"Units: \" & [UnitsFormsList].[ListCount]-1"

                    LayoutCachedLeft =9375
                    LayoutCachedTop =7080
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =7380
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15
                    Top =7095
                    Width =1365
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BackColor =16317438
                    Name ="Label27"
                    ControlSource ="=\"Units: \" & [UnitsList].[ListCount]-1"

                    LayoutCachedLeft =15
                    LayoutCachedTop =7095
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =7395
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =7740
                    Top =10950
                    Width =1395
                    Height =315
                    FontWeight =700
                    Name ="Label51"
                    Caption ="◄- move"
                    LayoutCachedLeft =7740
                    LayoutCachedTop =10950
                    LayoutCachedWidth =9135
                    LayoutCachedHeight =11265
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =7740
                    Top =9495
                    Width =1440
                    Height =315
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label52"
                    Caption ="move -►"
                    LayoutCachedLeft =7740
                    LayoutCachedTop =9495
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =9810
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7440
                    Top =2715
                    Width =1785
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label53"
                    Caption ="◄►"
                    LayoutCachedLeft =7440
                    LayoutCachedTop =2715
                    LayoutCachedWidth =9225
                    LayoutCachedHeight =2955
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3885
                    Top =7695
                    Width =570
                    FontSize =9
                    TabIndex =14
                    BackColor =13434879
                    Name ="SearchYear"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3885
                    LayoutCachedTop =7695
                    LayoutCachedWidth =4455
                    LayoutCachedHeight =7935
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4455
                    Top =7695
                    Width =1005
                    FontSize =9
                    TabIndex =15
                    BackColor =13434879
                    Name ="SearchMake"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4455
                    LayoutCachedTop =7695
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =7935
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5460
                    Top =7695
                    Width =965
                    FontSize =9
                    TabIndex =16
                    BackColor =13434879
                    Name ="SearchType"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5460
                    LayoutCachedTop =7695
                    LayoutCachedWidth =6425
                    LayoutCachedHeight =7935
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1410
                    Top =2640
                    Width =4350
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="GroupList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SGrpID, SGrpNO FROM SchGrp WHERE SGrpSchID=835 ORDER BY SGrpNO ;"
                    ColumnWidths ="0;1440"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1410
                    LayoutCachedTop =2640
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =4080
                    BackThemeColorIndex =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =3
                            Left =240
                            Top =2640
                            Width =1110
                            Height =315
                            FontSize =10
                            FontWeight =700
                            BackColor =15858167
                            Name ="Label20"
                            Caption ="Group:"
                            LayoutCachedLeft =240
                            LayoutCachedTop =2640
                            LayoutCachedWidth =1350
                            LayoutCachedHeight =2955
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9465
                    Top =2700
                    Width =3315
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BackColor =16381933
                    Name ="NewGroupList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SGrpID, SGrpNO FROM SchGrp WHERE SGrpSchID=835 AND SGrpID <>1971 ORDER BY"
                        " SGrpNO ;"
                    ColumnWidths ="0;1440"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =9465
                    LayoutCachedTop =2700
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =4140
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =1
                            Left =12771
                            Top =2700
                            Width =900
                            Height =315
                            FontSize =10
                            FontWeight =700
                            BackColor =-2147483613
                            Name ="Label29"
                            Caption =" :Group"
                            LayoutCachedLeft =12771
                            LayoutCachedTop =2700
                            LayoutCachedWidth =13671
                            LayoutCachedHeight =3015
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9465
                    Top =960
                    Width =3315
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BackColor =16381933
                    Name ="NewSchedule"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SchID, SchNo FROM Schedule WHERE SchMLRef=156 ORDER BY SCHNo1;"
                    ColumnWidths ="0;2880"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =9465
                    LayoutCachedTop =960
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =2400
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =1
                            Left =12756
                            Top =960
                            Width =915
                            Height =315
                            FontSize =10
                            FontWeight =700
                            BackColor =-2147483613
                            Name ="Label24"
                            Caption =" :Schedule"
                            LayoutCachedLeft =12756
                            LayoutCachedTop =960
                            LayoutCachedWidth =13671
                            LayoutCachedHeight =1275
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1380
                    Top =975
                    Width =4350
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="ScheduleList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SchID, SchNo FROM Schedule WHERE SchMLRef=156 ORDER BY SCHNo1;"
                    ColumnWidths ="0;2880"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1380
                    LayoutCachedTop =975
                    LayoutCachedWidth =5730
                    LayoutCachedHeight =1290
                    BackThemeColorIndex =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =3
                            Left =210
                            Top =975
                            Width =1110
                            Height =315
                            FontSize =10
                            FontWeight =700
                            BackColor =15858167
                            Name ="Label22"
                            Caption ="Schedule:"
                            LayoutCachedLeft =210
                            LayoutCachedTop =975
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =1290
                            ForeThemeColorIndex =2
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
' See "MoveGroup.cls"
