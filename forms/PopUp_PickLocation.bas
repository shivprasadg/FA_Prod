Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7245
    DatasheetFontHeight =11
    ItemSuffix =13
    Left =5880
    Top =5220
    Right =14430
    Bottom =8175
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xbf2692420e41e640
    End
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
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
        Begin PageHeader
            DisplayWhen =1
            Height =720
            Name ="PageHeaderSection"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Left =45
                    Top =540
                    Width =7200
                    Height =180
                    Name ="Image142"
                    Picture ="3_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedLeft =45
                    LayoutCachedTop =540
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =720
                    TabIndex =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6420
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    ForeColor =0
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

                    LayoutCachedLeft =6420
                    LayoutCachedTop =120
                    LayoutCachedWidth =6975
                    LayoutCachedHeight =348
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Width =2505
                    Height =525
                    FontSize =18
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="Pick A Location"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedWidth =2505
                    LayoutCachedHeight =525
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =690
                    Top =375
                    Width =1815
                    Height =240
                    FontSize =8
                    Name ="Label227"
                    Caption ="Any location"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =690
                    LayoutCachedTop =375
                    LayoutCachedWidth =2505
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
            Height =600
            Name ="Detail"
            Begin
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =6
                    ListWidth =7200
                    Left =1260
                    Top =120
                    Width =4020
                    Height =300
                    FontWeight =700
                    Name ="XLocationID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LAddrLine1, Locat"
                        "ions.LCity, Locations.LState, Locations.LocObsolete, Locations.LocObsolete, Loca"
                        "tions.LocObsdate, Locations.LocObsCmt FROM Locations WHERE (((Locations.LocObsol"
                        "ete)=No) AND ((Locations.ClientGroupID)=771)) ORDER BY Locations.LocationName, L"
                        "ocations.LocationType;"
                    ColumnWidths ="0;2160;2160;1440;720;720"
                    StatusBarText ="Link to Contacts table"
                    BottomPadding =150

                    LayoutCachedLeft =1260
                    LayoutCachedTop =120
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =420
                    ColumnStart =3
                    ColumnEnd =3
                    ThemeFontIndex =1
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =4
                    BorderTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =120
                            Width =1020
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label5165"
                            Caption ="Location"
                            BottomPadding =150
                            LayoutCachedLeft =180
                            LayoutCachedTop =120
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =420
                            ColumnStart =2
                            ColumnEnd =2
                            ThemeFontIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =119
                    PictureType =1
                    Left =5325
                    Top =105
                    Width =1155
                    Height =300
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="SubmitMe"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =5325
                    LayoutCachedTop =105
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =405
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =9
                    HoverTint =60.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =0
            Name ="PageFooterSection"
        End
    End
End
CodeBehindForm
' See "PopUp_PickLocation.cls"
