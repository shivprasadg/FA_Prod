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
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =27360
    DatasheetFontHeight =11
    ItemSuffix =458
    Left =11850
    Top =4020
    Right =-26071
    Bottom =18105
    RecSrcDt = Begin
        0x12da72863105e640
    End
    RecordSource ="Select * From Locations Where LocationId=0"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
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
        Begin CheckBox
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
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin ToggleButton
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
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
            PressedForeThemeColorIndex =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =720
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =5
            BackShade =75.0
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    BackStyle =1
                    PictureType =2
                    Top =540
                    Width =27360
                    Height =180
                    BorderColor =6108695
                    Name ="Image142"
                    Picture ="DrkRedOnLtGrey_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =540
                    LayoutCachedWidth =27360
                    LayoutCachedHeight =720
                    TabIndex =5
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =960
                    Width =3060
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label8"
                    Caption ="Location Tax Manager"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =960
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =420
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Image
                    SizeMode =0
                    PictureType =1
                    Left =120
                    Width =780
                    Height =540
                    Name ="EventImage"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ControlSource ="ImageType"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =900
                    LayoutCachedHeight =540
                    TabIndex =6
                    RowEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9985
                    Top =210
                    Width =899
                    Height =300
                    ColumnOrder =3
                    FontSize =12
                    FontWeight =700
                    BorderColor =14277081
                    ForeColor =0
                    Name ="UnitsShown"
                    ControlSource ="=IIf([LocationsList].[ListCount]<=1,0,Format([LocationsList].[ListCount]-1,\"#,#"
                        "##\"))"
                    GridlineColor =0

                    LayoutCachedLeft =9985
                    LayoutCachedTop =210
                    LayoutCachedWidth =10884
                    LayoutCachedHeight =510
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =255
                            TextAlign =2
                            Left =9960
                            Top =30
                            Width =960
                            Height =518
                            FontSize =7
                            BackColor =14211288
                            BorderColor =14211288
                            ForeColor =0
                            Name ="Label63"
                            Caption ="Locations"
                            FontName ="Segoe UI"
                            GridlineColor =0
                            LayoutCachedLeft =9960
                            LayoutCachedTop =30
                            LayoutCachedWidth =10920
                            LayoutCachedHeight =548
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin OptionGroup
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =5160
                    Top =60
                    Width =2645
                    Height =360
                    ColumnOrder =0
                    TabIndex =1
                    BackColor =15527148
                    BorderColor =0
                    Name ="UnitSelection"
                    DefaultValue ="1"
                    ControlTipText ="Select the type of fleet lists to show."

                    LayoutCachedLeft =5160
                    LayoutCachedTop =60
                    LayoutCachedWidth =7805
                    LayoutCachedHeight =420
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =21345
                    Top =135
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =2
                    ForeColor =0
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =21345
                    LayoutCachedTop =135
                    LayoutCachedWidth =22095
                    LayoutCachedHeight =363
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
                Begin CommandButton
                    OverlapFlags =215
                    Left =20790
                    Top =135
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =3
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

                    LayoutCachedLeft =20790
                    LayoutCachedTop =135
                    LayoutCachedWidth =21345
                    LayoutCachedHeight =363
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
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =14700
                    Top =120
                    Width =1680
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="ExcelReport"
                    Caption ="Excel Report"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddd87262ddddddd876262626dddddd6262626262262628 ,
                        0x2626262626fffff662ff6ff2626f62f2268f2f8626fffff6626fff62626f62f2 ,
                        0x2628f82626fffff66268f862626f62f2262fff2626fffff6628f6f82626f62f2 ,
                        0x26ff2ff626fffff662626262626f62f22626262626fffff6d872626262626268 ,
                        0xddddd87626dddddd
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14700
                    LayoutCachedTop =120
                    LayoutCachedWidth =16380
                    LayoutCachedHeight =390
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedColor =16777215
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =13380
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =23
                    Left =300
                    Top =1905
                    Width =13440
                    Height =11310
                    FontSize =9
                    TabIndex =44
                    BoundColumn =2
                    Name ="LocationsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Client, LocationName, LocationSN,LCity, LState, LCounty, LZip, LUnits, FA"
                        "Units, Distro, LTaxRate, LTaxOption, LocTaxes, StateSales, CountySales, CitySale"
                        "s, StateRReceipt, LStateID, LCountyID, LCityID, LocationID, ClientGroupId, clien"
                        "tID FROM LocationDistroTaxes  Order By Client, LUnits Desc"
                    ColumnWidths ="1440;2088;1800;1440;662;1440;1152;720;720;720;720;0;0;0;0;0;0;0;0;0;0;0;0;0"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    VerticalAnchor =2

                    LayoutCachedLeft =300
                    LayoutCachedTop =1905
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =13215
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    GridlineThemeColorIndex =-1
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9150
                    Top =1410
                    Width =930
                    FontSize =9
                    TabIndex =22
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchLZip"
                    Tag ="SearchX"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =9150
                    LayoutCachedTop =1410
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =1650
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =270
                    Top =1410
                    Width =1195
                    FontSize =9
                    TabIndex =10
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchClient"
                    Tag ="SearchX"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =270
                    LayoutCachedTop =1410
                    LayoutCachedWidth =1465
                    LayoutCachedHeight =1650
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1725
                    Top =1410
                    Width =1821
                    FontSize =9
                    TabIndex =12
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchLocation"
                    Tag ="SearchX"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =1725
                    LayoutCachedTop =1410
                    LayoutCachedWidth =3546
                    LayoutCachedHeight =1650
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5610
                    Top =1410
                    Width =1185
                    FontSize =9
                    TabIndex =16
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchLCity"
                    Tag ="SearchX"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =5610
                    LayoutCachedTop =1410
                    LayoutCachedWidth =6795
                    LayoutCachedHeight =1650
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =600
                    Top =1185
                    Width =9540
                    Height =225
                    FontSize =9
                    BorderColor =0
                    Name ="Label88"
                    Caption =" ┌─────────────────┬────────────────────┬────────────┬─────────────────────────┬"
                        "────────────────────────"
                    GridlineColor =0
                    LayoutCachedLeft =600
                    LayoutCachedTop =1185
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =1410
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =270
                    Top =420
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    ForeColor =0
                    Name ="OpenClient"
                    Caption ="View Client"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =270
                    LayoutCachedTop =420
                    LayoutCachedWidth =1710
                    LayoutCachedHeight =690
                    LayoutGroup =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =1
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7050
                    Top =1410
                    Width =420
                    FontSize =9
                    TabIndex =18
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchLState"
                    Tag ="SearchX"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =7050
                    LayoutCachedTop =1410
                    LayoutCachedWidth =7470
                    LayoutCachedHeight =1650
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7710
                    Top =1410
                    Width =1185
                    FontSize =9
                    TabIndex =20
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchLCounty"
                    Tag ="SearchX"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =7710
                    LayoutCachedTop =1410
                    LayoutCachedWidth =8895
                    LayoutCachedHeight =1650
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7740
                    Top =900
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    ForeColor =0
                    Name ="ResetForm"
                    Caption ="Reset Form"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Clear seaches and reset the list of units."
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7740
                    LayoutCachedTop =900
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =1170
                    ColumnStart =7
                    ColumnEnd =7
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =240
                    Width =11280
                    Height =345
                    FontWeight =700
                    ForeColor =0
                    Name ="Label92"
                    Caption ="❶ Type in the Search box(es)  ❷ Run Search  ❸ Select 'Desired' Location From lis"
                        "t  ❹ Edit as you please'"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =240
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =345
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =7470
                    Top =1410
                    Width =240
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =19
                    ForeColor =0
                    Name ="SearchState_Button"
                    Caption ="Search State"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to Search Schedule"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af2b17d4aff ,
                        0xb17d4a8b00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af2 ,
                        0xb17d4affb17d4a7e000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4a5bb17d4aefb17d4a79000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4aeab17d4affb17d4affb17d4a99b17d4a0400000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af3 ,
                        0xb17d4affb17d4a7db17d4a41b17d4af1b17d4affb17d4affb17d4ab6b17d4a0c ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af3b17d4aff ,
                        0xb17d4a8a0000000000000000b17d4a2eb17d4ae6b17d4affb17d4affb17d4ace ,
                        0xb17d4a7db17d4adcb17d4afbb17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x00000000000000000000000000000000b17d4a1fb17d4ad7b17d4affb17d4aff ,
                        0xb27f4cfccdac8cf6dcc4adffcdac8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x0000000000000000000000000000000000000000b17d4a12b17d4ad6b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded2ffb17e4bfdb17d4a8100000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adbcdab8bf7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4af9dbc3acff ,
                        0xffffffffffffffffffffffffffffffffffffffffdbc3acffb17d4af900000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adccdab8af7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a83b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded1ffb17e4bfdb17d4a8000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a0ab17d4ac5 ,
                        0xb27f4cfccdab8bf6dbc3abffcdab8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a09 ,
                        0xb17d4a81b17d4adcb17d4af9b17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =7470
                    LayoutCachedTop =1410
                    LayoutCachedWidth =7710
                    LayoutCachedHeight =1650
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =3
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =6795
                    Top =1410
                    Width =240
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =17
                    ForeColor =0
                    Name ="SearchCity_Button"
                    Caption ="Search City"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to Search Master Lease"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af2b17d4aff ,
                        0xb17d4a8b00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af2 ,
                        0xb17d4affb17d4a7e000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4a5bb17d4aefb17d4a79000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4aeab17d4affb17d4affb17d4a99b17d4a0400000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af3 ,
                        0xb17d4affb17d4a7db17d4a41b17d4af1b17d4affb17d4affb17d4ab6b17d4a0c ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af3b17d4aff ,
                        0xb17d4a8a0000000000000000b17d4a2eb17d4ae6b17d4affb17d4affb17d4ace ,
                        0xb17d4a7db17d4adcb17d4afbb17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x00000000000000000000000000000000b17d4a1fb17d4ad7b17d4affb17d4aff ,
                        0xb27f4cfccdac8cf6dcc4adffcdac8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x0000000000000000000000000000000000000000b17d4a12b17d4ad6b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded2ffb17e4bfdb17d4a8100000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adbcdab8bf7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4af9dbc3acff ,
                        0xffffffffffffffffffffffffffffffffffffffffdbc3acffb17d4af900000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adccdab8af7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a83b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded1ffb17e4bfdb17d4a8000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a0ab17d4ac5 ,
                        0xb27f4cfccdab8bf6dbc3abffcdab8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a09 ,
                        0xb17d4a81b17d4adcb17d4af9b17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6795
                    LayoutCachedTop =1410
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =1650
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =3
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =1470
                    Top =1410
                    Width =240
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    ForeColor =0
                    Name ="SearchClient_Button"
                    Caption ="Search Client"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to Search Client"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af2b17d4aff ,
                        0xb17d4a8b00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af2 ,
                        0xb17d4affb17d4a7e000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4a5bb17d4aefb17d4a79000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4aeab17d4affb17d4affb17d4a99b17d4a0400000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af3 ,
                        0xb17d4affb17d4a7db17d4a41b17d4af1b17d4affb17d4affb17d4ab6b17d4a0c ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af3b17d4aff ,
                        0xb17d4a8a0000000000000000b17d4a2eb17d4ae6b17d4affb17d4affb17d4ace ,
                        0xb17d4a7db17d4adcb17d4afbb17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x00000000000000000000000000000000b17d4a1fb17d4ad7b17d4affb17d4aff ,
                        0xb27f4cfccdac8cf6dcc4adffcdac8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x0000000000000000000000000000000000000000b17d4a12b17d4ad6b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded2ffb17e4bfdb17d4a8100000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adbcdab8bf7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4af9dbc3acff ,
                        0xffffffffffffffffffffffffffffffffffffffffdbc3acffb17d4af900000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adccdab8af7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a83b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded1ffb17e4bfdb17d4a8000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a0ab17d4ac5 ,
                        0xb27f4cfccdab8bf6dbc3abffcdab8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a09 ,
                        0xb17d4a81b17d4adcb17d4af9b17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =1470
                    LayoutCachedTop =1410
                    LayoutCachedWidth =1710
                    LayoutCachedHeight =1650
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =3
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =3540
                    Top =1410
                    Width =240
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    ForeColor =0
                    Name ="SearchLocation_Button"
                    Caption ="Search Location"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to Search FAID"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af2b17d4aff ,
                        0xb17d4a8b00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af2 ,
                        0xb17d4affb17d4a7e000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4a5bb17d4aefb17d4a79000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4aeab17d4affb17d4affb17d4a99b17d4a0400000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af3 ,
                        0xb17d4affb17d4a7db17d4a41b17d4af1b17d4affb17d4affb17d4ab6b17d4a0c ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af3b17d4aff ,
                        0xb17d4a8a0000000000000000b17d4a2eb17d4ae6b17d4affb17d4affb17d4ace ,
                        0xb17d4a7db17d4adcb17d4afbb17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x00000000000000000000000000000000b17d4a1fb17d4ad7b17d4affb17d4aff ,
                        0xb27f4cfccdac8cf6dcc4adffcdac8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x0000000000000000000000000000000000000000b17d4a12b17d4ad6b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded2ffb17e4bfdb17d4a8100000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adbcdab8bf7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4af9dbc3acff ,
                        0xffffffffffffffffffffffffffffffffffffffffdbc3acffb17d4af900000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adccdab8af7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a83b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded1ffb17e4bfdb17d4a8000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a0ab17d4ac5 ,
                        0xb27f4cfccdab8bf6dbc3abffcdab8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a09 ,
                        0xb17d4a81b17d4adcb17d4af9b17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =3540
                    LayoutCachedTop =1410
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =1650
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =3
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =8895
                    Top =1410
                    Width =240
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =21
                    ForeColor =0
                    Name ="SearchCounty_Button"
                    Caption ="Search County"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to Search Group"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af2b17d4aff ,
                        0xb17d4a8b00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af2 ,
                        0xb17d4affb17d4a7e000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4a5bb17d4aefb17d4a79000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4aeab17d4affb17d4affb17d4a99b17d4a0400000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af3 ,
                        0xb17d4affb17d4a7db17d4a41b17d4af1b17d4affb17d4affb17d4ab6b17d4a0c ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af3b17d4aff ,
                        0xb17d4a8a0000000000000000b17d4a2eb17d4ae6b17d4affb17d4affb17d4ace ,
                        0xb17d4a7db17d4adcb17d4afbb17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x00000000000000000000000000000000b17d4a1fb17d4ad7b17d4affb17d4aff ,
                        0xb27f4cfccdac8cf6dcc4adffcdac8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x0000000000000000000000000000000000000000b17d4a12b17d4ad6b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded2ffb17e4bfdb17d4a8100000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adbcdab8bf7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4af9dbc3acff ,
                        0xffffffffffffffffffffffffffffffffffffffffdbc3acffb17d4af900000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adccdab8af7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a83b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded1ffb17e4bfdb17d4a8000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a0ab17d4ac5 ,
                        0xb27f4cfccdab8bf6dbc3abffcdab8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a09 ,
                        0xb17d4a81b17d4adcb17d4af9b17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =8895
                    LayoutCachedTop =1410
                    LayoutCachedWidth =9135
                    LayoutCachedHeight =1650
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =3
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =3
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    PictureType =1
                    Left =2220
                    Top =900
                    Width =5184
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    ForeColor =16777215
                    Name ="SearchALL_Button"
                    Caption ="Multi-Column Search"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Excute a Search based on ALL of the Columns that have their search box filled ou"
                        "t"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =2220
                    LayoutCachedTop =900
                    LayoutCachedWidth =7404
                    LayoutCachedHeight =1170
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =9592887
                    BackTint =100.0
                    BackShade =75.0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =2
                    QuickStyleMask =-753
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =10080
                    Top =1410
                    Width =240
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =23
                    ForeColor =0
                    Name ="SearchZip_Button"
                    Caption ="Search Zip"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to Search AssetID"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af2b17d4aff ,
                        0xb17d4a8b00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af2 ,
                        0xb17d4affb17d4a7e000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4a5bb17d4aefb17d4a79000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4aeab17d4affb17d4affb17d4a99b17d4a0400000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af3 ,
                        0xb17d4affb17d4a7db17d4a41b17d4af1b17d4affb17d4affb17d4ab6b17d4a0c ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af3b17d4aff ,
                        0xb17d4a8a0000000000000000b17d4a2eb17d4ae6b17d4affb17d4affb17d4ace ,
                        0xb17d4a7db17d4adcb17d4afbb17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x00000000000000000000000000000000b17d4a1fb17d4ad7b17d4affb17d4aff ,
                        0xb27f4cfccdac8cf6dcc4adffcdac8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x0000000000000000000000000000000000000000b17d4a12b17d4ad6b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded2ffb17e4bfdb17d4a8100000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adbcdab8bf7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4af9dbc3acff ,
                        0xffffffffffffffffffffffffffffffffffffffffdbc3acffb17d4af900000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adccdab8af7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a83b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded1ffb17e4bfdb17d4a8000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a0ab17d4ac5 ,
                        0xb27f4cfccdab8bf6dbc3abffcdab8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a09 ,
                        0xb17d4a81b17d4adcb17d4af9b17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =10080
                    LayoutCachedTop =1410
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =1650
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =3
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =3
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =3024
                    Left =270
                    Top =885
                    Width =1470
                    Height =255
                    FontSize =9
                    TabIndex =7
                    BorderColor =8210719
                    ForeColor =8210719
                    Name ="ClientList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.clientID, vw_SixKeys.ClientShortName AS Client, vw_SixKeys.Cli"
                        "entGroupId FROM vw_SixKeys WHERE (((vw_SixKeys.MLOrig)=\"FA\")) GROUP BY vw_SixK"
                        "eys.clientID, vw_SixKeys.ClientShortName, vw_SixKeys.ClientGroupId HAVING (((vw_"
                        "SixKeys.ClientGroupId)<>27)) ORDER BY vw_SixKeys.ClientShortName;"
                    ColumnWidths ="0;2880;0"
                    OnClick ="[Event Procedure]"
                    Tag ="SearchX"
                    ControlTipText ="Select a Client to Show its Units"
                    GridlineColor =0

                    LayoutCachedLeft =270
                    LayoutCachedTop =885
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =1140
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =9960
                    Top =1020
                    Width =1935
                    Height =240
                    FontSize =9
                    FontWeight =700
                    BackColor =15527148
                    BorderColor =0
                    ForeColor =12349952
                    Name ="ClearSeachBoxText"
                    Caption ="Clear All Search Boxes"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear ALL of the Search Boxes below"
                    GridlineColor =0
                    LayoutCachedLeft =9960
                    LayoutCachedTop =1020
                    LayoutCachedWidth =11895
                    LayoutCachedHeight =1260
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    Visible = NotDefault
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =10260
                    Top =600
                    Width =1530
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =15527148
                    BorderColor =0
                    ForeColor =12349952
                    Name ="ClearSeachBoxText2"
                    Caption ="Clear All Seach Boxes"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear ALL of the Search Boxes below"
                    GridlineColor =0
                    LayoutCachedLeft =10260
                    LayoutCachedTop =600
                    LayoutCachedWidth =11790
                    LayoutCachedHeight =825
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15420
                    Top =1800
                    Width =5160
                    Height =300
                    FontSize =12
                    TabIndex =32
                    ForeColor =0
                    Name ="LocationName"
                    ControlSource ="LocationName"
                    GroupTable =20

                    LayoutCachedLeft =15420
                    LayoutCachedTop =1800
                    LayoutCachedWidth =20580
                    LayoutCachedHeight =2100
                    DatasheetCaption ="Location Name"
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =20
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14100
                            Top =1800
                            Width =1260
                            Height =300
                            FontSize =10
                            BorderColor =0
                            Name ="Label39"
                            Caption ="Location Name"
                            GroupTable =20
                            LayoutCachedLeft =14100
                            LayoutCachedTop =1800
                            LayoutCachedWidth =15360
                            LayoutCachedHeight =2100
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =4
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =20
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15420
                    Top =2160
                    Width =5160
                    Height =300
                    FontSize =12
                    TabIndex =33
                    ForeColor =0
                    Name ="Locationsn"
                    ControlSource ="Locationsn"
                    StatusBarText ="short name for locaition or Lessee number"
                    GroupTable =20

                    LayoutCachedLeft =15420
                    LayoutCachedTop =2160
                    LayoutCachedWidth =20580
                    LayoutCachedHeight =2460
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =20
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14100
                            Top =2160
                            Width =1260
                            Height =300
                            FontSize =10
                            BorderColor =0
                            Name ="Label40"
                            Caption ="Short Name"
                            GroupTable =20
                            LayoutCachedLeft =14100
                            LayoutCachedTop =2160
                            LayoutCachedWidth =15360
                            LayoutCachedHeight =2460
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =4
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =20
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15420
                    Top =2520
                    Width =5160
                    Height =300
                    FontSize =12
                    TabIndex =34
                    ForeColor =0
                    Name ="LocationStreet"
                    ControlSource ="LocationStreet"
                    GroupTable =20

                    LayoutCachedLeft =15420
                    LayoutCachedTop =2520
                    LayoutCachedWidth =20580
                    LayoutCachedHeight =2820
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =20
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14100
                            Top =2520
                            Width =1260
                            Height =300
                            FontSize =10
                            BorderColor =0
                            Name ="Label41"
                            Caption ="Address 1"
                            GroupTable =20
                            LayoutCachedLeft =14100
                            LayoutCachedTop =2520
                            LayoutCachedWidth =15360
                            LayoutCachedHeight =2820
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =4
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =20
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =20
                    ListWidth =2880
                    Left =15420
                    Top =3600
                    Width =1785
                    Height =300
                    FontSize =12
                    TabIndex =37
                    ForeColor =0
                    Name ="LState"
                    ControlSource ="LState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TaxStates.PostalAbbreviation, IIf([RentReceiptAvailable]=True And [Upfron"
                        "tAvailable]=True,\"RR/Upfront\",IIf([RentReceiptAvailable]=True,\"RR\",IIf([Upfr"
                        "ontAvailable]=True,\"Upfront\",\"None\"))) AS [Option], TaxStates.ISOCountryCode"
                        "2, TaxStates.ID FROM TaxStates GROUP BY TaxStates.PostalAbbreviation, IIf([RentR"
                        "eceiptAvailable]=True And [UpfrontAvailable]=True,\"RR/Upfront\",IIf([RentReceip"
                        "tAvailable]=True,\"RR\",IIf([UpfrontAvailable]=True,\"Upfront\",\"None\"))), Tax"
                        "States.ISOCountryCode2, TaxStates.ID, TaxStates.PostalAbbreviation HAVING (((Tax"
                        "States.ISOCountryCode2)=\"USA\" Or (TaxStates.ISOCountryCode2)=\"CAN\" Or (TaxSt"
                        "ates.ISOCountryCode2)=\"MEX\")) ORDER BY TaxStates.ISOCountryCode2 DESC , TaxSta"
                        "tes.PostalAbbreviation;"
                    ColumnWidths ="432;1152;432;0"
                    StatusBarText ="lookup in lookup_state table"
                    OnClick ="[Event Procedure]"
                    GroupTable =20
                    AllowValueListEdits =0

                    LayoutCachedLeft =15420
                    LayoutCachedTop =3600
                    LayoutCachedWidth =17205
                    LayoutCachedHeight =3900
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =20
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =14100
                            Top =3600
                            Width =1260
                            Height =300
                            FontSize =10
                            FontWeight =700
                            BorderColor =0
                            Name ="Label390"
                            Caption ="STATE >>"
                            GroupTable =20
                            LayoutCachedLeft =14100
                            LayoutCachedTop =3600
                            LayoutCachedWidth =15360
                            LayoutCachedHeight =3900
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =4
                            BackThemeColorIndex =4
                            BackShade =75.0
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            GroupTable =20
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18180
                    Top =3960
                    Width =2400
                    Height =300
                    FontSize =12
                    TabIndex =39
                    ForeColor =0
                    Name ="LZip"
                    ControlSource ="LZip"
                    InputMask ="00000\\-9999;;_"
                    GroupTable =20

                    LayoutCachedLeft =18180
                    LayoutCachedTop =3960
                    LayoutCachedWidth =20580
                    LayoutCachedHeight =4260
                    RowStart =7
                    RowEnd =7
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =20
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15420
                    Top =4680
                    Width =2707
                    Height =300
                    FontSize =12
                    TabIndex =41
                    ForeColor =0
                    Name ="LPhone"
                    ControlSource ="LPhone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =20

                    LayoutCachedLeft =15420
                    LayoutCachedTop =4680
                    LayoutCachedWidth =18127
                    LayoutCachedHeight =4980
                    RowStart =9
                    RowEnd =9
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =4
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =20
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14100
                            Top =4680
                            Width =1260
                            Height =300
                            FontSize =10
                            BorderColor =0
                            Name ="Label46"
                            Caption ="Phone | FAX"
                            GroupTable =20
                            LayoutCachedLeft =14100
                            LayoutCachedTop =4680
                            LayoutCachedWidth =15360
                            LayoutCachedHeight =4980
                            RowStart =9
                            RowEnd =9
                            LayoutGroup =4
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =20
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =18180
                    Top =4680
                    Width =2400
                    Height =300
                    FontSize =12
                    TabIndex =42
                    ForeColor =0
                    Name ="LFax"
                    ControlSource ="LFax"
                    InputMask ="!\\(999\") \"000\\-0000;;_"
                    GroupTable =20

                    LayoutCachedLeft =18180
                    LayoutCachedTop =4680
                    LayoutCachedWidth =20580
                    LayoutCachedHeight =4980
                    RowStart =9
                    RowEnd =9
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =20
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15420
                    Top =2880
                    Width =5160
                    Height =300
                    FontSize =12
                    TabIndex =35
                    ForeColor =0
                    Name ="Location Add2"
                    ControlSource ="LAddrLine1"
                    EventProcPrefix ="Location_Add2"
                    GroupTable =20

                    LayoutCachedLeft =15420
                    LayoutCachedTop =2880
                    LayoutCachedWidth =20580
                    LayoutCachedHeight =3180
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =20
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14100
                            Top =2880
                            Width =1260
                            Height =300
                            FontSize =10
                            BorderColor =0
                            Name ="Label60"
                            Caption ="Address 2"
                            GroupTable =20
                            LayoutCachedLeft =14100
                            LayoutCachedTop =2880
                            LayoutCachedWidth =15360
                            LayoutCachedHeight =3180
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =4
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =20
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =19020
                    Top =5760
                    Width =2040
                    Height =270
                    FontSize =8
                    TabIndex =50
                    Name ="SaveMe"
                    Caption ="Save / Update"
                    FontName ="Tahoma"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =19020
                    LayoutCachedTop =5760
                    LayoutCachedWidth =21060
                    LayoutCachedHeight =6030
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =1
                    BackTint =100.0
                    BackShade =95.0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =1
                    HoverTint =100.0
                    HoverShade =85.0
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =2
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =2
                    PressedForeTint =100.0
                    QuickStyle =22
                    QuickStyleMask =-629
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =8
                    ListWidth =2880
                    Left =15420
                    Top =5040
                    Width =5160
                    Height =300
                    FontSize =12
                    TabIndex =43
                    ForeColor =0
                    Name ="DivisionID"
                    ControlSource ="DivisionID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DivisionID, DivisionName, ClientID From ClientDivisions WHERE Deleted=Fal"
                        "se and ClientID= Order by DivisionName"
                    ColumnWidths ="0;2160;0"
                    GroupTable =20
                    AllowValueListEdits =0

                    LayoutCachedLeft =15420
                    LayoutCachedTop =5040
                    LayoutCachedWidth =20580
                    LayoutCachedHeight =5340
                    RowStart =10
                    RowEnd =10
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =20
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14100
                            Top =5040
                            Width =1260
                            Height =300
                            FontSize =10
                            BorderColor =0
                            Name ="Label566"
                            Caption ="Division"
                            GroupTable =20
                            LayoutCachedLeft =14100
                            LayoutCachedTop =5040
                            LayoutCachedWidth =15360
                            LayoutCachedHeight =5340
                            RowStart =10
                            RowEnd =10
                            LayoutGroup =4
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =20
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15420
                    Top =3240
                    Width =5160
                    Height =300
                    FontSize =12
                    TabIndex =36
                    ForeColor =0
                    Name ="Text579"
                    ControlSource ="LAddrLine2"
                    GroupTable =20

                    LayoutCachedLeft =15420
                    LayoutCachedTop =3240
                    LayoutCachedWidth =20580
                    LayoutCachedHeight =3540
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =20
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14100
                            Top =3240
                            Width =1260
                            Height =300
                            FontSize =10
                            BorderColor =0
                            Name ="Label580"
                            Caption ="Address 3"
                            GroupTable =20
                            LayoutCachedLeft =14100
                            LayoutCachedTop =3240
                            LayoutCachedWidth =15360
                            LayoutCachedHeight =3540
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =4
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =20
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15420
                    Top =3960
                    Width =2707
                    Height =300
                    FontSize =12
                    TabIndex =38
                    ForeColor =0
                    Name ="LCity"
                    ControlSource ="LCity"
                    RowSourceType ="Table/Query"
                    RowSource ="Select City From vw_CCSZ_Lookup Where State ='GA'  AND County= 'Hancock'  Group "
                        "By City Order By City;"
                    ColumnWidths ="2160"
                    OnClick ="[Event Procedure]"
                    GroupTable =20

                    LayoutCachedLeft =15420
                    LayoutCachedTop =3960
                    LayoutCachedWidth =18127
                    LayoutCachedHeight =4260
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =2
                    LayoutGroup =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =20
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14100
                            Top =3960
                            Width =1260
                            Height =300
                            FontSize =10
                            BorderColor =0
                            Name ="Label42"
                            Caption ="City -  Zip"
                            GroupTable =20
                            LayoutCachedLeft =14100
                            LayoutCachedTop =3960
                            LayoutCachedWidth =15360
                            LayoutCachedHeight =4260
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =4
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =20
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2880
                    Left =15420
                    Top =4320
                    Width =5160
                    Height =300
                    FontSize =12
                    TabIndex =40
                    ForeColor =0
                    Name ="LCounty"
                    ControlSource ="LCounty"
                    RowSourceType ="Table/Query"
                    RowSource ="Select County From vw_CCSZ_Lookup Where State ='GA' Group By County, CountyID Or"
                        "der By County;"
                    ColumnWidths ="2160;720"
                    StatusBarText ="link to counties table"
                    OnClick ="[Event Procedure]"
                    GroupTable =20

                    LayoutCachedLeft =15420
                    LayoutCachedTop =4320
                    LayoutCachedWidth =20580
                    LayoutCachedHeight =4620
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =20
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14100
                            Top =4320
                            Width =1260
                            Height =300
                            FontSize =10
                            BorderColor =0
                            Name ="Label43"
                            Caption ="County"
                            GroupTable =20
                            LayoutCachedLeft =14100
                            LayoutCachedTop =4320
                            LayoutCachedWidth =15360
                            LayoutCachedHeight =4620
                            RowStart =8
                            RowEnd =8
                            LayoutGroup =4
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =20
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =21900
                    Width =660
                    FontSize =8
                    BorderColor =12632256
                    Name ="ClientREF"
                    ControlSource ="ClientGroupId"
                    StatusBarText ="link to Clients table"
                    FontName ="Tahoma"
                    GridlineColor =0

                    LayoutCachedLeft =21900
                    LayoutCachedWidth =22560
                    LayoutCachedHeight =240
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =23280
                    Width =720
                    FontSize =8
                    TabIndex =2
                    BorderColor =12632256
                    Name ="LocationID"
                    ControlSource ="LocationID"
                    FontName ="Tahoma"
                    GridlineColor =0

                    LayoutCachedLeft =23280
                    LayoutCachedWidth =24000
                    LayoutCachedHeight =240
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    DecimalPlaces =4
                    SpecialEffect =3
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =23400
                    Top =1980
                    Width =2100
                    Height =420
                    FontSize =18
                    FontWeight =700
                    TabIndex =45
                    Name ="LTaxRate"
                    ControlSource ="LTaxRate"
                    Format ="Percent"
                    ValidationRule =">=0"
                    ValidationText ="Tax rate must be reater than or equal to 0"
                    DefaultValue ="0"

                    LayoutCachedLeft =23400
                    LayoutCachedTop =1980
                    LayoutCachedWidth =25500
                    LayoutCachedHeight =2400
                    RowStart =10
                    RowEnd =10
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =5
                    ForeTint =100.0
                    ForeShade =75.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =3
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =2
                    Left =21120
                    Top =1995
                    Width =2100
                    Height =420
                    FontSize =14
                    FontWeight =700
                    TabIndex =46
                    Name ="LTaxOption"
                    ControlSource ="LTaxOption"
                    RowSourceType ="Value List"
                    RowSource ="\"RR\";\"Upfront\""
                    StatusBarText ="For Optional states, customer selection of Upfront or Rent Receipts"
                    ControlTipText ="For Optional states, customer selection of Upfront or Rent Receipts"

                    LayoutCachedLeft =21120
                    LayoutCachedTop =1995
                    LayoutCachedWidth =23220
                    LayoutCachedHeight =2415
                    RowStart =10
                    RowEnd =10
                    ColumnStart =1
                    ColumnEnd =3
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =5
                    ForeShade =75.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =21120
                            Top =1680
                            Width =2100
                            Height =300
                            FontSize =12
                            BorderColor =0
                            Name ="Label80"
                            Caption ="Tax Option "
                            ControlTipText ="For Optional states, customer selection of Upfront or Rent Receipts"
                            LayoutCachedLeft =21120
                            LayoutCachedTop =1680
                            LayoutCachedWidth =23220
                            LayoutCachedHeight =1980
                            RowStart =10
                            RowEnd =10
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =23400
                    Top =1680
                    Width =2100
                    Height =300
                    FontSize =12
                    BorderColor =0
                    Name ="Label68"
                    Caption ="Location Tax Rate"
                    LayoutCachedLeft =23400
                    LayoutCachedTop =1680
                    LayoutCachedWidth =25500
                    LayoutCachedHeight =1980
                    RowStart =10
                    RowEnd =10
                    ColumnStart =2
                    ColumnEnd =2
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =24000
                    Width =720
                    FontSize =8
                    TabIndex =3
                    BorderColor =12632256
                    Name ="LStateID"
                    ControlSource ="LStateID"
                    FontName ="Tahoma"
                    GridlineColor =0

                    LayoutCachedLeft =24000
                    LayoutCachedWidth =24720
                    LayoutCachedHeight =240
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =24720
                    Width =720
                    FontSize =8
                    TabIndex =4
                    BorderColor =12632256
                    Name ="LCountyID"
                    ControlSource ="LCountyID"
                    FontName ="Tahoma"
                    GridlineColor =0

                    LayoutCachedLeft =24720
                    LayoutCachedWidth =25440
                    LayoutCachedHeight =240
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =25440
                    Width =720
                    FontSize =8
                    TabIndex =5
                    BorderColor =12632256
                    Name ="LCityID"
                    ControlSource ="LCityID"
                    FontName ="Tahoma"
                    GridlineColor =0

                    LayoutCachedLeft =25440
                    LayoutCachedWidth =26160
                    LayoutCachedHeight =240
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =13860
                    Top =6960
                    Width =3495
                    Height =225
                    FontSize =8
                    BorderColor =0
                    Name ="Label179"
                    Caption =" ┌─ Search   ─┬────────┬───────┐"
                    FontName ="Segoe UI"
                    GridlineColor =0
                    LayoutCachedLeft =13860
                    LayoutCachedTop =6960
                    LayoutCachedWidth =17355
                    LayoutCachedHeight =7185
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =11
                    Left =13860
                    Top =7680
                    Width =11232
                    Height =5460
                    FontSize =8
                    TabIndex =59
                    ForeColor =0
                    Name ="UnitsList"
                    RowSourceType ="Table/Query"
                    RowSource ="Select AssetID, MLNO, Schedule, [Unit#], Desc From UnitsParkedLocations Where Ga"
                        "rageLocationID=2341 Order by MLOrig, AssetID;"
                    ColumnWidths ="720;1152;864;1872;432;3240;1008;1008;432;0;0"
                    OnDblClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    AllowValueListEdits =0

                    LayoutCachedLeft =13860
                    LayoutCachedTop =7680
                    LayoutCachedWidth =25092
                    LayoutCachedHeight =13140
                    ColumnEnd =3
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =13860
                    Top =7200
                    Width =720
                    FontSize =9
                    TabIndex =52
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchAssetId"
                    OnChange ="[Event Procedure]"
                    GroupTable =34
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13860
                    LayoutCachedTop =7200
                    LayoutCachedWidth =14580
                    LayoutCachedHeight =7440
                    LayoutGroup =5
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =34
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =18480
                    Top =7200
                    Width =435
                    FontSize =9
                    TabIndex =55
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchStatus"
                    OnChange ="[Event Procedure]"
                    GroupTable =34
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18480
                    LayoutCachedTop =7200
                    LayoutCachedWidth =18915
                    LayoutCachedHeight =7440
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =5
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =34
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13860
                    Top =6600
                    Width =10317
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =51
                    BorderColor =0
                    Name ="Label629"
                    ControlSource ="=IIf([UnitsList].[ListCount]>0,[UnitsList].[ListCount]-1,0) & \" Units at Locati"
                        "on: \" & [Locationsn]"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =13860
                    LayoutCachedTop =6600
                    LayoutCachedWidth =24177
                    LayoutCachedHeight =6885
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15735
                    Top =7200
                    Width =870
                    FontSize =9
                    TabIndex =53
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchUnitNum"
                    OnChange ="[Event Procedure]"
                    GroupTable =34
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15735
                    LayoutCachedTop =7200
                    LayoutCachedWidth =16605
                    LayoutCachedHeight =7440
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =34
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =87
                            TextAlign =2
                            Left =14580
                            Top =7200
                            Width =1155
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BackColor =15527148
                            BorderColor =0
                            ForeColor =12349952
                            Name ="Label309"
                            Caption =" ---"
                            GroupTable =34
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =14580
                            LayoutCachedTop =7200
                            LayoutCachedWidth =15735
                            LayoutCachedHeight =7440
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =5
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =34
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =16605
                    Top =7200
                    Width =1875
                    FontSize =9
                    TabIndex =54
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchVIN"
                    OnChange ="[Event Procedure]"
                    GroupTable =34
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16605
                    LayoutCachedTop =7200
                    LayoutCachedWidth =18480
                    LayoutCachedHeight =7440
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =34
                End
                Begin TextBox
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =22560
                    Width =720
                    FontSize =8
                    TabIndex =1
                    BorderColor =12632256
                    Name ="ClientID"
                    FontName ="Tahoma"
                    GridlineColor =0

                    LayoutCachedLeft =22560
                    LayoutCachedWidth =23280
                    LayoutCachedHeight =240
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15420
                    Top =1440
                    Width =5160
                    Height =300
                    FontSize =12
                    FontWeight =700
                    TabIndex =31
                    ForeColor =0
                    Name ="Location Type"
                    ControlSource ="LocationType"
                    EventProcPrefix ="Location_Type"
                    GroupTable =20

                    LayoutCachedLeft =15420
                    LayoutCachedTop =1440
                    LayoutCachedWidth =20580
                    LayoutCachedHeight =1740
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =4
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =20
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14100
                            Top =1440
                            Width =1260
                            Height =300
                            FontSize =10
                            BorderColor =0
                            Name ="Label3"
                            Caption ="Type"
                            GroupTable =20
                            LayoutCachedLeft =14100
                            LayoutCachedTop =1440
                            LayoutCachedWidth =15360
                            LayoutCachedHeight =1740
                            LayoutGroup =4
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =20
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =16560
                    Top =5760
                    Width =2100
                    Height =420
                    FontSize =12
                    TabIndex =49
                    ForeColor =8210719
                    Name ="LookUpTaxZip"
                    Caption ="Look Up Tax (Zip)"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Get Tax Rate From The Web"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000072727254727272b5 ,
                        0x737373f8727272ff737373f8727272b572727254000000000000000000000000 ,
                        0x0000000000000000000000000000000072727215737373af808080e69f9f9ff7 ,
                        0xa5a5a5fff3f3f3ffa0a0a0ff9f9f9ff7808080e6737373af7272721500000000 ,
                        0x00000000000000000000000072727224747474e5b3b3b3e3f3f3f3ff838383ff ,
                        0xf8f8f8fffffffffff8f8f8ff848484ffefefefffb3b3b3e3747474e572727218 ,
                        0x000000000000000072727203737373d1c1c1c1e9ffffffffc7c7c7ffb4b4b4ff ,
                        0xffffffffffffffffffffffffb4b4b4ffbebebeffffffffffc1c1c1e9757575b2 ,
                        0x00000000000000007272725d858585e2ffffffffffffffff9d9d9dffd6d6d6ff ,
                        0xffffffffffffffffffffffffd6d6d6ff9d9d9dffffffffffffffffff858585e2 ,
                        0x7272725700000000737373b5727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x737373b500000000747474edf0f0f0ffffffffffffffffff797979fff8f8f8ff ,
                        0xfffffffffffffffffffffffff8f8f8ff797979fffffffffffffffffff0f0f0ff ,
                        0x747474eb00000000727272fffcfcfcffffffffffffffffff727272fffdfdfdff ,
                        0xffffffffffffffffffffffffffffffff727272fffffffffffffffffffcfcfcff ,
                        0x737373fa00000000747474edf0f0f0ffffffffffffffffff797979fff3f3f3ff ,
                        0xfffffffffffffffffffffffff8f8f8ff797979fffffffffffffffffff0f0f0ff ,
                        0x747474eb00000000737373b5727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x737373b5000000007272725a858585e2ffffffffffffffff9d9d9dffd6d6d6ff ,
                        0xffffffffffffffffffffffffd6d6d6ff9d9d9dffffffffffffffffff858585e2 ,
                        0x727272570000000072727203737373cebebebee9ffffffffc8c8c8ffb4b4b4ff ,
                        0xffffffffffffffffffffffffb4b4b4ffbebebeffffffffffbebebee9757575b2 ,
                        0x00000000000000000000000072727224747474e5aeaeaee5f2f2f2ff7f7f7fff ,
                        0xf8f8f8fffffffffff8f8f8ff838383ffedededffaeaeaee5747474e772727218 ,
                        0x0000000000000000000000000000000072727215737373af7f7f7fe89f9f9ff7 ,
                        0xa4a4a4fff8f8f8ffa3a3a3ff9f9f9ff77f7f7fe8737373af7272721500000000 ,
                        0x0000000000000000000000000000000000000000000000007272724b72727297 ,
                        0x757575eb727272fe757575eb727272977272724b000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =16560
                    LayoutCachedTop =5760
                    LayoutCachedWidth =18660
                    LayoutCachedHeight =6180
                    PictureCaptionArrangement =5
                    ColumnEnd =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =8210719
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =8210719
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                End
                Begin EmptyCell
                    Left =17265
                    Top =3600
                    Width =3315
                    Height =300
                    Name ="EmptyCell389"
                    GroupTable =20
                    LayoutCachedLeft =17265
                    LayoutCachedTop =3600
                    LayoutCachedWidth =20580
                    LayoutCachedHeight =3900
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =4
                    GroupTable =20
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14100
                    Top =5760
                    Width =2100
                    Height =420
                    FontSize =12
                    TabIndex =48
                    ForeColor =8210719
                    Name ="LookupStateTaxes"
                    Caption ="Sales Tax States"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Get Tax Rate From The Web"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000072727254727272b5 ,
                        0x737373f8727272ff737373f8727272b572727254000000000000000000000000 ,
                        0x0000000000000000000000000000000072727215737373af808080e69f9f9ff7 ,
                        0xa5a5a5fff3f3f3ffa0a0a0ff9f9f9ff7808080e6737373af7272721500000000 ,
                        0x00000000000000000000000072727224747474e5b3b3b3e3f3f3f3ff838383ff ,
                        0xf8f8f8fffffffffff8f8f8ff848484ffefefefffb3b3b3e3747474e572727218 ,
                        0x000000000000000072727203737373d1c1c1c1e9ffffffffc7c7c7ffb4b4b4ff ,
                        0xffffffffffffffffffffffffb4b4b4ffbebebeffffffffffc1c1c1e9757575b2 ,
                        0x00000000000000007272725d858585e2ffffffffffffffff9d9d9dffd6d6d6ff ,
                        0xffffffffffffffffffffffffd6d6d6ff9d9d9dffffffffffffffffff858585e2 ,
                        0x7272725700000000737373b5727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x737373b500000000747474edf0f0f0ffffffffffffffffff797979fff8f8f8ff ,
                        0xfffffffffffffffffffffffff8f8f8ff797979fffffffffffffffffff0f0f0ff ,
                        0x747474eb00000000727272fffcfcfcffffffffffffffffff727272fffdfdfdff ,
                        0xffffffffffffffffffffffffffffffff727272fffffffffffffffffffcfcfcff ,
                        0x737373fa00000000747474edf0f0f0ffffffffffffffffff797979fff3f3f3ff ,
                        0xfffffffffffffffffffffffff8f8f8ff797979fffffffffffffffffff0f0f0ff ,
                        0x747474eb00000000737373b5727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x737373b5000000007272725a858585e2ffffffffffffffff9d9d9dffd6d6d6ff ,
                        0xffffffffffffffffffffffffd6d6d6ff9d9d9dffffffffffffffffff858585e2 ,
                        0x727272570000000072727203737373cebebebee9ffffffffc8c8c8ffb4b4b4ff ,
                        0xffffffffffffffffffffffffb4b4b4ffbebebeffffffffffbebebee9757575b2 ,
                        0x00000000000000000000000072727224747474e5aeaeaee5f2f2f2ff7f7f7fff ,
                        0xf8f8f8fffffffffff8f8f8ff838383ffedededffaeaeaee5747474e772727218 ,
                        0x0000000000000000000000000000000072727215737373af7f7f7fe89f9f9ff7 ,
                        0xa4a4a4fff8f8f8ffa3a3a3ff9f9f9ff77f7f7fe8737373af7272721500000000 ,
                        0x0000000000000000000000000000000000000000000000007272724b72727297 ,
                        0x757575eb727272fe757575eb727272977272724b000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =14100
                    LayoutCachedTop =5760
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =6180
                    PictureCaptionArrangement =5
                    ColumnEnd =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =8210719
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =8210719
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                End
                Begin ToggleButton
                    OverlapFlags =95
                    Left =13860
                    Top =7440
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =56
                    ForeColor =8210719
                    Name ="ToggleSortAssetID1"
                    DefaultValue ="0"
                    Caption ="AssetID"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =34
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13860
                    LayoutCachedTop =7440
                    LayoutCachedWidth =14580
                    LayoutCachedHeight =7680
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =5
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =34
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =14580
                    Top =7440
                    Width =1155
                    Name ="EmptyCell398"
                    GroupTable =34
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14580
                    LayoutCachedTop =7440
                    LayoutCachedWidth =15735
                    LayoutCachedHeight =7680
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    GroupTable =34
                End
                Begin EmptyCell
                    Left =18480
                    Top =7440
                    Width =435
                    Name ="EmptyCell401"
                    GroupTable =34
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18480
                    LayoutCachedTop =7440
                    LayoutCachedWidth =18915
                    LayoutCachedHeight =7680
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =5
                    GroupTable =34
                End
                Begin ToggleButton
                    OverlapFlags =95
                    Left =15735
                    Top =7440
                    Width =870
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =57
                    ForeColor =8210719
                    Name ="ToggleSortUnitNum1"
                    DefaultValue ="0"
                    Caption ="UnitNum"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =34
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15735
                    LayoutCachedTop =7440
                    LayoutCachedWidth =16605
                    LayoutCachedHeight =7680
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =34
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =95
                    Left =16605
                    Top =7440
                    Width =1875
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =58
                    ForeColor =8210719
                    Name ="ToggleSortVIN1"
                    DefaultValue ="0"
                    Caption ="↓↑ VIN"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =34
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16605
                    LayoutCachedTop =7440
                    LayoutCachedWidth =18480
                    LayoutCachedHeight =7680
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =5
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =34
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    Left =1725
                    Top =1665
                    Width =2055
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =25
                    ForeColor =8210719
                    Name ="ToggleSortLocationName1"
                    DefaultValue ="0"
                    Caption ="↓↑ Location"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =1725
                    LayoutCachedTop =1665
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =1905
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =3
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    Left =270
                    Top =1665
                    Width =1440
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =24
                    ForeColor =8210719
                    Name ="ToggleSortClient1"
                    DefaultValue ="0"
                    Caption ="↓↑ Client"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =270
                    LayoutCachedTop =1665
                    LayoutCachedWidth =1710
                    LayoutCachedHeight =1905
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =3
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    Left =5610
                    Top =1665
                    Width =1425
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =27
                    ForeColor =8210719
                    Name ="ToggleSortLCity1"
                    DefaultValue ="0"
                    Caption ="↓↑ City"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =5610
                    LayoutCachedTop =1665
                    LayoutCachedWidth =7035
                    LayoutCachedHeight =1905
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =7
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =3
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    Left =7050
                    Top =1665
                    Width =660
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =28
                    ForeColor =8210719
                    Name ="ToggleSortLState1"
                    DefaultValue ="0"
                    Caption ="State"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =7050
                    LayoutCachedTop =1665
                    LayoutCachedWidth =7710
                    LayoutCachedHeight =1905
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =9
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =3
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =95
                    Left =7710
                    Top =1665
                    Width =1425
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =29
                    ForeColor =8210719
                    Name ="ToggleSortLCounty1"
                    DefaultValue ="0"
                    Caption ="↓↑ County"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =7710
                    LayoutCachedTop =1665
                    LayoutCachedWidth =9135
                    LayoutCachedHeight =1905
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =11
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =3
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    Left =9150
                    Top =1665
                    Width =1170
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =30
                    ForeColor =8210719
                    Name ="ToggleSortLZip1"
                    DefaultValue ="0"
                    Caption ="↓↑ Zip"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =9150
                    LayoutCachedTop =1665
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =1905
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =13
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =3
                    Overlaps =1
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =21000
                    Top =2820
                    Width =5580
                    Height =2580
                    TabIndex =47
                    Name ="LocationTax"
                    SourceObject ="Form.LocationTax"
                    LinkChildFields ="LocationID"
                    LinkMasterFields ="LocationID"

                    LayoutCachedLeft =21000
                    LayoutCachedTop =2820
                    LayoutCachedWidth =26580
                    LayoutCachedHeight =5400
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =21000
                            Top =2460
                            Width =1230
                            Height =315
                            Name ="Label444"
                            Caption ="LocationTax:"
                            LayoutCachedLeft =21000
                            LayoutCachedTop =2460
                            LayoutCachedWidth =22230
                            LayoutCachedHeight =2775
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =93
                    Left =3795
                    Top =1665
                    Width =1800
                    Height =240
                    FontSize =8
                    FontWeight =400
                    TabIndex =26
                    ForeColor =8210719
                    Name ="ToggleSortShort1"
                    DefaultValue ="0"
                    Caption ="↓↑ Short"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =3795
                    LayoutCachedTop =1665
                    LayoutCachedWidth =5595
                    LayoutCachedHeight =1905
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    GroupTable =3
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3795
                    Top =1410
                    Width =1484
                    FontSize =9
                    TabIndex =14
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchShort"
                    Tag ="SearchX"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =3795
                    LayoutCachedTop =1410
                    LayoutCachedWidth =5279
                    LayoutCachedHeight =1650
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =5280
                    Top =1410
                    Width =315
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    ForeColor =0
                    Name ="SearchShort_Button"
                    Caption ="Search Short"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to Search Master Lease"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af2b17d4aff ,
                        0xb17d4a8b00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af2 ,
                        0xb17d4affb17d4a7e000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4a5bb17d4aefb17d4a79000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4aeab17d4affb17d4affb17d4a99b17d4a0400000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb17d4af3 ,
                        0xb17d4affb17d4a7db17d4a41b17d4af1b17d4affb17d4affb17d4ab6b17d4a0c ,
                        0x0000000000000000000000000000000000000000b17d4a3cb17d4af3b17d4aff ,
                        0xb17d4a8a0000000000000000b17d4a2eb17d4ae6b17d4affb17d4affb17d4ace ,
                        0xb17d4a7db17d4adcb17d4afbb17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x00000000000000000000000000000000b17d4a1fb17d4ad7b17d4affb17d4aff ,
                        0xb27f4cfccdac8cf6dcc4adffcdac8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x0000000000000000000000000000000000000000b17d4a12b17d4ad6b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded2ffb17e4bfdb17d4a8100000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adbcdab8bf7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4af9dbc3acff ,
                        0xffffffffffffffffffffffffffffffffffffffffdbc3acffb17d4af900000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4adccdab8af7 ,
                        0xffffffffffffffffffffffffffffffffffffffffcdab8bf7b17d4ade00000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a83b17e4bfd ,
                        0xebded1ffffffffffffffffffffffffffebded1ffb17e4bfdb17d4a8000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a0ab17d4ac5 ,
                        0xb27f4cfccdab8bf6dbc3abffcdab8bf6b27f4cfcb17d4ac5b17d4a0a00000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a09 ,
                        0xb17d4a81b17d4adcb17d4af9b17d4adcb17d4a81b17d4a090000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5280
                    LayoutCachedTop =1410
                    LayoutCachedWidth =5595
                    LayoutCachedHeight =1650
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =3
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "LocationTaxManager.cls"
