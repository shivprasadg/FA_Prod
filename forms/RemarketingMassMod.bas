Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =27936
    DatasheetFontHeight =11
    ItemSuffix =83
    Left =4620
    Top =1575
    Right =-32731
    Bottom =22575
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xf80cbbca5255e440
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
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
        Begin Tab
            TextFontFamily =18
            FontSize =11
            FontName ="Cambria"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
            Height =765
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4440
                    Top =300
                    Width =7500
                    Height =360
                    ColumnOrder =4
                    FontSize =10
                    FontWeight =700
                    Name ="SchNo"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =4440
                    LayoutCachedTop =300
                    LayoutCachedWidth =11940
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4440
                            Top =120
                            Width =1800
                            Height =180
                            FontSize =8
                            Name ="Label13"
                            Caption ="You're Working with:"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =4440
                            LayoutCachedTop =120
                            LayoutCachedWidth =6240
                            LayoutCachedHeight =300
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16470
                    Width =480
                    Height =210
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =1
                    ForeColor =-2147483615
                    Name ="Text47"
                    ControlSource ="=[Units_List].[ListCount]-1"

                    LayoutCachedLeft =16470
                    LayoutCachedWidth =16950
                    LayoutCachedHeight =210
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =15975
                    Top =15
                    Width =450
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =13056
                    Name ="Label49"
                    Caption ="Units"
                    LayoutCachedLeft =15975
                    LayoutCachedTop =15
                    LayoutCachedWidth =16425
                    LayoutCachedHeight =210
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16455
                    Top =225
                    Width =480
                    Height =210
                    ColumnOrder =3
                    FontSize =8
                    TabIndex =2
                    ForeColor =-2147483615
                    Name ="Text63"
                    ControlSource ="=[Units_List].[ItemsSelected].[Count]"

                    LayoutCachedLeft =16455
                    LayoutCachedTop =225
                    LayoutCachedWidth =16935
                    LayoutCachedHeight =435
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =15960
                    Top =240
                    Width =450
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =13056
                    Name ="Label64"
                    Caption ="Units"
                    LayoutCachedLeft =15960
                    LayoutCachedTop =240
                    LayoutCachedWidth =16410
                    LayoutCachedHeight =435
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =540
                    Width =27936
                    Height =195
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="DrkRedOnLtGrey_Shadow"
                    GridlineColor =0
                    HorizontalAnchor =2

                    LayoutCachedTop =540
                    LayoutCachedWidth =27936
                    LayoutCachedHeight =735
                    TabIndex =8
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13765
                    Top =240
                    Width =689
                    Height =300
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    BorderColor =1643706
                    Name ="UnitCount"
                    ControlSource ="=[ReMarketingMassMod_sub_SUFParked].[Form]![UnitCount]"

                    LayoutCachedLeft =13765
                    LayoutCachedTop =240
                    LayoutCachedWidth =14454
                    LayoutCachedHeight =540
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =255
                            TextAlign =2
                            Left =13740
                            Top =60
                            Width =750
                            Height =540
                            FontSize =8
                            BackColor =1643706
                            BorderColor =1643706
                            Name ="Label77"
                            Caption ="Units"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =13740
                            LayoutCachedTop =60
                            LayoutCachedWidth =14490
                            LayoutCachedHeight =600
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =255
                    Top =60
                    Width =4365
                    Height =570
                    FontSize =20
                    Name ="Label31"
                    Caption ="Mass Modification Form"
                    FontName ="Segoe UI"
                    LayoutCachedTop =60
                    LayoutCachedWidth =4365
                    LayoutCachedHeight =630
                End
                Begin Label
                    OverlapFlags =247
                    Left =30
                    Width =1260
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =8872517
                    Name ="Label294"
                    Caption ="Remarketing"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =30
                    LayoutCachedWidth =1290
                    LayoutCachedHeight =240
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =20235
                    Top =120
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =4
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =20235
                    LayoutCachedTop =120
                    LayoutCachedWidth =20985
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
                Begin CommandButton
                    OverlapFlags =215
                    Left =19680
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =5
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

                    LayoutCachedLeft =19680
                    LayoutCachedTop =120
                    LayoutCachedWidth =20235
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
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =10200
                    Top =120
                    Width =3060
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="OpenExitOptions"
                    Caption ="EXIT Options Form"
                    OnClick ="[Event Procedure]"
                    Tag ="UAT"
                    ControlTipText ="Open EXIT Option Form with the Group shown on the button"
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =15
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10200
                    LayoutCachedTop =120
                    LayoutCachedWidth =13260
                    LayoutCachedHeight =360
                    ColumnStart =20
                    ColumnEnd =20
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    HoverThemeColorIndex =6
                    HoverShade =50.0
                    PressedColor =8210719
                    HoverForeColor =6108695
                    PressedForeThemeColorIndex =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17700
                    Top =120
                    Width =480
                    Height =210
                    FontSize =8
                    TabIndex =7
                    ForeColor =-2147483615
                    Name ="ClientGroupID"

                    LayoutCachedLeft =17700
                    LayoutCachedTop =120
                    LayoutCachedWidth =18180
                    LayoutCachedHeight =330
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =19980
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Top =225
                    Width =3060
                    Height =270
                    FontSize =9
                    TabIndex =12
                    BorderColor =967423
                    Name ="SearchClient"
                    OnChange ="[Event Procedure]"

                    LayoutCachedTop =225
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =495
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Width =3060
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label668"
                            Caption ="Type Search HERE:"
                            LayoutCachedWidth =3060
                            LayoutCachedHeight =240
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =120
                    Top =5550
                    Width =5475
                    Height =225
                    FontSize =9
                    FontWeight =700
                    Name ="Label58"
                    Caption ="┌──────  Search ───────────┬────────────┬──────┬────────────┐"
                    LayoutCachedLeft =120
                    LayoutCachedTop =5550
                    LayoutCachedWidth =5595
                    LayoutCachedHeight =5775
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Left =3945
                    Top =5580
                    Width =691
                    Height =225
                    FontSize =8
                    FontWeight =700
                    Name ="ClearLabel"
                    Caption ="◄Clear►"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =3945
                    LayoutCachedTop =5580
                    LayoutCachedWidth =4636
                    LayoutCachedHeight =5805
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    ForeTint =60.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =9
                    Left =30
                    Top =6060
                    Width =5730
                    Height =7500
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    BoundColumn =2
                    BackColor =15527148
                    Name ="Units_List"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SchID, sGrpID As GroupID, AssetID, UnitNum, VIN, [Status], Make, SubType,"
                        " MYear, UnitGroup FROM vw_SixKeys WHERE [Status] <> 'B' AND SchID=1545 Order BY "
                        "AssetID;"
                    ColumnWidths ="0;0;720;720;2016;288;1152;720;0"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =30
                    LayoutCachedTop =6060
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =13560
                    BorderThemeColorIndex =2
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =768
                    Top =5808
                    Width =720
                    FontSize =9
                    TabIndex =4
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchUnit"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =768
                    LayoutCachedTop =5808
                    LayoutCachedWidth =1488
                    LayoutCachedHeight =6048
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =95
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1500
                    Top =5808
                    Width =2016
                    FontSize =9
                    TabIndex =5
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchVin"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1500
                    LayoutCachedTop =5808
                    LayoutCachedWidth =3516
                    LayoutCachedHeight =6048
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3546
                    Top =5190
                    Width =2214
                    Height =315
                    FontSize =10
                    TabIndex =2
                    Name ="GetSelectedUnits"
                    Caption ="▼Selected (15) Units ►"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3546
                    LayoutCachedTop =5190
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =5505
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =4
                    BackShade =75.0
                    BorderColor =6108695
                    HoverThemeColorIndex =4
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1050
                    Top =5190
                    Width =2400
                    Height =315
                    FontSize =10
                    TabIndex =1
                    Name ="GetAll"
                    Caption ="▼All 15 Listed Units ►"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1050
                    LayoutCachedTop =5190
                    LayoutCachedWidth =3450
                    LayoutCachedHeight =5505
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderColor =6108695
                    HoverThemeColorIndex =6
                    HoverShade =75.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =36
                    Top =5808
                    Width =720
                    FontSize =9
                    TabIndex =3
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchAsset"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =36
                    LayoutCachedTop =5808
                    LayoutCachedWidth =756
                    LayoutCachedHeight =6048
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3528
                    Top =5808
                    Width =288
                    FontSize =9
                    TabIndex =6
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchStatus"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3528
                    LayoutCachedTop =5808
                    LayoutCachedWidth =3816
                    LayoutCachedHeight =6048
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4992
                    Top =5808
                    Width =828
                    FontSize =9
                    TabIndex =8
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchType"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4992
                    LayoutCachedTop =5808
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =6048
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =5
                    Top =495
                    Width =3060
                    Height =4575
                    FontSize =9
                    Name ="ClientList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.clientID, vw_SixKeys.ClientShortName, vw_SixKeys.CompanyName, "
                        "Count(vw_SixKeys.AssetID) AS [Unit Count], vw_SixKeys.ClientGroupId FROM vw_SixK"
                        "eys GROUP BY vw_SixKeys.clientID, vw_SixKeys.ClientShortName, vw_SixKeys.Company"
                        "Name, vw_SixKeys.ClientGroupId ORDER BY vw_SixKeys.CompanyName;"
                    ColumnWidths ="0;2160;0;0;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedTop =495
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =5070
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                End
                Begin Tab
                    OverlapFlags =85
                    Left =5895
                    Width =22020
                    Height =13575
                    FontSize =10
                    TabIndex =10
                    Name ="RMKTTabControl"
                    FontName ="Segoe UI"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =5895
                    LayoutCachedWidth =27915
                    LayoutCachedHeight =13575
                    ThemeFontIndex =-1
                    Shape =0
                    BackColor =16777215
                    BackShade =100.0
                    BorderColor =9211020
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeColor =8210719
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =5970
                            Top =418
                            Width =21870
                            Height =13082
                            Name ="SUF - Parked - Inspection Inv"
                            EventProcPrefix ="SUF___Parked___Inspection_Inv"
                            LayoutCachedLeft =5970
                            LayoutCachedTop =418
                            LayoutCachedWidth =27840
                            LayoutCachedHeight =13500
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    OldBorderStyle =0
                                    Left =5976
                                    Top =418
                                    Width =21510
                                    Height =13080
                                    Name ="ReMarketingMassMod_sub_SUFParked"
                                    SourceObject ="Form.ReMarketingMassMod_sub_SUFParked"
                                    HorizontalAnchor =2
                                    VerticalAnchor =2
                                    FilterOnEmptyMaster =0

                                    LayoutCachedLeft =5976
                                    LayoutCachedTop =418
                                    LayoutCachedWidth =27486
                                    LayoutCachedHeight =13498
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5970
                            Top =420
                            Width =21870
                            Height =13080
                            Name ="Inspection2"
                            Caption ="Inspections && Links"
                            LayoutCachedLeft =5970
                            LayoutCachedTop =420
                            LayoutCachedWidth =27840
                            LayoutCachedHeight =13500
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5970
                                    Top =420
                                    Width =18576
                                    Height =12816
                                    Name ="ReMarketingMassMod_sub_Inspection2"
                                    SourceObject ="Form.ReMarketingMassMod_sub_Inspection2"

                                    LayoutCachedLeft =5970
                                    LayoutCachedTop =420
                                    LayoutCachedWidth =24546
                                    LayoutCachedHeight =13236
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5970
                            Top =418
                            Width =21870
                            Height =13082
                            Name ="SetUnitValues"
                            Caption ="Set Unit Values / Bank Share"
                            LayoutCachedLeft =5970
                            LayoutCachedTop =418
                            LayoutCachedWidth =27840
                            LayoutCachedHeight =13500
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5976
                                    Top =418
                                    Width =19800
                                    Height =12816
                                    Name ="ReMarketingMassMod_sub_FMVPenalty"
                                    SourceObject ="Form.ReMarketingMassMod_sub_FMVPenalty"

                                    LayoutCachedLeft =5976
                                    LayoutCachedTop =418
                                    LayoutCachedWidth =25776
                                    LayoutCachedHeight =13234
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5970
                            Top =418
                            Width =21870
                            Height =13082
                            Name ="Sold - Invoicing"
                            EventProcPrefix ="Sold___Invoicing"
                            LayoutCachedLeft =5970
                            LayoutCachedTop =418
                            LayoutCachedWidth =27840
                            LayoutCachedHeight =13500
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5976
                                    Top =418
                                    Width =19776
                                    Height =12816
                                    Name ="ReMarketingMassMod_sub_SoldInvoice"
                                    SourceObject ="Form.ReMarketingMassMod_sub_SoldInvoice"

                                    LayoutCachedLeft =5976
                                    LayoutCachedTop =418
                                    LayoutCachedWidth =25752
                                    LayoutCachedHeight =13234
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5970
                            Top =420
                            Width =21870
                            Height =13080
                            Name ="Buy Out - Interim Funding"
                            EventProcPrefix ="Buy_Out___Interim_Funding"
                            LayoutCachedLeft =5970
                            LayoutCachedTop =420
                            LayoutCachedWidth =27840
                            LayoutCachedHeight =13500
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =6000
                                    Top =420
                                    Width =19767
                                    Height =12816
                                    Name ="ReMarketingMassMod_sub_BuyOutView"
                                    SourceObject ="Form.ReMarketingMassMod_sub_BuyOutView"

                                    LayoutCachedLeft =6000
                                    LayoutCachedTop =420
                                    LayoutCachedWidth =25767
                                    LayoutCachedHeight =13236
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5970
                            Top =420
                            Width =21870
                            Height =13080
                            Name ="Titles"
                            Caption ="Titles Check In"
                            LayoutCachedLeft =5970
                            LayoutCachedTop =420
                            LayoutCachedWidth =27840
                            LayoutCachedHeight =13500
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5999
                                    Top =420
                                    Width =21837
                                    Height =12816
                                    Name ="RemarketingMassMod_sub_Titles"
                                    SourceObject ="Form.RemarketingMassMod_sub_Titles"

                                    LayoutCachedLeft =5999
                                    LayoutCachedTop =420
                                    LayoutCachedWidth =27836
                                    LayoutCachedHeight =13236
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =3
                    Left =2460
                    Top =15
                    Width =600
                    Height =225
                    FontSize =8
                    FontWeight =700
                    Name ="ClearClientSearch"
                    Caption ="▼Clear"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =2460
                    LayoutCachedTop =15
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =240
                    ForeThemeColorIndex =2
                    ForeTint =60.0
                End
                Begin Tab
                    OverlapFlags =247
                    TextFontFamily =34
                    Left =3045
                    Width =2790
                    Height =5115
                    FontSize =9
                    TabIndex =11
                    Name ="TabCtl131"
                    FontName ="Calibri"
                    GridlineColor =0

                    LayoutCachedLeft =3045
                    LayoutCachedWidth =5835
                    LayoutCachedHeight =5115
                    ThemeFontIndex =-1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    BackColor =15921906
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =95.0
                    HoverThemeColorIndex =-1
                    PressedColor =14277081
                    PressedShade =85.0
                    HoverForeColor =-2147483617
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =2
                    PressedForeTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =60.0
                    Begin
                        Begin Page
                            OverlapFlags =119
                            Left =3120
                            Top =420
                            Width =2640
                            Height =4620
                            Name ="UnitDesc"
                            Caption ="Lease | Sch | Grp"
                            LayoutCachedLeft =3120
                            LayoutCachedTop =420
                            LayoutCachedWidth =5760
                            LayoutCachedHeight =5040
                            Begin
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =3
                                    Left =3195
                                    Top =420
                                    Width =2559
                                    Height =720
                                    FontSize =9
                                    Name ="LeaseList"
                                    RowSourceType ="Table/Query"
                                    ColumnWidths ="0;1872;576"
                                    OnClick ="[Event Procedure]"
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =3195
                                    LayoutCachedTop =420
                                    LayoutCachedWidth =5754
                                    LayoutCachedHeight =1140
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =85.0
                                End
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =4
                                    Left =3195
                                    Top =1365
                                    Width =2559
                                    Height =1875
                                    FontSize =9
                                    TabIndex =1
                                    Name ="ScheduleList"
                                    RowSourceType ="Table/Query"
                                    ColumnWidths ="0;1872;0;576"
                                    OnClick ="[Event Procedure]"
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =3195
                                    LayoutCachedTop =1365
                                    LayoutCachedWidth =5754
                                    LayoutCachedHeight =3240
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =85.0
                                End
                                Begin Label
                                    OverlapFlags =255
                                    Left =3195
                                    Top =1140
                                    Width =1350
                                    Height =225
                                    FontSize =9
                                    FontWeight =700
                                    Name ="Label23"
                                    Caption ="Schedules:"
                                    LayoutCachedLeft =3195
                                    LayoutCachedTop =1140
                                    LayoutCachedWidth =4545
                                    LayoutCachedHeight =1365
                                    ForeThemeColorIndex =2
                                End
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =4
                                    Left =3195
                                    Top =3480
                                    Width =2559
                                    Height =1500
                                    FontSize =9
                                    TabIndex =2
                                    Name ="GroupList"
                                    RowSourceType ="Table/Query"
                                    ColumnWidths ="0;1872;576;0"
                                    OnClick ="[Event Procedure]"
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =3195
                                    LayoutCachedTop =3480
                                    LayoutCachedWidth =5754
                                    LayoutCachedHeight =4980
                                    BackThemeColorIndex =1
                                    BorderThemeColorIndex =1
                                    BorderShade =85.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =3195
                                            Top =3240
                                            Width =2559
                                            Height =240
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label75"
                                            Caption ="Groups:"
                                            LayoutCachedLeft =3195
                                            LayoutCachedTop =3240
                                            LayoutCachedWidth =5754
                                            LayoutCachedHeight =3480
                                            ForeThemeColorIndex =2
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =3113
                            Top =413
                            Width =2647
                            Height =4632
                            Name ="Draws"
                            Caption ="Locations"
                            LayoutCachedLeft =3113
                            LayoutCachedTop =413
                            LayoutCachedWidth =5760
                            LayoutCachedHeight =5045
                            Begin
                                Begin ListBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =3195
                                    Top =720
                                    Width =2549
                                    Height =4320
                                    FontSize =9
                                    BorderColor =14211288
                                    Name ="LocationList"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Locations.LocationID, Locations.Locationsn FROM Locations WHERE (((Locati"
                                        "ons.ClientGroupId)=[Forms]![RemarketingMassMod]![ClientGroupID]));"
                                    ColumnWidths ="0;2448"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="2016 Freightliner Sleeper Tractor (D)"
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =3195
                                    LayoutCachedTop =720
                                    LayoutCachedWidth =5744
                                    LayoutCachedHeight =5040
                                    BackThemeColorIndex =1
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =3192
                                    Top =420
                                    Width =2544
                                    FontSize =9
                                    TabIndex =1
                                    BackColor =13434879
                                    BorderColor =967423
                                    Name ="SearchLocation"
                                    OnChange ="[Event Procedure]"
                                    GroupTable =5
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =3192
                                    LayoutCachedTop =420
                                    LayoutCachedWidth =5736
                                    LayoutCachedHeight =660
                                    LayoutGroup =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =5
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3828
                    Top =5808
                    Width =1152
                    FontSize =9
                    TabIndex =7
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchMake"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3828
                    LayoutCachedTop =5808
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =6048
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Top =5190
                    Width =979
                    Height =315
                    FontWeight =400
                    TabIndex =13
                    ForeColor =8355711
                    Name ="PickModeToggle"
                    DefaultValue ="0"
                    Caption ="▼Pick On"
                    OnClick ="[Event Procedure]"
                    GridlineColor =0

                    LayoutCachedTop =5190
                    LayoutCachedWidth =979
                    LayoutCachedHeight =5505
                    ForeTint =50.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15060409
                    BackTint =40.0
                    OldBorderStyle =1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =8
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedColor =0
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                End
            End
        End
        Begin FormFooter
            Height =270
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =23550
                    Width =705
                    Height =225
                    FontSize =9
                    BorderColor =9211020
                    ForeColor =9211020
                    Name ="SchID"

                    LayoutCachedLeft =23550
                    LayoutCachedWidth =24255
                    LayoutCachedHeight =225
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            Left =22920
                            Width =570
                            Height =225
                            FontSize =9
                            BorderColor =9211020
                            ForeColor =9211020
                            Name ="Label19"
                            Caption ="SchID:"
                            LayoutCachedLeft =22920
                            LayoutCachedWidth =23490
                            LayoutCachedHeight =225
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "RemarketingMassMod.cls"
