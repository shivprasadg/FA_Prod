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
    Width =28800
    DatasheetFontHeight =11
    ItemSuffix =87
    Left =4845
    Top =3315
    Right =-31636
    Bottom =17835
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
            Height =735
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
                    Name ="WorkingWith"
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
                    Width =28800
                    Height =195
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="DrkRedOnLtGrey_Shadow"
                    GridlineColor =0
                    HorizontalAnchor =2

                    LayoutCachedTop =540
                    LayoutCachedWidth =28800
                    LayoutCachedHeight =735
                    TabIndex =6
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
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
                    ControlSource ="=[AuditManager].[Form]![UnitCount]"

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
                    OverlapFlags =247
                    Top =60
                    Width =4365
                    Height =570
                    FontSize =20
                    Name ="Label31"
                    Caption ="Audit Manager"
                    FontName ="Segoe UI"
                    LayoutCachedTop =60
                    LayoutCachedWidth =4365
                    LayoutCachedHeight =630
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =22515
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

                    LayoutCachedLeft =22515
                    LayoutCachedTop =120
                    LayoutCachedWidth =23265
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
                    Left =21960
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

                    LayoutCachedLeft =21960
                    LayoutCachedTop =120
                    LayoutCachedWidth =22515
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =13575
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
                    TabIndex =11
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
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =3195
                    Top =255
                    Width =2559
                    Height =825
                    FontSize =9
                    TabIndex =1
                    Name ="LeaseList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MLID, MLNo, MLOrig, ClientID From vw_SixKeys GROUP BY MLID, MLNo, MLOrig,"
                        " ClientID Having ClientID =3327 ORDER BY MLOrig;"
                    ColumnWidths ="0;1872;576"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3195
                    LayoutCachedTop =255
                    LayoutCachedWidth =5754
                    LayoutCachedHeight =1080
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =3225
                    Top =1320
                    Width =2559
                    Height =2145
                    FontSize =9
                    TabIndex =2
                    Name ="ScheduleList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SchID, Schedule, MLID, Count(AssetID) AS Unt, ClientID, MLNo, MLOrig FROM"
                        " vw_SixKeys GROUP BY SchID, Schedule, MLID, ClientID, MLNo, MLOrig HAVING MLID=3"
                        "6 ORDER BY CDbl(NumInString(Left(Schedule,3))), Schedule;"
                    ColumnWidths ="0;1872;0;576"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3225
                    LayoutCachedTop =1320
                    LayoutCachedWidth =5784
                    LayoutCachedHeight =3465
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =3195
                    Top =15
                    Width =1350
                    Height =300
                    FontSize =9
                    FontWeight =700
                    Name ="Label22"
                    Caption ="Leases:"
                    LayoutCachedLeft =3195
                    LayoutCachedTop =15
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =3180
                    Top =1080
                    Width =1350
                    Height =225
                    FontSize =9
                    FontWeight =700
                    Name ="Label23"
                    Caption ="Schedules:"
                    LayoutCachedLeft =3180
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4530
                    LayoutCachedHeight =1305
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
                        " MYear, UnitGroup FROM vw_SixKeys WHERE SchID=293 ORDER BY AssetID;"
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
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =3
                    Top =495
                    Width =3060
                    Height =4965
                    FontSize =9
                    Name ="ClientList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT clientID, ClientShortName, CompanyName, Count(AssetID) AS [Unit Count] Fr"
                        "om vw_SixKeys GROUP BY clientID, ClientShortName, CompanyName ORDER BY CompanyNa"
                        "me;"
                    ColumnWidths ="0;2160;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedTop =495
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =5460
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =3180
                    Top =3720
                    Width =2559
                    Height =1740
                    FontSize =9
                    TabIndex =10
                    Name ="GroupList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SgrpID, UnitGroup, Count(AssetID) AS Unt, SchID FROM vw_SixKeys GROUP BY "
                        "SgrpID, UnitGroup,SchID HAVING SchID=293 ORDER BY UnitGroup;"
                    ColumnWidths ="0;1872;576;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3180
                    LayoutCachedTop =3720
                    LayoutCachedWidth =5739
                    LayoutCachedHeight =5460
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3180
                            Top =3480
                            Width =2559
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label75"
                            Caption ="Groups:"
                            LayoutCachedLeft =3180
                            LayoutCachedTop =3480
                            LayoutCachedWidth =5739
                            LayoutCachedHeight =3720
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
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
                Begin Subform
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =5940
                    Top =240
                    Width =16320
                    Height =13260
                    TabIndex =12
                    Name ="AuditManagerLog"
                    SourceObject ="Form.AuditManagerLog"

                    LayoutCachedLeft =5940
                    LayoutCachedTop =240
                    LayoutCachedWidth =22260
                    LayoutCachedHeight =13500
                End
                Begin Label
                    OverlapFlags =93
                    Left =5940
                    Width =1350
                    Height =300
                    FontSize =9
                    FontWeight =700
                    Name ="Label86"
                    Caption ="Changes:"
                    LayoutCachedLeft =5940
                    LayoutCachedWidth =7290
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =2
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    Visible = NotDefault
                    OverlapFlags =215
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =7
                    Left =22440
                    Top =225
                    Width =6300
                    Height =10275
                    FontSize =9
                    TabIndex =13
                    Name ="UnitRevDetail"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="864;2160;2160;0;0;1440;0"

                    LayoutCachedLeft =22440
                    LayoutCachedTop =225
                    LayoutCachedWidth =28740
                    LayoutCachedHeight =10500
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =22440
                            Width =6300
                            Height =240
                            FontSize =9
                            ForeColor =8872517
                            Name ="UnitRevDetailLabel"
                            Caption ="Unit Detail"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =22440
                            LayoutCachedWidth =28740
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =22440
                    Top =11550
                    Width =6300
                    Height =1605
                    FontSize =9
                    TabIndex =14
                    BorderColor =0
                    ForeColor =6108695
                    Name ="ReasonNote"
                    StatusBarText ="1st option lrf"
                    ValidationText ="Rent must be >=0"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =22440
                    LayoutCachedTop =11550
                    LayoutCachedWidth =28740
                    LayoutCachedHeight =13155
                    RowStart =1
                    RowEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =22440
                            Top =11310
                            Width =6300
                            Height =240
                            FontSize =9
                            ForeColor =8872517
                            Name ="ReasonNoteLabel"
                            Caption ="Action Reason / Note"
                            FontName ="Segoe UI"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =22440
                            LayoutCachedTop =11310
                            LayoutCachedWidth =28740
                            LayoutCachedHeight =11550
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =22440
                    Top =10860
                    Width =6300
                    Height =420
                    FontSize =9
                    TabIndex =15
                    BorderColor =0
                    ForeColor =6108695
                    Name ="ReasonAction"
                    StatusBarText ="1st option lrf"
                    DefaultValue ="0"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =22440
                    LayoutCachedTop =10860
                    LayoutCachedWidth =28740
                    LayoutCachedHeight =11280
                    RowStart =1
                    RowEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =22440
                            Top =10620
                            Width =6300
                            Height =240
                            FontSize =9
                            ForeColor =8872517
                            Name ="Label4385"
                            Caption ="Action Desc"
                            FontName ="Segoe UI"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =22440
                            LayoutCachedTop =10620
                            LayoutCachedWidth =28740
                            LayoutCachedHeight =10860
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =4
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =23460
                    Top =13230
                    Width =1455
                    Height =210
                    FontSize =8
                    TabIndex =16
                    Name ="ReasonINDX"
                    Format ="Percent"
                    StatusBarText ="int rt for int/pd bill"
                    Tag ="StayOpen"

                    LayoutCachedLeft =23460
                    LayoutCachedTop =13230
                    LayoutCachedWidth =24915
                    LayoutCachedHeight =13440
                End
            End
        End
        Begin FormFooter
            Height =225
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18030
                    Width =705
                    Height =225
                    FontSize =9
                    BorderColor =9211020
                    ForeColor =9211020
                    Name ="SchID"

                    LayoutCachedLeft =18030
                    LayoutCachedWidth =18735
                    LayoutCachedHeight =225
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =17460
                            Width =570
                            Height =225
                            FontSize =9
                            BorderColor =9211020
                            ForeColor =9211020
                            Name ="Label19"
                            Caption ="SchID:"
                            LayoutCachedLeft =17460
                            LayoutCachedWidth =18030
                            LayoutCachedHeight =225
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19530
                    Width =705
                    Height =225
                    FontSize =9
                    TabIndex =1
                    BorderColor =9211020
                    ForeColor =9211020
                    Name ="GroupID"

                    LayoutCachedLeft =19530
                    LayoutCachedWidth =20235
                    LayoutCachedHeight =225
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =18735
                            Width =795
                            Height =225
                            FontSize =9
                            BorderColor =9211020
                            ForeColor =9211020
                            Name ="Label81"
                            Caption ="GroupID:"
                            LayoutCachedLeft =18735
                            LayoutCachedWidth =19530
                            LayoutCachedHeight =225
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =119
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =21015
                    Width =705
                    Height =225
                    FontSize =9
                    TabIndex =2
                    BorderColor =9211020
                    ForeColor =9211020
                    Name ="UnitId"

                    LayoutCachedLeft =21015
                    LayoutCachedWidth =21720
                    LayoutCachedHeight =225
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =20220
                            Width =795
                            Height =225
                            FontSize =9
                            BorderColor =9211020
                            ForeColor =9211020
                            Name ="Label85"
                            Caption ="UnitID:"
                            LayoutCachedLeft =20220
                            LayoutCachedWidth =21015
                            LayoutCachedHeight =225
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "AuditManager.cls"
