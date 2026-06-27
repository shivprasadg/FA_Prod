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
    Width =28095
    DatasheetFontHeight =11
    ItemSuffix =136
    Left =8820
    Top =3840
    Right =20010
    Bottom =18390
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xf80cbbca5255e440
    End
    OnCurrent ="[Event Procedure]"
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
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =540
                    Width =27270
                    Height =60
                    BorderColor =10855845
                    Name ="Box22"
                    HorizontalAnchor =2
                    LayoutCachedTop =540
                    LayoutCachedWidth =27270
                    LayoutCachedHeight =600
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15685
                    Top =270
                    Width =1274
                    Height =435
                    ColumnOrder =3
                    FontSize =20
                    FontWeight =700
                    BorderColor =1643706
                    Name ="UnitCount"
                    ControlSource ="=ExitOptions_subform_Exchange.Form!UnitCount"

                    LayoutCachedLeft =15685
                    LayoutCachedTop =270
                    LayoutCachedWidth =16959
                    LayoutCachedHeight =705
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =223
                            TextFontCharSet =238
                            TextAlign =2
                            Left =15660
                            Top =60
                            Width =1335
                            Height =675
                            FontSize =9
                            FontWeight =700
                            BorderColor =-2147483615
                            ForeColor =16777215
                            Name ="Label63"
                            Caption ="Units Shown"
                            LayoutCachedLeft =15660
                            LayoutCachedTop =60
                            LayoutCachedWidth =16995
                            LayoutCachedHeight =735
                            BackThemeColorIndex =2
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Top =60
                    Width =2520
                    Height =540
                    FontSize =20
                    FontWeight =700
                    Name ="Label26"
                    Caption ="EXIT Options"
                    FontName ="Segoe UI"
                    LayoutCachedTop =60
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Top =210
                    Width =12180
                    Height =390
                    ColumnOrder =4
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    Name ="SchNo"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =210
                    LayoutCachedWidth =14880
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =20775
                    Top =75
                    Width =480
                    Height =210
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =3
                    ForeColor =-2147483615
                    Name ="Text47"
                    ControlSource ="=[Units_List].[ListCount]-1"

                    LayoutCachedLeft =20775
                    LayoutCachedTop =75
                    LayoutCachedWidth =21255
                    LayoutCachedHeight =285
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =20280
                    Top =90
                    Width =450
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =13056
                    Name ="Label49"
                    Caption ="Units"
                    LayoutCachedLeft =20280
                    LayoutCachedTop =90
                    LayoutCachedWidth =20730
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =23850
                    Top =60
                    Width =480
                    Height =210
                    ColumnOrder =1
                    FontSize =8
                    TabIndex =2
                    ForeColor =-2147483615
                    Name ="Text63"
                    ControlSource ="=[Units_List].[ItemsSelected].[Count]"

                    LayoutCachedLeft =23850
                    LayoutCachedTop =60
                    LayoutCachedWidth =24330
                    LayoutCachedHeight =270
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =23355
                    Top =75
                    Width =450
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =13056
                    Name ="Label64"
                    Caption ="Units"
                    LayoutCachedLeft =23355
                    LayoutCachedTop =75
                    LayoutCachedWidth =23805
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =22470
                    Top =60
                    Width =705
                    Height =225
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =4
                    Name ="SchID"

                    LayoutCachedLeft =22470
                    LayoutCachedTop =60
                    LayoutCachedWidth =23175
                    LayoutCachedHeight =285
                    BackThemeColorIndex =4
                    BackShade =50.0
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =21840
                            Top =60
                            Width =570
                            Height =225
                            FontSize =9
                            Name ="Label19"
                            Caption ="SchID:"
                            LayoutCachedLeft =21840
                            LayoutCachedTop =60
                            LayoutCachedWidth =22410
                            LayoutCachedHeight =285
                            BackThemeColorIndex =4
                            BackShade =50.0
                            BorderThemeColorIndex =6
                            BorderShade =75.0
                            ForeThemeColorIndex =1
                            ForeShade =85.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =17400
                    Top =420
                    Width =2220
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="ExitOptionsConsole"
                    Caption ="EXIT Options Dashboard"
                    OnClick ="[Event Procedure]"
                    GroupTable =4
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17400
                    LayoutCachedTop =420
                    LayoutCachedWidth =19620
                    LayoutCachedHeight =690
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =9211020
                    BorderColor =16777215
                    HoverColor =12566463
                    PressedColor =9211020
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =1
                    GroupTable =4
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =25920
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =6
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

                    LayoutCachedLeft =25920
                    LayoutCachedTop =120
                    LayoutCachedWidth =26475
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
                    OverlapFlags =215
                    PictureType =1
                    Left =26487
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =7
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1

                    LayoutCachedLeft =26487
                    LayoutCachedTop =120
                    LayoutCachedWidth =27042
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =13830
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Top =6660
                    Width =5295
                    Height =225
                    FontSize =9
                    FontWeight =700
                    Name ="Label58"
                    Caption ="┌────────┬──── Searches ──────┬───────────────────────────┐"
                    LayoutCachedTop =6660
                    LayoutCachedWidth =5295
                    LayoutCachedHeight =6885
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin ListBox
                    OverlapFlags =215
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =2946
                    Top =240
                    Width =2814
                    Height =1320
                    FontSize =9
                    TabIndex =3
                    Name ="LeaseList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MLID, MLNo, Client, MLOrig From MstrLease WHERE Client =4971 Order BY MLO"
                        "rig;"
                    ColumnWidths ="0;1800;0;936"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2946
                    LayoutCachedTop =240
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =1560
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                End
                Begin ListBox
                    OverlapFlags =215
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =2940
                    Top =1755
                    Width =2814
                    Height =1650
                    FontSize =9
                    TabIndex =5
                    Name ="ScheduleList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT * From EXITSchedule WHERE MLID=156 Order By SchID;"
                    ColumnWidths ="0;0;1800;936"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2940
                    LayoutCachedTop =1755
                    LayoutCachedWidth =5754
                    LayoutCachedHeight =3405
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =2946
                    Width =2814
                    Height =255
                    FontSize =9
                    FontWeight =700
                    Name ="Label22"
                    Caption ="MLA#                               Type"
                    LayoutCachedLeft =2946
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =255
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =2940
                    Top =1545
                    Width =2814
                    Height =225
                    FontSize =9
                    FontWeight =700
                    Name ="Label23"
                    Caption ="Schedules                         Units"
                    LayoutCachedLeft =2940
                    LayoutCachedTop =1545
                    LayoutCachedWidth =5754
                    LayoutCachedHeight =1770
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =93
                    Left =60
                    Width =2160
                    Height =255
                    FontSize =9
                    FontWeight =700
                    Name ="Label27"
                    Caption ="Clients"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =255
                    ForeThemeColorIndex =2
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =3
                    Left =240
                    Top =6600
                    Width =396
                    Height =225
                    FontSize =8
                    FontWeight =700
                    Name ="ClearLabel"
                    Caption ="Clear"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =240
                    LayoutCachedTop =6600
                    LayoutCachedWidth =636
                    LayoutCachedHeight =6825
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    ForeTint =60.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    TextFontCharSet =238
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =8
                    Top =7125
                    Width =5760
                    Height =6705
                    FontSize =8
                    FontWeight =700
                    TabIndex =22
                    BackColor =16250871
                    Name ="Units_List"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="720;288;1872;864;1037;2160;1008;1008"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Unit to see Unit's Description"
                    AllowValueListEdits =0

                    LayoutCachedTop =7125
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =13830
                    BorderThemeColorIndex =2
                End
                Begin TextBox
                    OverlapFlags =95
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4740
                    Top =6885
                    Width =990
                    FontSize =9
                    TabIndex =21
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchtDesc"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4740
                    LayoutCachedTop =6885
                    LayoutCachedWidth =5730
                    LayoutCachedHeight =7125
                    ColumnStart =5
                    ColumnEnd =5
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
                    Left =1005
                    Top =6885
                    Width =1872
                    FontSize =9
                    TabIndex =19
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchVin"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1005
                    LayoutCachedTop =6885
                    LayoutCachedWidth =2877
                    LayoutCachedHeight =7125
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1500
                    Top =5655
                    Width =1530
                    Height =840
                    FontSize =10
                    TabIndex =13
                    Name ="GetAll"
                    Caption =" Select All \015\012▼241 Units ▼"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1500
                    LayoutCachedTop =5655
                    LayoutCachedWidth =3030
                    LayoutCachedHeight =6495
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    HoverThemeColorIndex =4
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    PressedForeThemeColorIndex =0
                    QuickStyle =30
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Top =6885
                    Width =720
                    FontSize =9
                    TabIndex =17
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchAsset"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =6885
                    LayoutCachedWidth =720
                    LayoutCachedHeight =7125
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =720
                    Top =6885
                    Width =285
                    FontSize =9
                    TabIndex =18
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchStatus"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =720
                    LayoutCachedTop =6885
                    LayoutCachedWidth =1005
                    LayoutCachedHeight =7125
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =75
                    Top =465
                    Width =2664
                    Height =4639
                    FontSize =9
                    TabIndex =4
                    Name ="ClientList"
                    RowSourceType ="Table/Query"
                    RowSource ="ClientMasterLeasesTypeCount_Summed"
                    ColumnWidths ="0;1872;0;504;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =75
                    LayoutCachedTop =465
                    LayoutCachedWidth =2739
                    LayoutCachedHeight =5104
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =2940
                    Top =5145
                    Width =2814
                    Height =270
                    FontSize =9
                    TabIndex =7
                    BoundColumn =1
                    Name ="BankList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BankId, 'Assignee Bank: ' & Assignee as Bank, Sum(AssigneeBanks.Units) AS"
                        " Units, MLID, SchID From AssigneeBanks GROUP BY BankId, Assignee, MLID, SchID, S"
                        "chID HAVING AssigneeBanks.SchID =835;"
                    ColumnWidths ="0;1440;0;0;0;0"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2940
                    LayoutCachedTop =5145
                    LayoutCachedWidth =5754
                    LayoutCachedHeight =5415
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                End
                Begin Tab
                    OverlapFlags =85
                    Style =2
                    Left =5865
                    Top =120
                    Width =22230
                    Height =13200
                    FontSize =10
                    TabIndex =1
                    Name ="TabControl"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =5865
                    LayoutCachedTop =120
                    LayoutCachedWidth =28095
                    LayoutCachedHeight =13320
                    ThemeFontIndex =-1
                    OldBorderStyle =0
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =5940
                            Top =195
                            Width =22080
                            Height =13053
                            Name ="ExchangeITTab"
                            Caption ="ExchangeIT"
                            LayoutCachedLeft =5940
                            LayoutCachedTop =195
                            LayoutCachedWidth =28020
                            LayoutCachedHeight =13248
                            Begin
                                Begin Subform
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    OldBorderStyle =0
                                    Left =5971
                                    Top =864
                                    Width =21168
                                    Height =12384
                                    Name ="ExitOptions_subform_Exchange"
                                    SourceObject ="Form.ExitOptions_subform_Exchange"

                                    LayoutCachedLeft =5971
                                    LayoutCachedTop =864
                                    LayoutCachedWidth =27139
                                    LayoutCachedHeight =13248
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =6060
                                    Top =1800
                                    Width =12195
                                    Height =690
                                    FontSize =28
                                    ForeColor =-2147483609
                                    Name ="Label99"
                                    Caption ="◄Please Select: A Client, A Lease and Then a Schedule"
                                    LayoutCachedLeft =6060
                                    LayoutCachedTop =1800
                                    LayoutCachedWidth =18255
                                    LayoutCachedHeight =2490
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =6000
                                    Top =360
                                    Width =7320
                                    Height =492
                                    FontSize =18
                                    FontWeight =700
                                    ForeColor =8210719
                                    Name ="Label105"
                                    Caption ="ExchangeIT"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =6000
                                    LayoutCachedTop =360
                                    LayoutCachedWidth =13320
                                    LayoutCachedHeight =852
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5940
                            Top =195
                            Width =22080
                            Height =13053
                            Name ="ExtendITTab"
                            Caption ="ExtendIT"
                            LayoutCachedLeft =5940
                            LayoutCachedTop =195
                            LayoutCachedWidth =28020
                            LayoutCachedHeight =13248
                            Begin
                                Begin Subform
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5971
                                    Top =864
                                    Width =21168
                                    Height =12384
                                    Name ="ExitOptions_subform_Extend"
                                    SourceObject ="Form.ExitOptions_subform_Extend"
                                    HorizontalAnchor =2
                                    VerticalAnchor =2
                                    FilterOnEmptyMaster =0

                                    LayoutCachedLeft =5971
                                    LayoutCachedTop =864
                                    LayoutCachedWidth =27139
                                    LayoutCachedHeight =13248
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =5940
                                            Top =360
                                            Width =7320
                                            Height =492
                                            FontSize =18
                                            FontWeight =700
                                            Name ="Label107"
                                            Caption ="ExtendIT"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =5940
                                            LayoutCachedTop =360
                                            LayoutCachedWidth =13260
                                            LayoutCachedHeight =852
                                            ForeThemeColorIndex =0
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =6059
                                    Top =1800
                                    Width =12195
                                    Height =690
                                    FontSize =28
                                    ForeColor =-2147483609
                                    Name ="Label24"
                                    Caption ="◄Please Select: A Client, A Lease and Then a Schedule"
                                    LayoutCachedLeft =6059
                                    LayoutCachedTop =1800
                                    LayoutCachedWidth =18254
                                    LayoutCachedHeight =2490
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5940
                            Top =195
                            Width =22080
                            Height =13053
                            Name ="TerminateTab"
                            Caption ="Terminate"
                            LayoutCachedLeft =5940
                            LayoutCachedTop =195
                            LayoutCachedWidth =28020
                            LayoutCachedHeight =13248
                            Begin
                                Begin Subform
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5940
                                    Top =744
                                    Width =21150
                                    Height =12384
                                    Name ="ExitOptions_subform_Terminate"
                                    SourceObject ="Form.ExitOptions_subform_Terminate"

                                    LayoutCachedLeft =5940
                                    LayoutCachedTop =744
                                    LayoutCachedWidth =27090
                                    LayoutCachedHeight =13128
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =5968
                                            Top =240
                                            Width =7320
                                            Height =492
                                            FontSize =18
                                            FontWeight =700
                                            Name ="Label108"
                                            Caption ="Terminate Lease"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =5968
                                            LayoutCachedTop =240
                                            LayoutCachedWidth =13288
                                            LayoutCachedHeight =732
                                            ForeThemeColorIndex =5
                                            ForeShade =50.0
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =6060
                                    Top =1800
                                    Width =12195
                                    Height =690
                                    FontSize =28
                                    ForeColor =-2147483609
                                    Name ="Label104"
                                    Caption ="◄Please Select: A Client, A Lease and Then a Schedule"
                                    LayoutCachedLeft =6060
                                    LayoutCachedTop =1800
                                    LayoutCachedWidth =18255
                                    LayoutCachedHeight =2490
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5940
                            Top =195
                            Width =22080
                            Height =13053
                            Name ="HistoryTab"
                            LayoutCachedLeft =5940
                            LayoutCachedTop =195
                            LayoutCachedWidth =28020
                            LayoutCachedHeight =13248
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =5971
                                    Top =864
                                    Width =21168
                                    Height =12384
                                    Name ="ExitOptions_subform_History"
                                    SourceObject ="Form.ExitOptions_subform_History"

                                    LayoutCachedLeft =5971
                                    LayoutCachedTop =864
                                    LayoutCachedWidth =27139
                                    LayoutCachedHeight =13248
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =5970
                                    Top =360
                                    Width =7320
                                    Height =492
                                    FontSize =18
                                    FontWeight =700
                                    Name ="Label112"
                                    Caption ="History"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =5970
                                    LayoutCachedTop =360
                                    LayoutCachedWidth =13290
                                    LayoutCachedHeight =852
                                    ForeThemeColorIndex =0
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5940
                            Top =195
                            Width =22084
                            Height =13053
                            Name ="Get-ITTab"
                            EventProcPrefix ="Get_ITTab"
                            Caption ="Get-IT"
                            LayoutCachedLeft =5940
                            LayoutCachedTop =195
                            LayoutCachedWidth =28024
                            LayoutCachedHeight =13248
                            Begin
                                Begin Label
                                    OverlapFlags =215
                                    Left =5999
                                    Top =360
                                    Width =7320
                                    Height =492
                                    FontSize =18
                                    FontWeight =700
                                    Name ="Label120"
                                    Caption ="Buy Client's Portfolio Units"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =5999
                                    LayoutCachedTop =360
                                    LayoutCachedWidth =13319
                                    LayoutCachedHeight =852
                                    ForeThemeColorIndex =5
                                    ForeShade =50.0
                                End
                                Begin Subform
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5971
                                    Top =864
                                    Width =22053
                                    Height =12384
                                    Name ="ExitOptions_subform_BuyIT"
                                    SourceObject ="Form.ExitOptions_subform_BuyIT"

                                    LayoutCachedLeft =5971
                                    LayoutCachedTop =864
                                    LayoutCachedWidth =28024
                                    LayoutCachedHeight =13248
                                End
                            End
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =2940
                    Top =3645
                    Width =2814
                    FontSize =9
                    TabIndex =6
                    Name ="GroupList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT GroupId, UnitGroup, Count(AssetID) AS Units, SchID From ExitOptionsGroup "
                        "GROUP BY GroupId, UnitGroup, SchID HAVING SchID=835 Order BY UnitGroup; "
                    ColumnWidths ="0;1800;936;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2940
                    LayoutCachedTop =3645
                    LayoutCachedWidth =5754
                    LayoutCachedHeight =5085
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2940
                            Top =3465
                            Width =2814
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label75"
                            Caption ="Groups                              Units"
                            LayoutCachedLeft =2940
                            LayoutCachedTop =3465
                            LayoutCachedWidth =5754
                            LayoutCachedHeight =3705
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =3381
                    Top =5520
                    Width =2379
                    Height =315
                    FontSize =10
                    TabIndex =11
                    Name ="ExchangeSelected"
                    Caption ="ExchangeIT (1) Unit ►"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3381
                    LayoutCachedTop =5520
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =5835
                    Alignment =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8872517
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    PressedColor =8872517
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =30
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3360
                    Top =5910
                    Width =2379
                    Height =315
                    FontSize =10
                    TabIndex =14
                    Name ="ExtendSelected"
                    Caption ="ExtendIT (1) Unit ►"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3360
                    LayoutCachedTop =5910
                    LayoutCachedWidth =5739
                    LayoutCachedHeight =6225
                    Alignment =3
                    RowStart =1
                    RowEnd =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =0
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =5855577
                    PressedColor =0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3360
                    Top =6285
                    Width =2379
                    Height =315
                    FontSize =10
                    TabIndex =16
                    Name ="TerminateSelected"
                    Caption ="Terminate (1) Unit ►"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3360
                    LayoutCachedTop =6285
                    LayoutCachedWidth =5739
                    LayoutCachedHeight =6600
                    Alignment =3
                    RowStart =2
                    RowEnd =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =13209
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedColor =13209
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =35
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =75
                    Top =5655
                    Width =1245
                    Height =555
                    FontSize =10
                    TabIndex =12
                    ForeColor =0
                    Name ="History"
                    Caption ="Show Group's History"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =75
                    LayoutCachedTop =5655
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =6210
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    HoverThemeColorIndex =0
                    HoverTint =50.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    PressedForeColor =0
                    QuickStyle =30
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =6270
                    Width =1245
                    Height =240
                    FontSize =10
                    TabIndex =15
                    ForeColor =0
                    Name ="UnitHistory"
                    Caption ="Unit's History"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =60
                    LayoutCachedTop =6270
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =6510
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    HoverThemeColorIndex =0
                    HoverTint =50.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    PressedForeColor =0
                    QuickStyle =30
                    QuickStyleMask =-885
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =75
                    Top =225
                    Width =1872
                    FontSize =9
                    TabIndex =2
                    BackColor =13434879
                    BorderColor =-2147483609
                    ForeColor =-2147483615
                    Name ="SearchClient"
                    OnChange ="[Event Procedure]"
                    Tag ="MEdit"
                    ControlTipText ="Search Clients"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =75
                    LayoutCachedTop =225
                    LayoutCachedWidth =1947
                    LayoutCachedHeight =465
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3360
                    Top =5520
                    Width =2379
                    Height =315
                    FontSize =10
                    TabIndex =10
                    Name ="BuyITSelection"
                    Caption ="BuyIT (1) Unit ►"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3360
                    LayoutCachedTop =5520
                    LayoutCachedWidth =5739
                    LayoutCachedHeight =5835
                    Alignment =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8872517
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    PressedColor =8872517
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =30
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    ListWidth =720
                    Left =90
                    Top =5310
                    Width =720
                    Height =255
                    FontSize =9
                    TabIndex =8
                    BorderColor =8210719
                    Name ="EOLYear"
                    RowSourceType ="Value List"
                    RowSource ="2016;2017;2018;2019"
                    ColumnWidths ="720"
                    OnClick ="[Event Procedure]"
                    Tag ="SearchX"
                    ControlTipText ="Select a Client to Show its Units"
                    Format ="General Number"

                    LayoutCachedLeft =90
                    LayoutCachedTop =5310
                    LayoutCachedWidth =810
                    LayoutCachedHeight =5565
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =90
                            Top =5100
                            Width =2610
                            Height =240
                            FontSize =8
                            Name ="Label340"
                            Caption ="EOL Year ─────────────┐"
                            LayoutCachedLeft =90
                            LayoutCachedTop =5100
                            LayoutCachedWidth =2700
                            LayoutCachedHeight =5340
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2055
                    Width =720
                    Height =255
                    FontSize =8
                    Name ="ClientGroupID"
                    BottomPadding =45

                    LayoutCachedLeft =2055
                    LayoutCachedWidth =2775
                    LayoutCachedHeight =255
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =1170
                    Top =5310
                    Width =1418
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="GetLatestData"
                    Caption =" Sync|Get Data"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff00000000000000000000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffff727272ff00000000000000000000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffff727272ff00000000000000000000000000000000 ,
                        0x0000000000000000727272ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0x727272ff727272ff727272ff727272ff00000000727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0x727272ffffffffffffffffff727272ff00000000ffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0x727272ffffffffffffffffff727272ff00000000ffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ffb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4aff00000000b3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ff727272ffb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4aff00000000ffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ffb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4aff00000000ffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff000000000000000000000000000000000000000000000000 ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff000000000000000000000000000000000000000000000000 ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =1170
                    LayoutCachedTop =5310
                    LayoutCachedWidth =2588
                    LayoutCachedHeight =5550
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =2524406
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    HoverColor =8435191
                    PressedColor =2524406
                    HoverForeColor =1643706
                    QuickStyle =35
                    QuickStyleMask =-373
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2880
                    Top =6885
                    Width =1008
                    FontSize =9
                    TabIndex =20
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchUnitNum"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2880
                    LayoutCachedTop =6885
                    LayoutCachedWidth =3888
                    LayoutCachedHeight =7125
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =3885
                    Top =6885
                    Width =855
                    Name ="EmptyCell135"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3885
                    LayoutCachedTop =6885
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =7125
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =3
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
' See "EXITOptions.cls"
