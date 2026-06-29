Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =19275
    DatasheetFontHeight =11
    ItemSuffix =5
    Left =5175
    Top =11370
    Right =18765
    Bottom =19260
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x835b2967d6b9e440
    End
    RecordSource ="SELECT MstrLease.MLID, MstrLease.Client, MstrLease.MLNo, MstrLease.MLOrig FROM M"
        "strLease WHERE (((MstrLease.Client)=[forms]![ClientsView]![clientid])) ORDER BY "
        "MstrLease.MLNo, MstrLease.MLOrig;"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =7350
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =30
                    Top =570
                    Width =2520
                    Height =4995
                    FontSize =9
                    BackColor =-2147483643
                    Name ="MasterLeaseList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MstrLease.MLID, MstrLease.Client, MstrLease.MLNo, MstrLease.MLOrig FROM M"
                        "strLease WHERE (((MstrLease.Client)=[forms]![ClientsView]![clientid])) ORDER BY "
                        "MstrLease.MLOrig, MstrLease.MLNo;"
                    ColumnWidths ="0;0;1728;720"
                    OnClick ="[Event Procedure]"
                    GridlineColor =0
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =30
                    LayoutCachedTop =570
                    LayoutCachedWidth =2550
                    LayoutCachedHeight =5565
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =75.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17745
                    Top =60
                    Height =195
                    FontSize =8
                    TabIndex =1
                    Name ="MLID"
                    ControlSource ="MLID"

                    LayoutCachedLeft =17745
                    LayoutCachedTop =60
                    LayoutCachedWidth =19185
                    LayoutCachedHeight =255
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =2580
                    Top =660
                    Width =16560
                    Height =6690
                    TabIndex =2
                    Name ="ClientsView_Schedules2"
                    SourceObject ="Form.ClientsView_Schedules2"
                    LinkChildFields ="MLID"
                    LinkMasterFields ="MLID"

                    LayoutCachedLeft =2580
                    LayoutCachedTop =660
                    LayoutCachedWidth =19140
                    LayoutCachedHeight =7350
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =30
                    Top =45
                    Width =6675
                    Height =525
                    FontSize =26
                    FontWeight =700
                    BackColor =-2147483633
                    BorderColor =0
                    ForeColor =-2147483600
                    Name ="ScheduleLbl"
                    Caption ="Select a Schedule"
                    GridlineColor =0
                    LayoutCachedLeft =30
                    LayoutCachedTop =45
                    LayoutCachedWidth =6705
                    LayoutCachedHeight =570
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =5640
                    Width =2460
                    Height =288
                    FontSize =9
                    TabIndex =3
                    Name ="AddASchedule"
                    Caption ="Add A NEW Schedule"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =60
                    LayoutCachedTop =5640
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =5928
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =5
                    BackTint =100.0
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =9
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedThemeColorIndex =5
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =93
                    Left =2640
                    Top =870
                    Width =6795
                    Height =645
                    FontSize =26
                    FontWeight =700
                    BackColor =-2147483633
                    BorderColor =0
                    ForeColor =-2147483600
                    Name ="Label3"
                    Caption ="◄ Select Master Lease"
                    GridlineColor =0
                    LayoutCachedLeft =2640
                    LayoutCachedTop =870
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =1515
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =6030
                    Width =2460
                    Height =288
                    FontSize =9
                    TabIndex =4
                    Name ="ScheduleSortForm"
                    Caption ="Schedule Sort"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =60
                    LayoutCachedTop =6030
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =6318
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =12419407
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =8872517
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =12419407
                    PressedShade =100.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =75
                    Top =6480
                    Width =2430
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =0
                    Name ="LPAF"
                    Caption ="Unit Summary"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =75
                    LayoutCachedTop =6480
                    LayoutCachedWidth =2505
                    LayoutCachedHeight =6735
                    PictureCaptionArrangement =5
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =1
                    BackTint =100.0
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    HoverThemeColorIndex =1
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
' See "ClientsView_Schedules1.cls"
