Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =163
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8160
    DatasheetFontHeight =11
    ItemSuffix =64
    Left =12975
    Top =6930
    Right =21135
    Bottom =13830
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xf0eb4be7365fe540
    End
    RecordSource ="SELECT TOP 1 vw_SixKeys.ClientGroupID, vw_SixKeys.clientID, vw_SixKeys.SchID, vw"
        "_SixKeys.GroupID, vw_SixKeys.AssetsInGroup, vw_SixKeys.Schedule, vw_SixKeys.Unit"
        "Group FROM vw_SixKeys WHERE (((vw_SixKeys.GroupID)=2779));"
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
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
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
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =435
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =45
                    Width =5760
                    Height =435
                    ColumnOrder =0
                    FontSize =12
                    TabIndex =2
                    BorderColor =0
                    ForeColor =10040879
                    Name ="Label6"
                    ControlSource ="=\"Electronic Title Request\""

                    LayoutCachedLeft =45
                    LayoutCachedWidth =5805
                    LayoutCachedHeight =435
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =6720
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
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6720
                    LayoutCachedTop =60
                    LayoutCachedWidth =7275
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
                    Left =7291
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

                    LayoutCachedLeft =7291
                    LayoutCachedTop =60
                    LayoutCachedWidth =8041
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
            Height =6480
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =4020
                    Top =180
                    Width =1905
                    Height =480
                    FontSize =10
                    Name ="RunBatch"
                    Caption ="Run Title Request Batch"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Run The C of A as set"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =4020
                    LayoutCachedTop =180
                    LayoutCachedWidth =5925
                    LayoutCachedHeight =660
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =255
                    BackColor =9732411
                    BorderColor =16777215
                    HoverColor =616409
                    PressedColor =1643706
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6120
                    Top =180
                    Width =1905
                    Height =480
                    FontSize =10
                    TabIndex =1
                    Name ="RunHistory"
                    Caption ="Set Title Request History"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Run The C of A as set"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =6120
                    LayoutCachedTop =180
                    LayoutCachedWidth =8025
                    LayoutCachedHeight =660
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =255
                    BackThemeColorIndex =7
                    BackShade =75.0
                    BorderColor =16777215
                    HoverColor =616409
                    PressedColor =1643706
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-881
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2475
                    Top =540
                    Width =1215
                    Height =315
                    TabIndex =2
                    BorderColor =9211020
                    Name ="Commission"
                    Format ="Short Date"
                    RightPadding =75

                    LayoutCachedLeft =2475
                    LayoutCachedTop =540
                    LayoutCachedWidth =3690
                    LayoutCachedHeight =855
                    RowStart =19
                    RowEnd =19
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =150
                            Top =540
                            Width =2160
                            Height =375
                            FontSize =10
                            FontWeight =700
                            LeftMargin =45
                            TopMargin =45
                            RightMargin =45
                            BottomMargin =45
                            Name ="Label37"
                            Caption ="Commssion Ends:"
                            LeftPadding =60
                            RightPadding =75
                            LayoutCachedLeft =150
                            LayoutCachedTop =540
                            LayoutCachedWidth =2310
                            LayoutCachedHeight =915
                            RowStart =19
                            RowEnd =19
                            ColumnStart =1
                            ColumnEnd =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2475
                    Top =120
                    Width =1215
                    Height =315
                    TabIndex =3
                    BorderColor =9211020
                    Name ="Signed"
                    Format ="Short Date"
                    RightPadding =75

                    LayoutCachedLeft =2475
                    LayoutCachedTop =120
                    LayoutCachedWidth =3690
                    LayoutCachedHeight =435
                    RowStart =20
                    RowEnd =20
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =540
                            Top =120
                            Width =1770
                            Height =375
                            FontSize =10
                            FontWeight =700
                            LeftMargin =45
                            TopMargin =45
                            RightMargin =45
                            BottomMargin =45
                            Name ="Label39"
                            Caption ="Date Signed / Sent:"
                            LeftPadding =60
                            RightPadding =75
                            LayoutCachedLeft =540
                            LayoutCachedTop =120
                            LayoutCachedWidth =2310
                            LayoutCachedHeight =495
                            RowStart =20
                            RowEnd =20
                            ColumnStart =1
                            ColumnEnd =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =6
                    Left =120
                    Top =1260
                    Width =8040
                    Height =5100
                    TabIndex =4
                    Name ="BatchList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TOP 100 ElectronicTitleRequest.MYear, ElectronicTitleRequest.Make, Electr"
                        "onicTitleRequest.VIN, ElectronicTitleRequest.SchGrp, ElectronicTitleRequest.Last"
                        "sent, ElectronicTitleRequest.tAssetID AS AssetID FROM ElectronicTitleRequest WHE"
                        "RE (((ElectronicTitleRequest.Sentdate) Is Null)) ORDER BY ElectronicTitleRequest"
                        ".Lastsent, ElectronicTitleRequest.VIN;"
                    ColumnWidths ="720;1440;2448;1728;1152;1152"

                    LayoutCachedLeft =120
                    LayoutCachedTop =1260
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =6360
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =360
                            Top =960
                            Width =1770
                            Height =315
                            Name ="Label63"
                            Caption ="Batch List of Units:"
                            LayoutCachedLeft =360
                            LayoutCachedTop =960
                            LayoutCachedWidth =2130
                            LayoutCachedHeight =1275
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =4020
                    Top =840
                    Width =3967
                    Height =315
                    TabIndex =5
                    Name ="ExcelExport"
                    Caption =" Excel Batch Export"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
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

                    LayoutCachedLeft =4020
                    LayoutCachedTop =840
                    LayoutCachedWidth =7987
                    LayoutCachedHeight =1155
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
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
' See "ElectronciTitlePopUp.cls"
