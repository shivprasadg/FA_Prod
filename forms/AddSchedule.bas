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
    Width =6675
    DatasheetFontHeight =11
    ItemSuffix =25
    Left =4530
    Top =2715
    Right =13395
    Bottom =7230
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xd41d4fb412e8e440
    End
    Caption ="Add a New Schedule"
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
                    OverlapFlags =85
                    TextFontCharSet =238
                    Width =4395
                    Height =405
                    FontSize =16
                    FontWeight =700
                    ForeColor =10040879
                    Name ="Label6"
                    Caption ="Add New Schedule and Group 'A'"
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =405
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =5745
                    Top =30
                    Width =750
                    Height =228
                    FontSize =8
                    FontWeight =700
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    BackStyle =0

                    LayoutCachedLeft =5745
                    LayoutCachedTop =30
                    LayoutCachedWidth =6495
                    LayoutCachedHeight =258
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    HoverForeThemeColorIndex =4
                    HoverForeShade =75.0
                    PressedForeThemeColorIndex =4
                    PressedForeShade =75.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =5160
                    Top =30
                    Width =555
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="GoHome"
                    Caption ="Home"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End
                    BackStyle =0

                    LayoutCachedLeft =5160
                    LayoutCachedTop =30
                    LayoutCachedWidth =5715
                    LayoutCachedHeight =258
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    HoverColor =-2
                    HoverTint =80.0
                    PressedColor =-2
                    PressedShade =80.0
                    HoverForeThemeColorIndex =4
                    HoverForeShade =75.0
                    PressedForeThemeColorIndex =4
                    PressedForeShade =75.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =2280
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =5040
                    Left =1560
                    Top =75
                    Width =3960
                    Height =315
                    BackColor =13434879
                    Name ="SelecetMasterLease"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MstrLease.MLID, MstrLease.MLOrig, MstrLease.MLNo FROM MstrLease WHERE ((("
                        "MstrLease.MLOrig)=\"FA\") AND ((MstrLease.MLNo) Not Like \"*Demo*\") AND ((MstrL"
                        "ease.Client) Not Like \"*Demo*\")) ORDER BY MstrLease.MLOrig, MstrLease.MLNo;"
                    ColumnWidths ="0;0;2160"
                    GroupTable =1

                    LayoutCachedLeft =1560
                    LayoutCachedTop =75
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =390
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =30
                            Top =75
                            Width =1470
                            Height =315
                            Name ="Label1"
                            Caption ="Master Lease"
                            GroupTable =1
                            LayoutCachedLeft =30
                            LayoutCachedTop =75
                            LayoutCachedWidth =1500
                            LayoutCachedHeight =390
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =5700
                    Top =120
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =5
                    ForeColor =5026082
                    Name ="SourceSchID"

                    LayoutCachedLeft =5700
                    LayoutCachedTop =120
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1560
                    Top =450
                    Width =3960
                    Height =315
                    TabIndex =1
                    BackColor =13434879
                    Name ="SchNo"
                    GroupTable =1

                    LayoutCachedLeft =1560
                    LayoutCachedTop =450
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =765
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =30
                            Top =450
                            Width =1470
                            Height =315
                            Name ="Label4"
                            Caption ="Schedule Num"
                            GroupTable =1
                            LayoutCachedLeft =30
                            LayoutCachedTop =450
                            LayoutCachedWidth =1500
                            LayoutCachedHeight =765
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1560
                    Top =825
                    Width =3960
                    Height =315
                    ColumnWidth =4335
                    TabIndex =2
                    Name ="SchShortDesc"
                    GroupTable =1

                    LayoutCachedLeft =1560
                    LayoutCachedTop =825
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =1140
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =30
                            Top =825
                            Width =1470
                            Height =315
                            Name ="Label5"
                            Caption ="Schedule Desc"
                            GroupTable =1
                            LayoutCachedLeft =30
                            LayoutCachedTop =825
                            LayoutCachedWidth =1500
                            LayoutCachedHeight =1140
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1560
                    Top =1575
                    Width =3960
                    Height =600
                    TabIndex =4
                    Name ="Reason"
                    GroupTable =1

                    LayoutCachedLeft =1560
                    LayoutCachedTop =1575
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =2175
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =30
                            Top =1575
                            Width =1470
                            Height =600
                            Name ="Label12"
                            Caption ="Reason for Adding"
                            GroupTable =1
                            LayoutCachedLeft =30
                            LayoutCachedTop =1575
                            LayoutCachedWidth =1500
                            LayoutCachedHeight =2175
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5625
                    Top =1260
                    Width =1020
                    Height =300
                    FontSize =9
                    TabIndex =8
                    Name ="MakeSchedule"
                    Caption ="Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Make the Schedule defined in this Schedule Rule"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000ab598980ab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989800000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffd9b4caffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffd9b4caffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5625
                    LayoutCachedTop =1260
                    LayoutCachedWidth =6645
                    LayoutCachedHeight =1560
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =255
                    BackThemeColorIndex =3
                    BorderColor =16777215
                    HoverColor =616409
                    PressedColor =1643706
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-881
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5700
                    Top =450
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =6
                    ForeColor =5026082
                    Name ="MLID"

                    LayoutCachedLeft =5700
                    LayoutCachedTop =450
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =660
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5700
                    Top =780
                    Width =870
                    Height =210
                    FontSize =10
                    TabIndex =7
                    ForeColor =5026082
                    Name ="FAID"

                    LayoutCachedLeft =5700
                    LayoutCachedTop =780
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =990
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1560
                    Top =1200
                    Width =3960
                    Height =315
                    TabIndex =3
                    Name ="NumOfUnits"
                    GroupTable =1

                    LayoutCachedLeft =1560
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =1515
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =30
                            Top =1200
                            Width =1470
                            Height =315
                            Name ="Label20"
                            Caption ="Units In Group"
                            GroupTable =1
                            LayoutCachedLeft =30
                            LayoutCachedTop =1200
                            LayoutCachedWidth =1500
                            LayoutCachedHeight =1515
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
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
' See "AddSchedule.cls"
