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
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7020
    DatasheetFontHeight =11
    ItemSuffix =71
    Left =11085
    Top =6660
    Right =19290
    Bottom =9675
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x135f4f982d0fe540
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            BackColor =15527148
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =3120
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =50
                    IMESentenceMode =3
                    Left =1095
                    Top =60
                    Width =3195
                    Height =315
                    FontSize =12
                    Name ="UnitVIN"
                    FontName ="OCR A Extended"
                    BottomPadding =150

                    LayoutCachedLeft =1095
                    LayoutCachedTop =60
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =375
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =165
                            Top =60
                            Width =900
                            Height =315
                            Name ="Label1"
                            Caption ="VIN "
                            BottomPadding =150
                            LayoutCachedLeft =165
                            LayoutCachedTop =60
                            LayoutCachedWidth =1065
                            LayoutCachedHeight =375
                            ColumnStart =1
                            ColumnEnd =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5655
                    Top =60
                    Width =1155
                    Height =315
                    FontSize =12
                    TabIndex =1
                    Name ="UnitYr"
                    BottomPadding =150

                    LayoutCachedLeft =5655
                    LayoutCachedTop =60
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =375
                    ColumnStart =4
                    ColumnEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4470
                            Top =60
                            Width =1125
                            Height =315
                            Name ="Label5"
                            Caption ="Model Year"
                            BottomPadding =150
                            LayoutCachedLeft =4470
                            LayoutCachedTop =60
                            LayoutCachedWidth =5595
                            LayoutCachedHeight =375
                            ColumnStart =3
                            ColumnEnd =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3735
                    Top =1770
                    Width =1320
                    Height =390
                    TabIndex =5
                    Name ="LogTitleIn"
                    Caption ="Log Title IN"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Log This Title IN"
                    LeftPadding =120
                    TopPadding =105
                    RightPadding =135
                    BottomPadding =120

                    LayoutCachedLeft =3735
                    LayoutCachedTop =1770
                    LayoutCachedWidth =5055
                    LayoutCachedHeight =2160
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderColor =5855577
                    HoverThemeColorIndex =4
                    HoverShade =75.0
                    PressedColor =1643706
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1095
                    Top =490
                    Width =3195
                    Height =315
                    FontSize =12
                    TabIndex =2
                    Name ="UnitTitleNumber"
                    BottomPadding =150

                    LayoutCachedLeft =1095
                    LayoutCachedTop =490
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =805
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =165
                            Top =490
                            Width =900
                            Height =315
                            Name ="Label12"
                            Caption ="Title #  "
                            BottomPadding =150
                            LayoutCachedLeft =165
                            LayoutCachedTop =490
                            LayoutCachedWidth =1065
                            LayoutCachedHeight =805
                            RowStart =1
                            RowEnd =1
                            ColumnEnd =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1095
                    Top =1365
                    Width =5730
                    Height =330
                    TabIndex =4
                    BackColor =16515069
                    BorderColor =12632256
                    ForeColor =0
                    Name ="EntryReason"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_TitleReasons.ReasonDescription, LookUp_TitleReasons.ReasonType, Lo"
                        "okUp_TitleReasons.ReasonId FROM LookUp_TitleReasons ORDER BY LookUp_TitleReasons"
                        ".ReasonDescription;"
                    ColumnWidths ="5616;0;0"
                    DefaultValue ="\"Received from State\""
                    OnClick ="[Event Procedure]"
                    BottomPadding =150

                    LayoutCachedLeft =1095
                    LayoutCachedTop =1365
                    LayoutCachedWidth =6825
                    LayoutCachedHeight =1695
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =165
                            Top =1365
                            Width =900
                            Height =330
                            Name ="Label4"
                            Caption ="Reason "
                            BottomPadding =150
                            LayoutCachedLeft =165
                            LayoutCachedTop =1365
                            LayoutCachedWidth =1065
                            LayoutCachedHeight =1695
                            RowStart =2
                            RowEnd =2
                            ColumnEnd =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5655
                    Top =945
                    Width =1155
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =6
                    Name ="UnitID"
                    BottomPadding =150

                    LayoutCachedLeft =5655
                    LayoutCachedTop =945
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =1260
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4470
                            Top =945
                            Width =1125
                            Height =315
                            Name ="Label60"
                            Caption ="AssetID"
                            BottomPadding =150
                            LayoutCachedLeft =4470
                            LayoutCachedTop =945
                            LayoutCachedWidth =5595
                            LayoutCachedHeight =1260
                            RowStart =1
                            RowEnd =1
                            ColumnStart =3
                            ColumnEnd =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5475
                    Top =1785
                    Width =1320
                    Height =390
                    TabIndex =7
                    Name ="LogTitleOut"
                    Caption ="Log Title OUT"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Log This Title OUT"
                    LeftPadding =120
                    TopPadding =105
                    RightPadding =135
                    BottomPadding =120

                    LayoutCachedLeft =5475
                    LayoutCachedTop =1785
                    LayoutCachedWidth =6795
                    LayoutCachedHeight =2175
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =0
                    BorderColor =5855577
                    HoverColor =13017476
                    PressedColor =1643706
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5655
                    Top =490
                    Width =1155
                    Height =330
                    FontSize =12
                    TabIndex =3
                    BorderColor =12632256
                    ForeColor =0
                    Name ="TitledState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_States.StateAbbrev FROM LookUp_States ORDER BY LookUp_States.State"
                        "Abbrev;"
                    BottomPadding =150

                    LayoutCachedLeft =5655
                    LayoutCachedTop =490
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =820
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4470
                            Top =490
                            Width =1125
                            Height =315
                            Name ="Label2"
                            Caption ="Title State"
                            BottomPadding =150
                            LayoutCachedLeft =4470
                            LayoutCachedTop =490
                            LayoutCachedWidth =5595
                            LayoutCachedHeight =805
                            RowStart =1
                            RowEnd =1
                            ColumnStart =3
                            ColumnEnd =3
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1095
                    Top =935
                    Width =1785
                    Height =315
                    FontSize =12
                    TabIndex =8
                    Name ="UnitTitleLogDate"
                    Format ="Short Date"
                    ControlTipText ="You may change this to the date needed"
                    BottomPadding =150

                    LayoutCachedLeft =1095
                    LayoutCachedTop =935
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =1250
                    ColumnStart =4
                    ColumnEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =165
                            Top =935
                            Width =900
                            Height =315
                            Name ="Label59"
                            Caption ="Log Date "
                            BottomPadding =150
                            LayoutCachedLeft =165
                            LayoutCachedTop =935
                            LayoutCachedWidth =1065
                            LayoutCachedHeight =1250
                            ColumnStart =3
                            ColumnEnd =3
                            BackThemeColorIndex =1
                            BackShade =85.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =2700
                    Width =6780
                    Height =360
                    FontSize =12
                    FontWeight =700
                    TabIndex =9
                    Name ="UnitDesc"
                    FontName ="Segoe UI"
                    BottomPadding =150

                    LayoutCachedLeft =60
                    LayoutCachedTop =2700
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =3060
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    ForeThemeColorIndex =2
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =2340
                    Width =6780
                    Height =360
                    FontSize =12
                    FontWeight =700
                    TabIndex =10
                    Name ="SchGrp"
                    FontName ="Segoe UI"
                    BottomPadding =150

                    LayoutCachedLeft =60
                    LayoutCachedTop =2340
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =2700
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =480
                    Top =1860
                    Width =750
                    Height =180
                    FontSize =8
                    TabIndex =11
                    Name ="EntryReasonId"
                    BottomPadding =150

                    LayoutCachedLeft =480
                    LayoutCachedTop =1860
                    LayoutCachedWidth =1230
                    LayoutCachedHeight =2040
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "TitleCheckInSub.cls"
