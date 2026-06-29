Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4740
    DatasheetFontHeight =11
    ItemSuffix =277
    Left =570
    Top =11040
    Right =5085
    Bottom =17505
    RecSrcDt = Begin
        0xe43512f2b6cae540
    End
    RecordSource ="SELECT Schedule.SchNo, SchGrp.SGrpNO, EO.UnitID, EO.UserName, EO.DatePosted, EO."
        "Notes, XD.ReplacedByUnitID, EO.ReplacedByScheduleID, EO.ReplacedByGroupID, XD.St"
        "artOrPurposedDate, XD.EndOrActualDate, XD.FinalRentDate, XD.oGroupID, XD.nGroupI"
        "D, XD.oBLCD, XD.BDay, XD.exStartDate, XD.exFinalRent, XD.exActual, XD.oProjEOL, "
        "XD.nEstDelv, XD.nActualDelv, XD.nAccepted, XD.nPDStart, XD.nInServ, XD.nOLDate F"
        "ROM ((UnitExitOption AS EO LEFT JOIN vw_EXchangeITDates AS XD ON EO.UnitID = XD."
        "UnitID) LEFT JOIN SchGrp ON EO.ReplacedByGroupID = SchGrp.SGrpID) LEFT JOIN Sche"
        "dule ON EO.ReplacedByScheduleID = Schedule.SchID WHERE (((EO.UnitID)=0) AND ((EO"
        ".OptionTypeId)=1));"
    OnCurrent ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =405
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =1
                    PictureAlignment =0
                    PictureType =2
                    Top =240
                    Width =4740
                    Height =165
                    Name ="Image324"
                    Picture ="1_WhiteShadow"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =240
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =405
                    TabIndex =1
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =1
                    Width =4740
                    Height =270
                    FontSize =8
                    BorderColor =10040879
                    Name ="Label325"
                    Caption ="ExchangeIT "
                    FontName ="Segoe UI"
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =270
                    ThemeFontIndex =-1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =3480
                    Top =30
                    Width =1170
                    Height =210
                    FontSize =8
                    FontWeight =700
                    Name ="ModifyMe"
                    Caption ="Add Option"
                    OnClick ="[Event Procedure]"
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3480
                    LayoutCachedTop =30
                    LayoutCachedWidth =4650
                    LayoutCachedHeight =240
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =9592887
                    BackTint =100.0
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =5069
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =1140
                    Top =2406
                    Height =233
                    Name ="Box275"
                    LayoutCachedLeft =1140
                    LayoutCachedTop =2406
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =2639
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =95
                    Left =1140
                    Top =2647
                    Height =292
                    Name ="Box274"
                    LayoutCachedLeft =1140
                    LayoutCachedTop =2647
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =2939
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =223
                    Left =1800
                    Top =2395
                    Width =120
                    Height =648
                    Name ="Box276"
                    LayoutCachedLeft =1800
                    LayoutCachedTop =2395
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =3043
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =1140
                    Top =840
                    Height =225
                    Name ="Box272"
                    LayoutCachedLeft =1140
                    LayoutCachedTop =840
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =1065
                    BorderThemeColorIndex =2
                    BorderTint =60.0
                    BorderShade =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =1140
                    Top =1365
                    Height =300
                    Name ="Box271"
                    LayoutCachedLeft =1140
                    LayoutCachedTop =1365
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =1665
                    BorderThemeColorIndex =2
                    BorderTint =60.0
                    BorderShade =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =1140
                    Top =1065
                    Height =922
                    Name ="Box270"
                    LayoutCachedLeft =1140
                    LayoutCachedTop =1065
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =1987
                    BorderThemeColorIndex =2
                    BorderTint =60.0
                    BorderShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1560
                    Top =4800
                    Width =165
                    FontSize =10
                    TabIndex =12
                    BorderColor =12632256
                    ForeColor =0
                    Name ="UnitID"
                    ControlSource ="UnitID"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =4800
                    LayoutCachedWidth =1725
                    LayoutCachedHeight =5040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =30
                    Top =4799
                    Width =1455
                    Height =270
                    FontSize =8
                    TabIndex =10
                    BorderColor =12632256
                    ForeColor =8355711
                    Name ="UserName"
                    ControlSource ="UserName"

                    LayoutCachedLeft =30
                    LayoutCachedTop =4799
                    LayoutCachedWidth =1485
                    LayoutCachedHeight =5069
                    RowStart =8
                    RowEnd =8
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1545
                    Top =4799
                    Width =2880
                    Height =270
                    FontSize =8
                    TabIndex =11
                    BorderColor =12632256
                    ForeColor =8355711
                    Name ="DatePosted"
                    ControlSource ="DatePosted"

                    LayoutCachedLeft =1545
                    LayoutCachedTop =4799
                    LayoutCachedWidth =4425
                    LayoutCachedHeight =5069
                    RowStart =8
                    RowEnd =8
                    ColumnStart =2
                    ColumnEnd =7
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =3240
                    Width =4395
                    Height =1500
                    FontSize =9
                    TabIndex =9
                    BackColor =16252927
                    BorderColor =12632256
                    ForeColor =0
                    Name ="Notes"
                    ControlSource ="Notes"

                    LayoutCachedLeft =60
                    LayoutCachedTop =3240
                    LayoutCachedWidth =4455
                    LayoutCachedHeight =4740
                    RowStart =7
                    RowEnd =7
                    ColumnEnd =7
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =3
                    Top =2220
                    Width =2850
                    Height =270
                    FontSize =8
                    FontWeight =700
                    BorderColor =0
                    Name ="Label9"
                    Caption ="This Unit's Final Dates ▼"
                    FontName ="Segoe UI"
                    LayoutCachedTop =2220
                    LayoutCachedWidth =2850
                    LayoutCachedHeight =2490
                    ColumnEnd =2
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Width =4140
                    Height =270
                    FontSize =8
                    FontWeight =700
                    BorderColor =0
                    Name ="Label20"
                    Caption ="EXchangedIT with  Schedule            |             Group"
                    FontName ="Segoe UI"
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =270
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2701
                    Top =330
                    Width =135
                    Height =255
                    FontSize =8
                    FontWeight =700
                    BorderColor =0
                    ForeColor =0
                    Name ="Label4"
                    Caption ="|"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =2701
                    LayoutCachedTop =330
                    LayoutCachedWidth =2836
                    LayoutCachedHeight =585
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1785
                    Top =4800
                    Width =165
                    FontSize =10
                    TabIndex =13
                    BorderColor =12632256
                    ForeColor =0
                    Name ="ReplacedByGroupID"
                    ControlSource ="ReplacedByGroupID"

                    LayoutCachedLeft =1785
                    LayoutCachedTop =4800
                    LayoutCachedWidth =1950
                    LayoutCachedHeight =5040
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =2025
                    Top =4800
                    Width =165
                    FontSize =10
                    TabIndex =14
                    BorderColor =12632256
                    ForeColor =0
                    Name ="ReplacedByScheduleID"
                    ControlSource ="ReplacedByScheduleID"

                    LayoutCachedLeft =2025
                    LayoutCachedTop =4800
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =5040
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Top =330
                    Width =2610
                    Height =285
                    FontSize =10
                    BorderColor =12632256
                    ForeColor =0
                    Name ="ScheduleName"
                    ControlSource ="SchNo"

                    LayoutCachedTop =330
                    LayoutCachedWidth =2610
                    LayoutCachedHeight =615
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =4
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2970
                    Top =330
                    Width =1425
                    Height =285
                    FontSize =10
                    TabIndex =1
                    BorderColor =12632256
                    ForeColor =0
                    Name ="GroupName"
                    ControlSource ="SGrpNO"

                    LayoutCachedLeft =2970
                    LayoutCachedTop =330
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =615
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =7
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2928
                    Top =2520
                    Width =1236
                    Height =255
                    FontSize =10
                    TabIndex =7
                    BorderColor =12632256
                    ForeColor =0
                    Name ="FinalRentDate"
                    ControlSource ="FinalRentDate"
                    Format ="Short Date"
                    GroupTable =1

                    LayoutCachedLeft =2928
                    LayoutCachedTop =2520
                    LayoutCachedWidth =4164
                    LayoutCachedHeight =2775
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =8
                    BackTint =40.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =1
                            Left =1536
                            Top =2520
                            Width =1320
                            Height =255
                            FontSize =8
                            BorderColor =0
                            Name ="Label212"
                            Caption ="Final Rent Date"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =1536
                            LayoutCachedTop =2520
                            LayoutCachedWidth =2856
                            LayoutCachedHeight =2775
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2928
                    Top =2844
                    Width =1236
                    Height =252
                    FontSize =10
                    TabIndex =8
                    BackColor =14151142
                    BorderColor =12632256
                    ForeColor =0
                    Name ="cProJOffLease"
                    Format ="Short Date"
                    GroupTable =1

                    LayoutCachedLeft =2928
                    LayoutCachedTop =2844
                    LayoutCachedWidth =4164
                    LayoutCachedHeight =3096
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =1
                            Left =1536
                            Top =2844
                            Width =1320
                            Height =252
                            FontSize =8
                            BorderColor =0
                            Name ="Label214"
                            Caption ="Proj Off-lease "
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =1536
                            LayoutCachedTop =2844
                            LayoutCachedWidth =2856
                            LayoutCachedHeight =3096
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2970
                    Top =660
                    Width =1230
                    FontSize =10
                    TabIndex =2
                    BorderColor =12632256
                    Name ="ReplacedByUnitID"
                    ControlSource ="ReplacedByUnitID"

                    LayoutCachedLeft =2970
                    LayoutCachedTop =660
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =900
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =7
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextAlign =3
                            Top =660
                            Width =2850
                            Height =240
                            FontSize =8
                            FontWeight =700
                            BorderColor =0
                            Name ="Label5"
                            Caption ="EXchangedIT with AssetID ♦"
                            FontName ="Segoe UI"
                            LayoutCachedTop =660
                            LayoutCachedWidth =2850
                            LayoutCachedHeight =900
                            RowStart =4
                            RowEnd =4
                            ColumnStart =4
                            ColumnEnd =5
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =223
                    Left =1800
                    Top =907
                    Width =120
                    Height =1008
                    Name ="Box273"
                    LayoutCachedLeft =1800
                    LayoutCachedTop =907
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =1915
                End
                Begin TextBox
                    SpecialEffect =5
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2940
                    Top =960
                    Width =1230
                    Height =255
                    FontSize =10
                    TabIndex =3
                    BorderColor =12632256
                    ForeColor =0
                    Name ="nEstDelv"
                    ControlSource ="nEstDelv"
                    GroupTable =2

                    LayoutCachedLeft =2940
                    LayoutCachedTop =960
                    LayoutCachedWidth =4170
                    LayoutCachedHeight =1215
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =1
                            Left =1560
                            Top =960
                            Width =1320
                            Height =255
                            FontSize =8
                            BorderColor =0
                            Name ="Label1"
                            Caption ="Est.Delivery"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =1560
                            LayoutCachedTop =960
                            LayoutCachedWidth =2880
                            LayoutCachedHeight =1215
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =5
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2940
                    Top =1272
                    Width =1230
                    Height =252
                    FontSize =10
                    TabIndex =4
                    BorderColor =12632256
                    ForeColor =0
                    Name ="nActualDelv"
                    ControlSource ="nActualDelv"
                    Format ="Short Date"
                    GroupTable =2

                    LayoutCachedLeft =2940
                    LayoutCachedTop =1272
                    LayoutCachedWidth =4170
                    LayoutCachedHeight =1524
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =1
                            Left =1560
                            Top =1272
                            Width =1320
                            Height =252
                            FontSize =8
                            BorderColor =0
                            Name ="Label2"
                            Caption ="Actual Delivery"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =1560
                            LayoutCachedTop =1272
                            LayoutCachedWidth =2880
                            LayoutCachedHeight =1524
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2940
                    Top =1584
                    Width =1230
                    Height =252
                    FontSize =10
                    TabIndex =5
                    BorderColor =12632256
                    ForeColor =0
                    Name ="nAccepted"
                    ControlSource ="nAccepted"
                    Format ="Short Date"
                    GroupTable =2

                    LayoutCachedLeft =2940
                    LayoutCachedTop =1584
                    LayoutCachedWidth =4170
                    LayoutCachedHeight =1836
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =1
                            Left =1560
                            Top =1584
                            Width =1320
                            Height =252
                            FontSize =8
                            BorderColor =0
                            Name ="Label259"
                            Caption ="Accepted"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =1560
                            LayoutCachedTop =1584
                            LayoutCachedWidth =2880
                            LayoutCachedHeight =1836
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =5
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2940
                    Top =1896
                    Width =1230
                    Height =252
                    FontSize =10
                    TabIndex =6
                    BorderColor =12632256
                    ForeColor =0
                    Name ="nInServ"
                    ControlSource ="nInServ"
                    Format ="Short Date"
                    GroupTable =2

                    LayoutCachedLeft =2940
                    LayoutCachedTop =1896
                    LayoutCachedWidth =4170
                    LayoutCachedHeight =2148
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =1
                            Left =1560
                            Top =1896
                            Width =1320
                            Height =252
                            FontSize =8
                            BorderColor =0
                            Name ="Label265"
                            Caption ="In Serviced"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =1560
                            LayoutCachedTop =1896
                            LayoutCachedWidth =2880
                            LayoutCachedHeight =2148
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =2
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
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
' See "Units_Exit_ExchangeIT.cls"
