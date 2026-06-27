Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =166
    Left =4095
    Top =2460
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x5fe1e8c312a6e540
    End
    RecordSource ="SELECT * FROM TFRCompare_Summary WHERE (((TFRCompare_Summary.[Assigned_Indicator"
        "])<>\"H\"));"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
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
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="FundStatus"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="FundStatus"
        End
        Begin BreakLevel
            ControlSource ="AccountName"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =525
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Top =75
                    Width =11160
                    Height =360
                    FontSize =14
                    FontWeight =700
                    BorderColor =16777215
                    ForeColor =4210752
                    Name ="Text133"
                    ControlSource ="=\"Summary Report \" & [OpenArgs]"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedTop =75
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =435
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7620
                    Width =2310
                    Height =315
                    TabIndex =1
                    ForeColor =8355711
                    Name ="Label135"
                    ControlSource ="=\"As of: \" & Date()"
                    GridlineColor =10921638

                    LayoutCachedLeft =7620
                    LayoutCachedWidth =9930
                    LayoutCachedHeight =315
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =840
            BreakLevel =1
            Name ="GroupHeader0"
            AlternateBackColor =14151142
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Left =255
                    Top =390
                    Width =1485
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label1"
                    Caption ="Account "
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =255
                    LayoutCachedTop =390
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =840
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =1
                    Left =1740
                    Top =390
                    Width =540
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label2"
                    Caption ="Units"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =1740
                    LayoutCachedTop =390
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =840
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =2280
                    Top =390
                    Width =1485
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label3"
                    Caption ="Total \015\012OEC"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2280
                    LayoutCachedTop =390
                    LayoutCachedWidth =3765
                    LayoutCachedHeight =840
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =6750
                    Top =390
                    Width =600
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label53"
                    Caption ="PD \015\012% "
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =6750
                    LayoutCachedTop =390
                    LayoutCachedWidth =7350
                    LayoutCachedHeight =840
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =7350
                    Top =390
                    Width =600
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label57"
                    Caption ="UTA \015\012% "
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =7350
                    LayoutCachedTop =390
                    LayoutCachedWidth =7950
                    LayoutCachedHeight =840
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =7950
                    Top =390
                    Width =855
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label64"
                    Caption ="Total Fee % "
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =7950
                    LayoutCachedTop =390
                    LayoutCachedWidth =8805
                    LayoutCachedHeight =840
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =10140
                    Top =390
                    Width =1335
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label71"
                    Caption ="Total \015\012Fee $"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =10140
                    LayoutCachedTop =390
                    LayoutCachedWidth =11475
                    LayoutCachedHeight =840
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    IMESentenceMode =3
                    Width =11520
                    Height =360
                    FontSize =12
                    BorderColor =16777215
                    ForeColor =4210752
                    Name ="Text112"
                    ControlSource ="=\"  <b>\" & IIf([FundStatus]=\"Syndicated\",\"Active \",\"TMT \") & [FundStatus"
                        "] & \"</b> Schedules and Groups\""
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedWidth =11520
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3765
                    Top =390
                    Width =1335
                    Height =450
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="Label4"
                    ControlSource ="=IIf([FundStatus]=\"Syndicated\",\"Bank\",\"Target Bank\")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =3765
                    LayoutCachedTop =390
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =840
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5925
                    Top =390
                    Width =825
                    Height =450
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="Label38"
                    ControlSource ="=IIf([FundStatus]=\"Syndicated\",\"Net Synd Fee\",\"Target Fee\")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =5925
                    LayoutCachedTop =390
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =840
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5100
                    Top =390
                    Width =825
                    Height =450
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="Text150"
                    ControlSource ="=IIf([FundStatus]=\"Syndicated\",\"Gross Fee\",\"Target Fee\")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =5100
                    LayoutCachedTop =390
                    LayoutCachedWidth =5925
                    LayoutCachedHeight =840
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =8805
                    Top =390
                    Width =1335
                    Height =450
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label158"
                    Caption ="Gross \015\012Fee $"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =8805
                    LayoutCachedTop =390
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =840
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeTint =100.0
                    GroupTable =2
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            Name ="Detail"
            AlternateBackColor =16777215
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1740
                    Width =540
                    Height =255
                    ColumnWidth =1650
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="UnitCount"
                    ControlSource ="UnitCount"
                    Format ="Standard"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1740
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2280
                    Width =1485
                    Height =255
                    ColumnWidth =1620
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="OEC"
                    ControlSource ="OEC"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =2280
                    LayoutCachedWidth =3765
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =2
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =255
                    Width =1485
                    Height =255
                    ColumnWidth =1590
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="AccountName"
                    ControlSource ="AccountName"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =255
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3765
                    Width =1335
                    Height =255
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TargetBank"
                    ControlSource ="TargetBank"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =3765
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5925
                    Width =825
                    Height =255
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="NetSyndFeePct"
                    ControlSource ="=[NetSyndFee]/[OEC]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =5925
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Width =600
                    Height =255
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text55"
                    ControlSource ="=[PDRent]/[OEC]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =6750
                    LayoutCachedWidth =7350
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7350
                    Width =600
                    Height =255
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text58"
                    ControlSource ="=[UTA]/[OEC]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =7350
                    LayoutCachedWidth =7950
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7950
                    Width =855
                    Height =255
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text65"
                    ControlSource ="=[TotalFee]/[OEC]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =7950
                    LayoutCachedWidth =8805
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10140
                    Width =1335
                    Height =255
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text72"
                    ControlSource ="TotalFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =10140
                    LayoutCachedWidth =11475
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =2
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5100
                    Width =825
                    Height =255
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text136"
                    ControlSource ="=[SyndFeePlus]/[OEC]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =5100
                    LayoutCachedWidth =5925
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8805
                    Width =1335
                    Height =255
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text152"
                    ControlSource ="=[SyndFeePlus]"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =8805
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =2
                    CurrencySymbol ="$"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =330
            BreakLevel =1
            Name ="GroupFooter0"
            AlternateBackColor =14151142
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Top =50
                    Width =1815
                    Height =270
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text130"
                    ControlSource ="=[FundStatus] & \" Sub Totals →\""
                    GridlineColor =10921638

                    LayoutCachedTop =50
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =320
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1740
                    Top =15
                    Width =540
                    Height =255
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text78"
                    ControlSource ="=Sum([UnitCount])"
                    Format ="Standard"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1740
                    LayoutCachedTop =15
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =270
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2280
                    Top =15
                    Width =1485
                    Height =255
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text79"
                    ControlSource ="=Sum([OEC])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =2280
                    LayoutCachedTop =15
                    LayoutCachedWidth =3765
                    LayoutCachedHeight =270
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3765
                    Top =15
                    Width =1335
                    Height =255
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text80"
                    ControlSource ="=\"\""
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =3765
                    LayoutCachedTop =15
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =270
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5925
                    Top =15
                    Width =825
                    Height =255
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text81"
                    ControlSource ="=Sum([NetSyndFee])/Sum([OEC])"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =5925
                    LayoutCachedTop =15
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =270
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6750
                    Top =15
                    Width =600
                    Height =255
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text82"
                    ControlSource ="=Sum([PDRent])/Sum([OEC])"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =6750
                    LayoutCachedTop =15
                    LayoutCachedWidth =7350
                    LayoutCachedHeight =270
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7350
                    Top =15
                    Width =600
                    Height =255
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text83"
                    ControlSource ="=Sum([UTA])/Sum([OEC])"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =7350
                    LayoutCachedTop =15
                    LayoutCachedWidth =7950
                    LayoutCachedHeight =270
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7950
                    Top =15
                    Width =855
                    Height =255
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text84"
                    ControlSource ="=Sum([TotalFee])/Sum([OEC])"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =7950
                    LayoutCachedTop =15
                    LayoutCachedWidth =8805
                    LayoutCachedHeight =270
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10140
                    Top =15
                    Width =1380
                    Height =255
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text85"
                    ControlSource ="=Sum([TotalFee])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =10140
                    LayoutCachedTop =15
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =270
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                    CurrencySymbol ="$"
                End
                Begin Rectangle
                    Top =270
                    Width =11520
                    Height =60
                    BorderColor =16777215
                    Name ="Box116"
                    GridlineColor =10921638
                    LayoutCachedTop =270
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =330
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Line
                    BorderWidth =2
                    Width =11520
                    BorderColor =8355711
                    Name ="Line106"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedWidth =11520
                    BorderTint =50.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5100
                    Top =15
                    Width =825
                    Height =255
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text147"
                    ControlSource ="=Sum([SyndFeePlus])/Sum([OEC])"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =5100
                    LayoutCachedTop =15
                    LayoutCachedWidth =5925
                    LayoutCachedHeight =270
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8805
                    Top =15
                    Width =1335
                    Height =255
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text160"
                    ControlSource ="=Sum([SyndFeePlus])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =8805
                    LayoutCachedTop =15
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =270
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                    CurrencySymbol ="$"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =1510
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1740
                    Top =60
                    Width =540
                    Height =255
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text117"
                    ControlSource ="=Sum([UnitCount])"
                    Format ="Standard"
                    FontName ="Segoe UI"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1740
                    LayoutCachedTop =60
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =315
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2280
                    Top =60
                    Width =1485
                    Height =255
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text118"
                    ControlSource ="=Sum([OEC])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =2280
                    LayoutCachedTop =60
                    LayoutCachedWidth =3765
                    LayoutCachedHeight =315
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =9
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3765
                    Top =60
                    Width =1335
                    Height =255
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text119"
                    ControlSource ="=\"\""
                    FontName ="Segoe UI"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =3765
                    LayoutCachedTop =60
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =315
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5925
                    Top =60
                    Width =825
                    Height =255
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text120"
                    ControlSource ="=Sum([NetSyndFee])/Sum([OEC])"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =5925
                    LayoutCachedTop =60
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =315
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6750
                    Top =60
                    Width =600
                    Height =255
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text121"
                    ControlSource ="=Sum([PDRent])/Sum([OEC])"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =6750
                    LayoutCachedTop =60
                    LayoutCachedWidth =7350
                    LayoutCachedHeight =315
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7350
                    Top =60
                    Width =600
                    Height =255
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text122"
                    ControlSource ="=Sum([UTA])/Sum([OEC])"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =7350
                    LayoutCachedTop =60
                    LayoutCachedWidth =7950
                    LayoutCachedHeight =315
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7950
                    Top =60
                    Width =855
                    Height =255
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text123"
                    ControlSource ="=Sum([TotalFee])/Sum([OEC])"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =7950
                    LayoutCachedTop =60
                    LayoutCachedWidth =8805
                    LayoutCachedHeight =315
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10140
                    Top =60
                    Width =1380
                    Height =255
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text124"
                    ControlSource ="=Sum([TotalFee])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =10140
                    LayoutCachedTop =60
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =315
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =9
                    CurrencySymbol ="$"
                End
                Begin Line
                    BorderWidth =1
                    Width =11520
                    BorderColor =4210752
                    Name ="Line127"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedWidth =11520
                    BorderTint =75.0
                End
                Begin Line
                    BorderWidth =1
                    Top =60
                    Width =11520
                    BorderColor =4210752
                    Name ="Line128"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedTop =60
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =60
                    BorderTint =75.0
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Top =60
                    Width =1920
                    Height =259
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text132"
                    ControlSource ="=\"Synd/ Forecast Totals →\""
                    GridlineColor =10921638

                    LayoutCachedTop =60
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =319
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =360
                    Width =9960
                    Height =1150
                    TabIndex =11
                    Name ="TFRCompareSummary_Portfolio"
                    SourceObject ="Report.TFRCompareSummary_Portfolio"
                    GridlineColor =10921638

                    LayoutCachedTop =360
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =1510
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5100
                    Top =60
                    Width =825
                    Height =255
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text144"
                    ControlSource ="=Sum([SyndFeePlus])/Sum([OEC])"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =5100
                    LayoutCachedTop =60
                    LayoutCachedWidth =5925
                    LayoutCachedHeight =315
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =9
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8805
                    Top =60
                    Width =1335
                    Height =255
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text163"
                    ControlSource ="=Sum([SyndFeePlus])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =9
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =8805
                    LayoutCachedTop =60
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =315
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =3
                    ThemeFontIndex =-1
                    GroupTable =9
                    CurrencySymbol ="$"
                End
            End
        End
    End
End
