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
    Width =9840
    DatasheetFontHeight =11
    ItemSuffix =167
    Left =4725
    Top =2610
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x048ff9f412a6e540
    End
    RecordSource ="Select * From TFRCompare_Summary Where Assigned_Indicator= 'H'"
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
            Height =0
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =852
            BreakLevel =1
            BackColor =8435191
            Name ="GroupHeader0"
            AlternateBackColor =7961551
            Begin
                Begin Label
                    Left =252
                    Top =396
                    Width =1488
                    Height =456
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label1"
                    Caption ="Account "
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =252
                    LayoutCachedTop =396
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =852
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =3
                    Left =1740
                    Top =396
                    Width =735
                    Height =456
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label2"
                    Caption ="Units"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1740
                    LayoutCachedTop =396
                    LayoutCachedWidth =2475
                    LayoutCachedHeight =852
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
                    Left =2472
                    Top =396
                    Width =1488
                    Height =456
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label3"
                    Caption ="Total OEC"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2472
                    LayoutCachedTop =396
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =852
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    OverlapFlags =4
                    TextAlign =2
                    Left =3960
                    Top =396
                    Width =1440
                    Height =456
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label4"
                    Caption ="Target Bank"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3960
                    LayoutCachedTop =396
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =852
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =5400
                    Top =396
                    Width =825
                    Height =456
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label38"
                    Caption ="Avg\015\012Rent "
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5400
                    LayoutCachedTop =396
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =852
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin Label
                    TextAlign =2
                    Left =8280
                    Top =396
                    Width =600
                    Height =456
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label53"
                    Caption ="Rent \015\012% "
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8280
                    LayoutCachedTop =396
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =852
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
                    Left =8880
                    Top =396
                    Width =600
                    Height =456
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label57"
                    Caption ="PD \015\012% "
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8880
                    LayoutCachedTop =396
                    LayoutCachedWidth =9480
                    LayoutCachedHeight =852
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
                    Left =6228
                    Top =396
                    Width =852
                    Height =456
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label64"
                    Caption ="Rent Monthly "
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6228
                    LayoutCachedTop =396
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =852
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    IMESentenceMode =3
                    Width =9840
                    Height =360
                    FontSize =12
                    BorderColor =16777215
                    Name ="Text112"
                    ControlSource ="=\"  <b>\" & [FundStatus] & \"</b> Schedules and Groups\""
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedWidth =9840
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                End
                Begin Label
                    TextAlign =2
                    Left =7080
                    Top =396
                    Width =1200
                    Height =456
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label155"
                    Caption ="PD\015\012Rent "
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7080
                    LayoutCachedTop =396
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =852
                    ColumnStart =6
                    ColumnEnd =6
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
                    Width =735
                    Height =255
                    ColumnWidth =1650
                    FontSize =9
                    TabIndex =1
                    Name ="UnitCount"
                    ControlSource ="UnitCount"
                    Format ="Standard"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1740
                    LayoutCachedWidth =2475
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
                    Left =2472
                    Width =1488
                    Height =255
                    ColumnWidth =1620
                    FontSize =9
                    TabIndex =2
                    Name ="OEC"
                    ControlSource ="OEC"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2472
                    LayoutCachedWidth =3960
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
                    Left =252
                    Width =1488
                    Height =255
                    ColumnWidth =1590
                    FontSize =9
                    Name ="AccountName"
                    ControlSource ="AccountName"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =252
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
                    Left =3960
                    Height =255
                    FontSize =9
                    TabIndex =3
                    Name ="TargetBank"
                    ControlSource ="TargetBank"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3960
                    LayoutCachedWidth =5400
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5400
                    Width =825
                    Height =255
                    FontSize =9
                    TabIndex =4
                    Name ="RentIncome"
                    ControlSource ="RentIncome"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5400
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
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
                    Left =8280
                    Width =600
                    Height =255
                    FontSize =9
                    TabIndex =7
                    Name ="Text55"
                    ControlSource ="=([RentIncome])/[OEC]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8280
                    LayoutCachedWidth =8880
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
                    Left =8880
                    Width =600
                    Height =255
                    FontSize =9
                    TabIndex =8
                    Name ="Text58"
                    ControlSource ="=[TotPDRent]/[OEC]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8880
                    LayoutCachedWidth =9480
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
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6228
                    Width =852
                    Height =255
                    FontSize =9
                    TabIndex =5
                    Name ="Text65"
                    ControlSource ="=[RentIncome]*[UnitCount]"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6228
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =2
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7080
                    Width =1200
                    Height =255
                    FontSize =9
                    TabIndex =6
                    Name ="Text161"
                    ControlSource ="TotPDRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7080
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
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
            Height =360
            BreakLevel =1
            BackColor =8435191
            Name ="GroupFooter0"
            AlternateBackColor =14151142
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Top =50
                    Width =1920
                    Height =270
                    FontSize =9
                    TabIndex =8
                    Name ="Text130"
                    ControlSource ="=[FundStatus] & \" Sub Totals →\""

                    LayoutCachedTop =50
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =320
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1740
                    Top =12
                    Width =735
                    Height =252
                    FontSize =9
                    Name ="Text78"
                    ControlSource ="=Sum([UnitCount])"
                    Format ="Standard"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1740
                    LayoutCachedTop =12
                    LayoutCachedWidth =2475
                    LayoutCachedHeight =264
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
                    Left =2472
                    Top =12
                    Width =1488
                    Height =252
                    FontSize =9
                    TabIndex =1
                    Name ="Text79"
                    ControlSource ="=Sum([OEC])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2472
                    LayoutCachedTop =12
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =264
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
                    Left =3960
                    Top =12
                    Height =252
                    FontSize =9
                    TabIndex =2
                    Name ="Text80"
                    ControlSource ="=\"\""
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3960
                    LayoutCachedTop =12
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =264
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5400
                    Top =12
                    Width =825
                    Height =252
                    FontSize =9
                    TabIndex =3
                    Name ="Text81"
                    ControlSource ="=Sum([RentIncome])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5400
                    LayoutCachedTop =12
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =264
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8280
                    Top =12
                    Width =600
                    Height =252
                    FontSize =9
                    TabIndex =6
                    Name ="Text82"
                    ControlSource ="=Sum([RentIncome])/Sum([OEC])"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8280
                    LayoutCachedTop =12
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =264
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
                    Left =8880
                    Top =12
                    Width =600
                    Height =252
                    FontSize =9
                    TabIndex =7
                    Name ="Text83"
                    ControlSource ="=Sum([TotPDRent])/Sum([OEC])"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8880
                    LayoutCachedTop =12
                    LayoutCachedWidth =9480
                    LayoutCachedHeight =264
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6228
                    Top =12
                    Width =852
                    Height =252
                    FontSize =9
                    TabIndex =4
                    Name ="Text84"
                    ControlSource ="=Sum([RentIncome]*[UnitCount])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6228
                    LayoutCachedTop =12
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =264
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                    CurrencySymbol ="$"
                End
                Begin Rectangle
                    Top =270
                    Width =9840
                    Height =90
                    BorderColor =16777215
                    Name ="Box116"
                    LayoutCachedTop =270
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Line
                    BorderWidth =2
                    Width =9840
                    BorderColor =8355711
                    Name ="Line106"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedWidth =9840
                    BorderTint =50.0
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7080
                    Top =12
                    Width =1200
                    Height =252
                    FontSize =9
                    TabIndex =5
                    Name ="Text163"
                    ControlSource ="=Sum([TotPDRent])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7080
                    LayoutCachedTop =12
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =264
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    GroupTable =8
                    CurrencySymbol ="$"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
