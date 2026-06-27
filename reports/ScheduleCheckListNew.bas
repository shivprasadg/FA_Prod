Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15345
    DatasheetFontHeight =11
    ItemSuffix =126
    Left =4110
    Top =5265
    Filter ="SchID=1877"
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0xa4b65dc7afcde440
    End
    RecordSource ="ScheduleCheckListQ"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =0
            FontSize =11
            FontName ="Calibri"
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
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
        End
        Begin ListBox
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="MLNo"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="SchNo"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="SGrpNO"
        End
        Begin BreakLevel
            ControlSource ="UnitUnitNum"
        End
        Begin BreakLevel
            ControlSource ="unitunitnum2"
        End
        Begin BreakLevel
            ControlSource ="UnitType"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="UnitID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =420
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =15300
                    Height =420
                    ColumnOrder =0
                    FontSize =14
                    Name ="Text1"
                    ControlSource ="=\"<B>\" & [CLIENTSHNM] & \"</B>:  MLA# <B>\" & [MLNO] & \"</B>  Schedule <B>\" "
                        "& [SCHNO] & \"</B> Check List\""
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedWidth =15300
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =4
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =14760
                    Width =480
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =1
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =14760
                    LayoutCachedWidth =15240
                    LayoutCachedHeight =240
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =29
            BackColor =0
            Name ="GroupHeader3"
            AlternateBackColor =0
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =60
            BreakLevel =1
            Name ="GroupHeader0"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =3240
                    Height =14
                    FontSize =9
                    FontWeight =700
                    Name ="Text6"
                    ControlSource ="=IIf(IsNull([SCHIN]),\"Schedule O/S: SENT: \" & [SCHOUT],\"Schedule Back\")"
                    Format =">"

                    LayoutCachedWidth =3240
                    LayoutCachedHeight =14
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10380
                    Width =4860
                    Height =14
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="Text115"
                    ControlSource ="=\"Assignee Bank: \" & IIf(IsNull([Assignee]),\"Pending\",[Assignee])"

                    LayoutCachedLeft =10380
                    LayoutCachedWidth =15240
                    LayoutCachedHeight =14
                    ForeThemeColorIndex =2
                End
                Begin CheckBox
                    Visible = NotDefault
                    Left =3701
                    Top =46
                    Height =14
                    TabIndex =2
                    Name ="MLINSBLANKET"
                    ControlSource ="MLINSBLANKET"
                    StatusBarText ="IF YES AT LEASE LEVEL IS NO AT SCHEDULE LEVEL"

                    LayoutCachedLeft =3701
                    LayoutCachedTop =46
                    LayoutCachedWidth =3961
                    LayoutCachedHeight =60
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5460
                    Width =2100
                    Height =14
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="InsExpires"
                    ControlSource ="=IIf([MLinsblanket]=No,Null,IIf([mlinsExp]<Date(),\"INS EXPIRED \",\"Ins Expires"
                        " \")) & [MLInsExp]"

                    LayoutCachedLeft =5460
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =14
                    ForeThemeColorIndex =2
                End
                Begin Label
                    Visible = NotDefault
                    TextFontFamily =34
                    Left =3915
                    Width =1455
                    Height =14
                    FontSize =9
                    FontWeight =700
                    Name ="Label7"
                    Caption ="Blanket Insurance"
                    LayoutCachedLeft =3915
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =14
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    Visible = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7620
                    Width =4380
                    Height =14
                    FontSize =9
                    TabIndex =4
                    Name ="Text5"
                    ControlSource ="=IIf(IsNull([MLGUARANTOR]),Null,IIf(IsNull([MLGUARBACK]),\"GUARANTY O/S: SENT: \""
                        " & [MLGUARSENT],\"GUARANTY BACK\"))"

                    LayoutCachedLeft =7620
                    LayoutCachedWidth =12000
                    LayoutCachedHeight =14
                    ForeThemeColorIndex =2
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =945
            BreakLevel =2
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =810
                    Top =180
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    BackColor =4754549
                    Name ="SGrpNO"
                    ControlSource ="SGrpNO"
                    GroupTable =2

                    LayoutCachedLeft =810
                    LayoutCachedTop =180
                    LayoutCachedWidth =2250
                    LayoutCachedHeight =420
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =1
                    GroupTable =2
                    Begin
                        Begin Label
                            TextAlign =1
                            TextFontFamily =34
                            Left =75
                            Top =180
                            Width =675
                            Height =240
                            FontSize =9
                            ForeColor =2646351
                            Name ="Label10"
                            Caption ="Grp No:"
                            GroupTable =2
                            LayoutCachedLeft =75
                            LayoutCachedTop =180
                            LayoutCachedWidth =750
                            LayoutCachedHeight =420
                            LayoutGroup =2
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3180
                    Top =180
                    Width =675
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    BackColor =4754549
                    Name ="UnitsInGroup\015\012\015\012\015\012\015\012"
                    ControlSource ="UnitsInGroup"
                    EventProcPrefix ="UnitsInGroup________"
                    GroupTable =2

                    LayoutCachedLeft =3180
                    LayoutCachedTop =180
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =420
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =1
                    GroupTable =2
                    Begin
                        Begin Label
                            TextAlign =1
                            TextFontFamily =34
                            Left =2310
                            Top =180
                            Width =810
                            Height =240
                            FontSize =9
                            ForeColor =2646351
                            Name ="Label30"
                            Caption ="No Units:"
                            GroupTable =2
                            LayoutCachedLeft =2310
                            LayoutCachedTop =180
                            LayoutCachedWidth =3120
                            LayoutCachedHeight =420
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8775
                    Top =180
                    Width =1425
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    BackColor =4754549
                    Name ="OrigEquipCost"
                    ControlSource ="=Sum([TotUnit])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =2

                    LayoutCachedLeft =8775
                    LayoutCachedTop =180
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =420
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =1
                    GroupTable =2
                    Begin
                        Begin Label
                            TextAlign =1
                            TextFontFamily =34
                            Left =7200
                            Top =180
                            Width =1515
                            Height =240
                            FontSize =9
                            ForeColor =2646351
                            Name ="Label31"
                            Caption =" Total Equip Cost:"
                            GroupTable =2
                            LayoutCachedLeft =7200
                            LayoutCachedTop =180
                            LayoutCachedWidth =8715
                            LayoutCachedHeight =420
                            ColumnStart =5
                            ColumnEnd =5
                            LayoutGroup =2
                            BorderThemeColorIndex =6
                            BorderShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =900
                    Top =675
                    Width =1950
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =2646351
                    Name ="Label11"
                    Caption ="VIN"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =900
                    LayoutCachedTop =675
                    LayoutCachedWidth =2850
                    LayoutCachedHeight =945
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =7305
                    Top =675
                    Width =1650
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =2646351
                    Name ="Label14"
                    Caption ="Invoice"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7305
                    LayoutCachedTop =675
                    LayoutCachedWidth =8955
                    LayoutCachedHeight =945
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =2850
                    Top =675
                    Width =810
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label86"
                    Caption ="AssetId"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2850
                    LayoutCachedTop =675
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =945
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =3690
                    Top =675
                    Width =3555
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =2646351
                    Name ="Label93"
                    Caption ="Description"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3690
                    LayoutCachedTop =675
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =945
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =9015
                    Top =675
                    Width =1110
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =2646351
                    Name ="Label15"
                    Caption ="Equip Cost"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9015
                    LayoutCachedTop =675
                    LayoutCachedWidth =10125
                    LayoutCachedHeight =945
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =10185
                    Top =675
                    Width =810
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =2646351
                    Name ="Label18"
                    Caption ="Paid"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10185
                    LayoutCachedTop =675
                    LayoutCachedWidth =10995
                    LayoutCachedHeight =945
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =11055
                    Top =675
                    Width =750
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =2646351
                    Name ="Label19"
                    Caption ="Accepted"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11055
                    LayoutCachedTop =675
                    LayoutCachedWidth =11805
                    LayoutCachedHeight =945
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =11865
                    Top =675
                    Width =750
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =2646351
                    Name ="Label20"
                    Caption ="COA Back"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11865
                    LayoutCachedTop =675
                    LayoutCachedWidth =12615
                    LayoutCachedHeight =945
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextFontFamily =34
                    Left =13845
                    Top =675
                    Width =495
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =2646351
                    Name ="Label35"
                    Caption ="Dom"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13845
                    LayoutCachedTop =675
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =945
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =12675
                    Top =675
                    Width =1140
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =2646351
                    Name ="Label94"
                    Caption ="Title Event"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12675
                    LayoutCachedTop =675
                    LayoutCachedWidth =13815
                    LayoutCachedHeight =945
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Top =675
                    Width =900
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =2646351
                    Name ="Label89"
                    Caption ="Unit#"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =675
                    LayoutCachedWidth =900
                    LayoutCachedHeight =945
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3915
                    Top =180
                    Width =3225
                    Name ="EmptyCell104"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =3915
                    LayoutCachedTop =180
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =420
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    GroupTable =2
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11880
                    Width =2880
                    FontSize =9
                    FontWeight =700
                    Name ="Text3"
                    ControlSource ="=IIf(IsNull([MLBACK]),\"MASTER LEASE O/S: SENT: \" & [MLBACK],\"MASTER LEASE BAC"
                        "K\")"

                    LayoutCachedLeft =11880
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =240
                    ForeThemeColorIndex =2
                End
                Begin Rectangle
                    OldBorderStyle =0
                    Top =480
                    Width =15336
                    Height =29
                    BackColor =0
                    BorderColor =10921638
                    Name ="Box117"
                    GridlineColor =10921638
                    LayoutCachedTop =480
                    LayoutCachedWidth =15336
                    LayoutCachedHeight =509
                    BackThemeColorIndex =-1
                End
                Begin EmptyCell
                    Left =14370
                    Top =675
                    Width =600
                    Height =270
                    Name ="EmptyCell125"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =14370
                    LayoutCachedTop =675
                    LayoutCachedWidth =14970
                    LayoutCachedHeight =945
                    ColumnStart =11
                    ColumnEnd =12
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =240
            BreakLevel =6
            BackColor =14545371
            Name ="GroupHeader2"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =900
                    Width =1950
                    FontSize =9
                    TabIndex =3
                    Name ="Text66"
                    ControlSource ="UnitVIN"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =900
                    LayoutCachedWidth =2850
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7305
                    Width =1650
                    FontSize =9
                    TabIndex =6
                    Name ="Text68"
                    ControlSource ="UnitInvoiceNum"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7305
                    LayoutCachedWidth =8955
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3690
                    Width =3555
                    FontSize =9
                    TabIndex =5
                    Name ="UnitDesc"
                    ControlSource ="UnitDesc"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3690
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10185
                    Width =810
                    FontSize =9
                    TabIndex =8
                    Name ="Text70"
                    ControlSource ="UnitVendorPytDate"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =10185
                    LayoutCachedWidth =10995
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11055
                    Width =750
                    FontSize =9
                    TabIndex =9
                    Name ="Text71"
                    ControlSource ="UnitAcceptDate"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =11055
                    LayoutCachedWidth =11805
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11865
                    Width =750
                    FontSize =9
                    TabIndex =10
                    Name ="Text72"
                    ControlSource ="UnitCOABack"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =11865
                    LayoutCachedWidth =12615
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13845
                    Width =495
                    FontSize =9
                    TabIndex =12
                    Name ="Text74"
                    ControlSource ="LState"
                    StatusBarText ="lookup in lookup_state table"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13845
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9015
                    Width =1110
                    FontSize =9
                    TabIndex =7
                    Name ="UnitFinalCost"
                    ControlSource ="UnitCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9015
                    LayoutCachedWidth =10125
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =0
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2850
                    Width =810
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2850
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =7
                    ForeShade =75.0
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =900
                    ColumnWidth =2205
                    FontSize =9
                    TabIndex =2
                    Name ="UnitUnitNum"
                    ControlSource ="UnitUnitNum"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =900
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =30
                    Top =240
                    Width =660
                    Height =0
                    FontSize =9
                    TabIndex =1
                    Name ="UnitUnitNum2"
                    ControlSource ="unitunitnum2"

                    LayoutCachedLeft =30
                    LayoutCachedTop =240
                    LayoutCachedWidth =690
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12675
                    Width =1140
                    FontSize =9
                    TabIndex =11
                    Name ="Text91"
                    ControlSource ="=IIf(Nz([lastofunittitlelogdate],\"\")=\"\",Null,[Action] & \": \" & Format([las"
                        "tofunittitlelogdate],\"mm/dd/yy\"))"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12675
                    LayoutCachedWidth =13815
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =4320
                    Top =60
                    Width =240
                    Height =120
                    Name ="SumOfchildFinalCost"
                    ControlSource ="SumOfchildFinalCost"

                    LayoutCachedLeft =4320
                    LayoutCachedTop =60
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =180
                End
                Begin EmptyCell
                    Left =14370
                    Width =180
                    Name ="EmptyCell119"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =14370
                    LayoutCachedWidth =14550
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =14610
                    Width =360
                    Name ="EmptyCell124"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =14610
                    LayoutCachedWidth =14970
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1140
            Name ="Detail"
            Begin
                Begin Subform
                    Locked = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =540
                    Width =13485
                    Height =1140
                    Name ="SchChecklistSub"
                    SourceObject ="Report.SchChecklistSub"
                    LinkChildFields ="UnitREF"
                    LinkMasterFields ="UnitID"

                    LayoutCachedLeft =540
                    LayoutCachedWidth =14025
                    LayoutCachedHeight =1140
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =240
            BreakLevel =6
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8955
                    Width =1170
                    FontSize =9
                    FontWeight =700
                    Name ="TotChild"
                    ControlSource ="TotUnit"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8955
                    LayoutCachedWidth =10125
                    LayoutCachedHeight =240
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin Line
                    Left =8970
                    Width =1170
                    Name ="Line116"
                    LayoutCachedLeft =8970
                    LayoutCachedWidth =10140
                End
            End
        End
        Begin PageFooter
            Height =240
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =2040
                    FontSize =9
                    Name ="Text44"
                    ControlSource ="=Date()"
                    Format ="mmm dd\", \"yyyy"

                    LayoutCachedWidth =2040
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11070
                    Width =3600
                    FontSize =9
                    TabIndex =1
                    Name ="Text50"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =11070
                    LayoutCachedWidth =14670
                    LayoutCachedHeight =240
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
