Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =28110
    DatasheetFontHeight =11
    ItemSuffix =1188
    Left =7800
    Top =7290
    Right =-29371
    Bottom =20550
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x4e47c5bd521ce540
    End
    RecordSource ="Select * From TFR_Work"
    Caption ="Inventory Sales"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
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
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
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
        Begin ToggleButton
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
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
            PressedForeThemeColorIndex =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =3375
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Line
                    OverlapFlags =93
                    Left =645
                    Top =885
                    Width =12960
                    BorderColor =12349952
                    Name ="Line88"
                    LayoutCachedLeft =645
                    LayoutCachedTop =885
                    LayoutCachedWidth =13605
                    LayoutCachedHeight =885
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =11916
                    Top =795
                    Width =1553
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="ClearAndRun"
                    Caption ="Clear Filters and Run"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =11916
                    LayoutCachedTop =795
                    LayoutCachedWidth =13469
                    LayoutCachedHeight =1035
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =9211020
                    BorderColor =8210719
                    HoverThemeColorIndex =2
                    PressedColor =9211020
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =14571
                    Top =1290
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearState"
                    Caption ="Clear▼"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =14571
                    LayoutCachedTop =1290
                    LayoutCachedWidth =15111
                    LayoutCachedHeight =1515
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Width =28080
                    Height =540
                    BorderColor =14211288
                    Name ="Box49"
                    HorizontalAnchor =2
                    LayoutCachedWidth =28080
                    LayoutCachedHeight =540
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =3180
                    Width =28080
                    Height =195
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"

                    LayoutCachedTop =3180
                    LayoutCachedWidth =28080
                    LayoutCachedHeight =3375
                    TabIndex =30
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =17631
                    Top =960
                    Width =2280
                    Height =1320
                    BorderColor =10855845
                    Name ="Box179"
                    LayoutCachedLeft =17631
                    LayoutCachedTop =960
                    LayoutCachedWidth =19911
                    LayoutCachedHeight =2280
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =4488
                    Top =2820
                    Width =936
                    Height =465
                    FontSize =10
                    FontWeight =700
                    ForeColor =13158
                    Name ="Label6"
                    Caption ="Fund\015\012Date"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4488
                    LayoutCachedTop =2820
                    LayoutCachedWidth =5424
                    LayoutCachedHeight =3285
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =25800
                    Top =2820
                    Width =1152
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label15"
                    Caption ="BDEX\015\012Comm"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =25800
                    LayoutCachedTop =2820
                    LayoutCachedWidth =26952
                    LayoutCachedHeight =3285
                    ColumnStart =25
                    ColumnEnd =25
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =2472
                    Top =2820
                    Width =2016
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label18"
                    Caption ="Account Name"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2472
                    LayoutCachedTop =2820
                    LayoutCachedWidth =4488
                    LayoutCachedHeight =3285
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =780
                    Top =2820
                    Width =540
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label21"
                    Caption ="Fund\015\012Ind"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =780
                    LayoutCachedTop =2820
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =3285
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10380
                    Width =29
                    Height =270
                    ColumnOrder =20
                    FontWeight =700
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =10380
                    LayoutCachedWidth =10409
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =9360
                    Top =240
                    Width =765
                    Height =210
                    ColumnOrder =21
                    FontSize =8
                    TabIndex =1
                    Name ="txtCurrent"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =240
                    LayoutCachedWidth =10125
                    LayoutCachedHeight =450
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =6600
                    Top =2820
                    Width =2952
                    Height =465
                    FontSize =10
                    FontWeight =700
                    ForeColor =13158
                    Name ="Label70"
                    Caption ="Unit Description"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6600
                    LayoutCachedTop =2820
                    LayoutCachedWidth =9552
                    LayoutCachedHeight =3285
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ListBox
                    OverlapFlags =223
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2460
                    Top =960
                    Width =2016
                    Height =1800
                    ColumnOrder =15
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    BoundColumn =1
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstAccountName"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TFR.AccountName, ClientId FROM TFR_Work As TFR WHERE TFR.AccountName Is N"
                        "ot Null GROUP BY TFR.AccountName, TFR.ClientId ORDER BY TFR.AccountName;"
                    ColumnWidths ="2016;0"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =2460
                    LayoutCachedTop =960
                    LayoutCachedWidth =4476
                    LayoutCachedHeight =2760
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ListBox
                    OverlapFlags =223
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =4470
                    Top =960
                    Width =936
                    Height =1800
                    ColumnOrder =16
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstFundYear"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TFR.FYear FROM TFR_Work As TFR WHERE TFR.[Sch_Grp] Is Not Null GROUP BY T"
                        "FR.FYear ORDER BY TFR.FYear DESC;"
                    ColumnWidths ="1440"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =4470
                    LayoutCachedTop =960
                    LayoutCachedWidth =5406
                    LayoutCachedHeight =2760
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ListBox
                    OverlapFlags =223
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =5406
                    Top =960
                    Width =1140
                    Height =1800
                    ColumnOrder =17
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstBankType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TFR.Bank FROM TFR_Work As TFR WHERE TFR.SGrpID Is Not Null GROUP BY TFR.B"
                        "ank ORDER BY TFR.Bank;"
                    ColumnWidths ="1296"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =5406
                    LayoutCachedTop =960
                    LayoutCachedWidth =6546
                    LayoutCachedHeight =2760
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ListBox
                    OverlapFlags =223
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =6546
                    Top =960
                    Width =2955
                    Height =1800
                    ColumnOrder =18
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstDesc"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitDesc From TFR_Work WHERE ClientID is not null   GROUP BY UnitDesc ORD"
                        "ER BY UnitDesc;"
                    ColumnWidths ="1296"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =6546
                    LayoutCachedTop =960
                    LayoutCachedWidth =9501
                    LayoutCachedHeight =2760
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ListBox
                    OverlapFlags =223
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =765
                    Top =975
                    Width =539
                    Height =1800
                    ColumnOrder =19
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstStatus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TFR.Assigned_Indicator FROM TFR_Work as TFR WHERE TFR.Assigned_Indicator "
                        "<>'X' GROUP BY TFR.Assigned_Indicator ORDER BY TFR.Assigned_Indicator;"
                    ColumnWidths ="404"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =765
                    LayoutCachedTop =975
                    LayoutCachedWidth =1304
                    LayoutCachedHeight =2775
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11916
                    Top =1095
                    Width =1553
                    Height =1140
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    Name ="cmdSetFilters"
                    Caption ="After Setting Filters Click HERE\015\012Run the Results"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =11916
                    LayoutCachedTop =1095
                    LayoutCachedWidth =13469
                    LayoutCachedHeight =2235
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12750136
                    BorderColor =12750136
                    HoverColor =10040879
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =2460
                    Top =795
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearClients"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =2460
                    LayoutCachedTop =795
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =1020
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =4470
                    Top =795
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearYear"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =4470
                    LayoutCachedTop =795
                    LayoutCachedWidth =5010
                    LayoutCachedHeight =1020
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =5406
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearBankType"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =5406
                    LayoutCachedTop =780
                    LayoutCachedWidth =5946
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =6546
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearSch_Grp"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =6546
                    LayoutCachedTop =780
                    LayoutCachedWidth =7086
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =765
                    Top =795
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearStatus"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =765
                    LayoutCachedTop =795
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =1020
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =14046
                    Top =810
                    Width =1800
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearAll"
                    Caption ="◄ Clear All Selections "
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =14046
                    LayoutCachedTop =810
                    LayoutCachedWidth =15846
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =11508
                    Top =2820
                    Width =1296
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label121"
                    Caption ="OEC"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11508
                    LayoutCachedTop =2820
                    LayoutCachedWidth =12804
                    LayoutCachedHeight =3285
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Image
                    BackStyle =1
                    PictureType =2
                    Left =17346
                    Top =1350
                    Width =480
                    Height =420
                    Name ="Image161"
                    Picture ="1_Excel-365-Green"

                    LayoutCachedLeft =17346
                    LayoutCachedTop =1350
                    LayoutCachedWidth =17826
                    LayoutCachedHeight =1770
                    TabIndex =29
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =18966
                    Top =2355
                    Width =5805
                    Height =270
                    FontSize =8
                    Name ="Label162"
                    Caption ="┌─────────────────────  Unit Valuation ─────────────────────┐"
                    LayoutCachedLeft =18966
                    LayoutCachedTop =2355
                    LayoutCachedWidth =24771
                    LayoutCachedHeight =2625
                    ForeThemeColorIndex =2
                End
                Begin ListBox
                    Visible = NotDefault
                    OverlapFlags =215
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =17871
                    Top =960
                    Width =2040
                    Height =1320
                    ColumnOrder =14
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    BackColor =14277081
                    ForeColor =0
                    BorderColor =4210752
                    Name ="ReportsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LaunchReportsList.ReportPublicName, LaunchReportsList.ReportDesc, LaunchR"
                        "eportsList.Command, LaunchReportsList.FormCalled, LaunchReportsList.SecurityLeve"
                        "l, LaunchReportsList.Tags, LaunchReportsList.FormGroup FROM LaunchReportsList WH"
                        "ERE (((LaunchReportsList.SecurityLevel)<=9) AND ((LaunchReportsList.FormGroup)=\""
                        "RmktDashboard\")) ORDER BY LaunchReportsList.ReportPublicName;"
                    ColumnWidths ="2016;0;0;0"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="ReportList"
                    ControlTipText ="Sold Report Report of the currently filtered Units"
                    GridlineColor =0

                    LayoutCachedLeft =17871
                    LayoutCachedTop =960
                    LayoutCachedWidth =19911
                    LayoutCachedHeight =2280
                    ThemeFontIndex =-1
                    BackShade =85.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11833
                    Top =240
                    Width =689
                    Height =300
                    ColumnOrder =13
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="UnitsShown"
                    ControlSource ="=Count([ClientID])"

                    LayoutCachedLeft =11833
                    LayoutCachedTop =240
                    LayoutCachedWidth =12522
                    LayoutCachedHeight =540
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    Begin
                        Begin Label
                            Visible = NotDefault
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =223
                            TextAlign =2
                            Left =11808
                            Top =60
                            Width =750
                            Height =540
                            FontSize =7
                            BackColor =14211288
                            BorderColor =14211288
                            Name ="Label63"
                            Caption ="Deals"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =11808
                            LayoutCachedTop =60
                            LayoutCachedWidth =12558
                            LayoutCachedHeight =600
                        End
                    End
                End
                Begin Label
                    OverlapFlags =223
                    Left =120
                    Top =60
                    Width =3780
                    Height =555
                    FontSize =20
                    ForeColor =4754549
                    Name ="Label31"
                    Caption ="TFR Dashboard"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =615
                End
                Begin Label
                    OverlapFlags =215
                    Left =150
                    Width =1260
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label294"
                    Caption ="Syndication"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =150
                    LayoutCachedWidth =1410
                    LayoutCachedHeight =240
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =16884
                    Top =2820
                    Width =1152
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label376"
                    Caption ="Net \015\012Synd Fee"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16884
                    LayoutCachedTop =2820
                    LayoutCachedWidth =18036
                    LayoutCachedHeight =3285
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =1920
                    Top =90
                    Width =840
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label421"
                    Caption ="Manager"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =1920
                    LayoutCachedTop =90
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    BackStyle =0
                    IMESentenceMode =3
                    Left =21015
                    Top =2535
                    Width =1155
                    Height =270
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =24
                    Name ="GLLine"
                    ControlSource ="=Sum([UnitSalePrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =21015
                    LayoutCachedTop =2535
                    LayoutCachedWidth =22170
                    LayoutCachedHeight =2805
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18996
                    Top =2535
                    Width =1155
                    Height =270
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =21
                    Name ="Text422"
                    ControlSource ="=Sum([BuyOut])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =18996
                    LayoutCachedTop =2535
                    LayoutCachedWidth =20151
                    LayoutCachedHeight =2805
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =22170
                    Top =2535
                    Width =1155
                    Height =270
                    ColumnOrder =11
                    FontSize =9
                    TabIndex =25
                    Name ="Text444"
                    ControlSource ="=Sum([GainLoss])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =22170
                    LayoutCachedTop =2535
                    LayoutCachedWidth =23325
                    LayoutCachedHeight =2805
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =26952
                    Top =2820
                    Width =1152
                    Height =465
                    FontSize =10
                    FontWeight =700
                    BackColor =12566463
                    Name ="Label442"
                    Caption ="Net\015\012Revenue"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =26952
                    LayoutCachedTop =2820
                    LayoutCachedWidth =28104
                    LayoutCachedHeight =3285
                    ColumnStart =26
                    ColumnEnd =26
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =9552
                    Top =2820
                    Width =1956
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label480"
                    Caption ="Lease Sch |Group"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9552
                    LayoutCachedTop =2820
                    LayoutCachedWidth =11508
                    LayoutCachedHeight =3285
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =13320
                    Top =2820
                    Width =1176
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label486"
                    Caption ="Gross\015\012Synd Fee"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13320
                    LayoutCachedTop =2820
                    LayoutCachedWidth =14496
                    LayoutCachedHeight =3285
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =9501
                    Top =960
                    Width =1950
                    Height =1800
                    ColumnOrder =9
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstSch_Grp"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Sch_Grp From TFR_Work WHERE ClientID is not null   GROUP BY Sch_Grp ORDER"
                        " BY Sch_Grp;"
                    ColumnWidths ="1296"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =9501
                    LayoutCachedTop =960
                    LayoutCachedWidth =11451
                    LayoutCachedHeight =2760
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =9501
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearModel"
                    Caption ="Clear▼"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =9501
                    LayoutCachedTop =780
                    LayoutCachedWidth =10041
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =95
                    TextAlign =1
                    Left =105
                    Top =795
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearBDEX"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =105
                    LayoutCachedTop =795
                    LayoutCachedWidth =645
                    LayoutCachedHeight =1020
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =5424
                    Top =2820
                    Width =1176
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label609"
                    Caption ="Assignee \015\012Bank"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5424
                    LayoutCachedTop =2820
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =3285
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =20292
                    Top =2820
                    Width =1056
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label610"
                    Caption ="Net\015\012Per diem "
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =20292
                    LayoutCachedTop =2820
                    LayoutCachedWidth =21348
                    LayoutCachedHeight =3285
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15906
                    Top =1995
                    Width =1215
                    Height =270
                    ColumnOrder =8
                    FontSize =8
                    FontWeight =700
                    TabIndex =19
                    BorderColor =4210752
                    Name ="EndDateRange"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =15906
                    LayoutCachedTop =1995
                    LayoutCachedWidth =17121
                    LayoutCachedHeight =2265
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15906
                    Top =1665
                    Width =1215
                    Height =270
                    ColumnOrder =7
                    FontSize =8
                    FontWeight =700
                    TabIndex =17
                    BorderColor =4210752
                    Name ="StartDateRange"
                    Format ="Short Date"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =15906
                    LayoutCachedTop =1665
                    LayoutCachedWidth =17121
                    LayoutCachedHeight =1935
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =15900
                    Top =1428
                    Width =285
                    Height =228
                    FontSize =7
                    TabIndex =13
                    ForeColor =12349952
                    Name ="Fund0"
                    Caption ="19"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =15900
                    LayoutCachedTop =1428
                    LayoutCachedWidth =16185
                    LayoutCachedHeight =1656
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =6
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =16212
                    Top =1428
                    Width =288
                    Height =228
                    FontSize =7
                    TabIndex =14
                    ForeColor =12349952
                    Name ="Fund1"
                    Caption ="18"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =16212
                    LayoutCachedTop =1428
                    LayoutCachedWidth =16500
                    LayoutCachedHeight =1656
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =6
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =16524
                    Top =1428
                    Width =288
                    Height =228
                    FontSize =7
                    TabIndex =15
                    ForeColor =12349952
                    Name ="Fund2"
                    Caption ="17"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =16524
                    LayoutCachedTop =1428
                    LayoutCachedWidth =16812
                    LayoutCachedHeight =1656
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =6
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =108
                    Top =2820
                    Width =660
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label628"
                    Caption ="BDEX"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =108
                    LayoutCachedTop =2820
                    LayoutCachedWidth =768
                    LayoutCachedHeight =3285
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =21960
                    Top =2820
                    Width =492
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label5"
                    Caption ="PD\015\012Days"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =21960
                    LayoutCachedTop =2820
                    LayoutCachedWidth =22452
                    LayoutCachedHeight =3285
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =15096
                    Top =2820
                    Width =1188
                    Height =465
                    FontSize =10
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label709"
                    Caption ="Term \015\012Penalty"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15096
                    LayoutCachedTop =2820
                    LayoutCachedWidth =16284
                    LayoutCachedHeight =3285
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =23964
                    Top =2820
                    Width =1236
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label711"
                    Caption ="Total\015\012Revenue"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =23964
                    LayoutCachedTop =2820
                    LayoutCachedWidth =25200
                    LayoutCachedHeight =3285
                    ColumnStart =23
                    ColumnEnd =23
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =255
                    BackStyle =0
                    IMESentenceMode =3
                    Left =20151
                    Top =2535
                    Width =1155
                    Height =270
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =23
                    Name ="Text725"
                    ControlSource ="=Sum([TermPenalty])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =20151
                    LayoutCachedTop =2535
                    LayoutCachedWidth =21306
                    LayoutCachedHeight =2805
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19860
                    Top =2535
                    Width =1155
                    Height =270
                    ColumnOrder =6
                    FontSize =9
                    TabIndex =22
                    Name ="Text726"
                    ControlSource ="=Sum([BookValue])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =19860
                    LayoutCachedTop =2535
                    LayoutCachedWidth =21015
                    LayoutCachedHeight =2805
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =22452
                    Top =2820
                    Width =912
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label746"
                    Caption ="UTA\015\012(rebate)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =22452
                    LayoutCachedTop =2820
                    LayoutCachedWidth =23364
                    LayoutCachedHeight =3285
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =14496
                    Top =2820
                    Width =600
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label796"
                    Caption ="\015\012%"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14496
                    LayoutCachedTop =2820
                    LayoutCachedWidth =15096
                    LayoutCachedHeight =3285
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =23364
                    Top =2820
                    Width =600
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label802"
                    Caption ="\015\012%"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =23364
                    LayoutCachedTop =2820
                    LayoutCachedWidth =23964
                    LayoutCachedHeight =3285
                    ColumnStart =22
                    ColumnEnd =22
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =16284
                    Top =2820
                    Width =600
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label806"
                    Caption ="\015\012%"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16284
                    LayoutCachedTop =2820
                    LayoutCachedWidth =16884
                    LayoutCachedHeight =3285
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =18036
                    Top =2820
                    Width =600
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label814"
                    Caption ="\015\012%"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18036
                    LayoutCachedTop =2820
                    LayoutCachedWidth =18636
                    LayoutCachedHeight =3285
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =21348
                    Top =2820
                    Width =600
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label832"
                    Caption ="\015\012%"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =21348
                    LayoutCachedTop =2820
                    LayoutCachedWidth =21948
                    LayoutCachedHeight =3285
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =25200
                    Top =2820
                    Width =600
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label842"
                    Caption ="\015\012%"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =25200
                    LayoutCachedTop =2820
                    LayoutCachedWidth =25800
                    LayoutCachedHeight =3285
                    ColumnStart =24
                    ColumnEnd =24
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =19692
                    Top =2820
                    Width =600
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label852"
                    Caption ="\015\012%"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19692
                    LayoutCachedTop =2820
                    LayoutCachedWidth =20292
                    LayoutCachedHeight =3285
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =18636
                    Top =2820
                    Width =1056
                    Height =465
                    FontSize =10
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label860"
                    Caption ="Interim\015\012Expense "
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18636
                    LayoutCachedTop =2820
                    LayoutCachedWidth =19692
                    LayoutCachedHeight =3285
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =105
                    Top =975
                    Width =660
                    Height =1800
                    ColumnOrder =4
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstBDEX"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TFR.BDEX FROM TFR_Work AS TFR WHERE (((TFR.Assigned_Indicator)<>'X')) GRO"
                        "UP BY TFR.BDEX ORDER BY TFR.BDEX;"
                    ColumnWidths ="404"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =105
                    LayoutCachedTop =975
                    LayoutCachedWidth =765
                    LayoutCachedHeight =2775
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =16836
                    Top =1428
                    Width =288
                    Height =228
                    FontSize =7
                    TabIndex =16
                    ForeColor =12349952
                    Name ="Fund3"
                    Caption ="16"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =16836
                    LayoutCachedTop =1428
                    LayoutCachedWidth =17124
                    LayoutCachedHeight =1656
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =6
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14076
                    Top =1845
                    Width =1545
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =18
                    ForeColor =0
                    Name ="SendToExcel"
                    Caption =" Send To Excel"
                    OnClick ="[Event Procedure]"
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

                    LayoutCachedLeft =14076
                    LayoutCachedTop =1845
                    LayoutCachedWidth =15621
                    LayoutCachedHeight =2145
                    PictureCaptionArrangement =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =6405021
                    BorderColor =16777215
                    HoverColor =6405021
                    PressedColor =6405021
                    HoverForeColor =0
                    PressedForeColor =0
                    QuickStyle =32
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11931
                    Top =2355
                    Width =1545
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =20
                    ForeColor =10040879
                    Name ="GetData"
                    Caption ="Refresh Data"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000090a468ff90a4686690a468ff90a468ff90a468ff ,
                        0x90a4684b00000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000090a468ff90a468ff90a4685a0000000090a46869 ,
                        0x90a468ff90a4685a00000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff0000000090a468ff90a468ff90a468ff0000000000000000 ,
                        0x90a4685a90a468ff00000000727272ffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffffffffffffffffff0f0000000000000000000000000000000000000000 ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffffffffffffffffff0f90a468ff90a4685a000000000000000090a468ff ,
                        0x90a468ff90a468ff00000000727272ffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffffffffffffffffff5490a4685a90a468ff90a468690000000090a4685a ,
                        0x90a468ff90a468ff00000000727272ffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffffffffffffffffffd8ffffff1290a4684b90a468ff90a468ff90a468ff ,
                        0x90a4686690a468ff00000000727272ffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffffffffffffffffffffb3b3b3ffffffff57ffffff1bffffff1b00000000 ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffffffffffffffffffffb3b3b3ffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffffffffffffffffffffb3b3b3ffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffffffffffffffffffffb3b3b3ffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x000000000000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x000000000000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =11931
                    LayoutCachedTop =2355
                    LayoutCachedWidth =13476
                    LayoutCachedHeight =2655
                    PictureCaptionArrangement =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =13942903
                    BorderColor =16777215
                    HoverColor =1643706
                    PressedThemeColorIndex =8
                    PressedTint =60.0
                    PressedForeColor =10040879
                    QuickStyle =32
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =12804
                    Top =2820
                    Width =516
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label889"
                    Caption ="Units"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12804
                    LayoutCachedTop =2820
                    LayoutCachedWidth =13320
                    LayoutCachedHeight =3285
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =1320
                    Top =2820
                    Width =1152
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="Label917"
                    Caption ="Client \015\012Group"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1320
                    LayoutCachedTop =2820
                    LayoutCachedWidth =2472
                    LayoutCachedHeight =3285
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =1304
                    Top =795
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearCG"
                    Caption ="Clear▼"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =1304
                    LayoutCachedTop =795
                    LayoutCachedWidth =1844
                    LayoutCachedHeight =1020
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1304
                    Top =960
                    Width =1152
                    Height =1800
                    ColumnOrder =0
                    FontSize =8
                    FontWeight =700
                    TabIndex =26
                    BoundColumn =1
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstCGroup"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TFR.ClientGroup, TFR.ClientGroupId FROM TFR_Work AS TFR WHERE (((TFR.Acco"
                        "untName) Is Not Null)) GROUP BY TFR.ClientGroup, TFR.ClientGroupId ORDER BY TFR."
                        "ClientGroup;"
                    ColumnWidths ="2016;0"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =1304
                    LayoutCachedTop =960
                    LayoutCachedWidth =2456
                    LayoutCachedHeight =2760
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =26340
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =27
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

                    LayoutCachedLeft =26340
                    LayoutCachedTop =120
                    LayoutCachedWidth =26895
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
                    Left =26907
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =28
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1

                    LayoutCachedLeft =26907
                    LayoutCachedTop =120
                    LayoutCachedWidth =27462
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
            Height =298
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =22452
                    Width =912
                    Height =270
                    FontSize =10
                    TabIndex =26
                    BorderColor =0
                    Name ="UTA"
                    ControlSource ="UTA"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =22452
                    LayoutCachedWidth =23364
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =25800
                    Width =1152
                    Height =270
                    FontSize =10
                    TabIndex =30
                    BorderColor =0
                    Name ="Text789"
                    ControlSource ="BDEXComm"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =25800
                    LayoutCachedWidth =26952
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =25
                    ColumnEnd =25
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =23964
                    Width =1236
                    Height =270
                    FontSize =10
                    TabIndex =28
                    BorderColor =0
                    Name ="TotalFee"
                    ControlSource ="TotalFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =23964
                    LayoutCachedWidth =25200
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =23
                    ColumnEnd =23
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5424
                    Width =1176
                    Height =270
                    FontSize =10
                    TabIndex =10
                    BorderColor =0
                    Name ="Bank"
                    ControlSource ="Bank"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =5424
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9552
                    Width =1956
                    Height =270
                    FontSize =9
                    TabIndex =12
                    BorderColor =0
                    Name ="Sch_Grp"
                    ControlSource ="Sch_Grp"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =9552
                    LayoutCachedWidth =11508
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15096
                    Width =1188
                    Height =270
                    FontSize =10
                    TabIndex =17
                    BorderColor =0
                    Name ="Text703"
                    ControlSource ="TermPenalty"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =15096
                    LayoutCachedWidth =16284
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =20292
                    Width =1056
                    Height =270
                    FontSize =10
                    TabIndex =23
                    BorderColor =0
                    Name ="PDRent"
                    ControlSource ="PDRent"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =20292
                    LayoutCachedWidth =21348
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13320
                    Width =1176
                    Height =270
                    FontSize =10
                    TabIndex =15
                    BorderColor =0
                    Name ="Text487"
                    ControlSource ="SyndFee"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =13320
                    LayoutCachedWidth =14496
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16884
                    Width =1152
                    Height =270
                    FontSize =10
                    TabIndex =19
                    BorderColor =0
                    Name ="BuyOut"
                    ControlSource ="NetSyndFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =16884
                    LayoutCachedWidth =18036
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6600
                    Width =2952
                    Height =270
                    FontSize =9
                    TabIndex =11
                    BorderColor =0
                    Name ="UnitDesc"
                    ControlSource ="UnitDesc"
                    StatusBarText ="unit make"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1
                    ShowDatePicker =0

                    LayoutCachedLeft =6600
                    LayoutCachedWidth =9552
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4488
                    Width =936
                    Height =270
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =9
                    BorderColor =0
                    Name ="UnitID"
                    ControlSource ="FundDate"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =4488
                    LayoutCachedWidth =5424
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2472
                    Width =2016
                    Height =270
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =8
                    BorderColor =14211288
                    ForeColor =10040879
                    Name ="AccountName"
                    ControlSource ="AccountName"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2472
                    LayoutCachedWidth =4488
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11508
                    Width =1296
                    Height =270
                    FontSize =10
                    TabIndex =13
                    BorderColor =0
                    Name ="Odometer"
                    ControlSource ="GroupOEC"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =11508
                    LayoutCachedWidth =12804
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14496
                    Width =600
                    Height =270
                    FontSize =8
                    TabIndex =16
                    BorderColor =0
                    Name ="PnLTarget"
                    ControlSource ="SyndFeePct"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =14496
                    LayoutCachedWidth =15096
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =108
                    Width =660
                    Height =270
                    FontSize =10
                    TabIndex =5
                    BorderColor =0
                    Name ="Text626"
                    ControlSource ="BDEX"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =108
                    LayoutCachedWidth =768
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =26952
                    Width =1152
                    Height =270
                    FontSize =10
                    TabIndex =31
                    BorderColor =0
                    Name ="NetTotalFee"
                    ControlSource ="NetTotalFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =26952
                    LayoutCachedWidth =28104
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =26
                    ColumnEnd =26
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =23364
                    Width =600
                    Height =270
                    FontSize =8
                    TabIndex =27
                    BorderColor =0
                    Name ="Text803"
                    ControlSource ="UTApct"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =23364
                    LayoutCachedWidth =23964
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =22
                    ColumnEnd =22
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16284
                    Width =600
                    Height =270
                    FontSize =8
                    TabIndex =18
                    BorderColor =0
                    Name ="Text807"
                    ControlSource ="TermPenaltyPct"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =16284
                    LayoutCachedWidth =16884
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =18036
                    Width =600
                    Height =270
                    FontSize =8
                    TabIndex =20
                    BorderColor =0
                    Name ="Text815"
                    ControlSource ="TermPenaltyPct"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =18036
                    LayoutCachedWidth =18636
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =21348
                    Width =600
                    Height =270
                    FontSize =8
                    TabIndex =24
                    BorderColor =0
                    Name ="Text833"
                    ControlSource ="PDRentPct"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =21348
                    LayoutCachedWidth =21948
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =25200
                    Width =600
                    Height =270
                    FontSize =8
                    TabIndex =29
                    BorderColor =0
                    Name ="Text843"
                    ControlSource ="TotalFeePct"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =25200
                    LayoutCachedWidth =25800
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =24
                    ColumnEnd =24
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =19692
                    Width =600
                    Height =270
                    FontSize =8
                    TabIndex =22
                    BorderColor =0
                    Name ="Text853"
                    ControlSource ="InterimExpPct"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =19692
                    LayoutCachedWidth =20292
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =18636
                    Width =1056
                    Height =270
                    FontSize =10
                    TabIndex =21
                    BorderColor =0
                    Name ="Text861"
                    ControlSource ="InterimExp"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =18636
                    LayoutCachedWidth =19692
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Top =240
                    Width =25200
                    Height =58
                    TabIndex =4
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00470072006f0075007000490044005d003d005b0074007800740043007500 ,
                        0x7200720065006e0074005d0000000000
                    End

                    LayoutCachedTop =240
                    LayoutCachedWidth =25200
                    LayoutCachedHeight =298
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00160000005b00 ,
                        0x470072006f0075007000490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =780
                    Width =540
                    Height =270
                    ColumnOrder =1
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    Name ="FundingStatus"
                    ControlSource ="Assigned_Indicator"
                    OnDblClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x010000006a000000010000000000000002000000000000000400000001010000 ,
                        0xffffff00758c4800000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x22004600220000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    ConditionalFormat14 = Begin
                        0x010001000000000000000200000001010000ffffff00758c4800030000002200 ,
                        0x46002200000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =21960
                    Width =492
                    Height =270
                    FontSize =10
                    TabIndex =25
                    Name ="PDDays"
                    ControlSource ="PDDaysAvg"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =21960
                    LayoutCachedWidth =22452
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12804
                    Width =516
                    Height =270
                    FontSize =10
                    TabIndex =14
                    BorderColor =0
                    Name ="Text883"
                    ControlSource ="UnitCount"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =12804
                    LayoutCachedWidth =13320
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1320
                    Width =1152
                    Height =270
                    FontSize =10
                    TabIndex =7
                    BorderColor =0
                    Name ="Text911"
                    ControlSource ="ClientGroup"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000100 ,
                        0xba141900ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =1320
                    LayoutCachedWidth =2472
                    LayoutCachedHeight =270
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x010001000000000000000500000001000100ba141900ffffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =135
                    Width =45
                    Height =210
                    FontSize =8
                    TabIndex =1
                    Name ="ClientID"
                    ControlSource ="ClientID"

                    LayoutCachedLeft =135
                    LayoutCachedWidth =180
                    LayoutCachedHeight =210
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =315
                    Width =45
                    Height =210
                    FontSize =8
                    TabIndex =2
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =315
                    LayoutCachedWidth =360
                    LayoutCachedHeight =210
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =510
                    Width =45
                    Height =210
                    FontSize =8
                    TabIndex =3
                    Name ="GroupID"
                    ControlSource ="SGrpID"

                    LayoutCachedLeft =510
                    LayoutCachedWidth =555
                    LayoutCachedHeight =210
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =45
                    Height =210
                    FontSize =8
                    Name ="ClientGroupId"
                    ControlSource ="ClientGroupId"

                    LayoutCachedWidth =45
                    LayoutCachedHeight =210
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
' See "TFRDashboard.cls"
