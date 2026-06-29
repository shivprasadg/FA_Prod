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
    Width =25920
    DatasheetFontHeight =11
    ItemSuffix =699
    Left =6135
    Top =2715
    Right =32055
    Bottom =19095
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x598229f9cbbee440
    End
    RecordSource ="SELECT * FROM RemarketingReportAllUnits WHERE [UNITId] is not null  AND ([Client"
        "Id]=4046)  Order By SourceCompany;"
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
            Height =3195
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Width =25920
                    Height =540
                    BorderColor =14211288
                    Name ="Box49"
                    HorizontalAnchor =2
                    LayoutCachedWidth =25920
                    LayoutCachedHeight =540
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =3000
                    Width =25920
                    Height =195
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =3000
                    LayoutCachedWidth =25920
                    LayoutCachedHeight =3195
                    TabIndex =26
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =13035
                    Top =1200
                    Width =2280
                    Height =1080
                    BorderColor =10855845
                    Name ="Box179"
                    LayoutCachedLeft =13035
                    LayoutCachedTop =1200
                    LayoutCachedWidth =15315
                    LayoutCachedHeight =2280
                End
                Begin Line
                    OverlapFlags =93
                    Left =645
                    Top =885
                    Width =23400
                    BorderColor =12349952
                    Name ="Line88"
                    LayoutCachedLeft =645
                    LayoutCachedTop =885
                    LayoutCachedWidth =24045
                    LayoutCachedHeight =885
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =6540
                    Top =2796
                    Width =708
                    Height =288
                    FontSize =10
                    FontWeight =700
                    ForeColor =13158
                    Name ="Label6"
                    Caption ="Asset"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =6540
                    LayoutCachedTop =2796
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =3084
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =2184
                    Top =2796
                    Width =4320
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Unit Description"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2184
                    LayoutCachedTop =2796
                    LayoutCachedWidth =6504
                    LayoutCachedHeight =3084
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =19800
                    Top =2796
                    Width =1452
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label10"
                    Caption ="City"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19800
                    LayoutCachedTop =2796
                    LayoutCachedWidth =21252
                    LayoutCachedHeight =3084
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =132
                    Top =2796
                    Width =2016
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label18"
                    Caption ="Account Name"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =132
                    LayoutCachedTop =2796
                    LayoutCachedWidth =2148
                    LayoutCachedHeight =3084
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =186
                    TextAlign =2
                    Left =11244
                    Top =2796
                    Width =540
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label21"
                    Caption ="STS"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11244
                    LayoutCachedTop =2796
                    LayoutCachedWidth =11784
                    LayoutCachedHeight =3084
                    ColumnStart =5
                    ColumnEnd =5
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
                    ColumnOrder =18
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
                    ColumnOrder =19
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
                    TextAlign =2
                    Left =21288
                    Top =2796
                    Width =576
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label61"
                    Caption ="State"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =21288
                    LayoutCachedTop =2796
                    LayoutCachedWidth =21864
                    LayoutCachedHeight =3084
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =120
                    Top =960
                    Width =2016
                    Height =1800
                    ColumnOrder =12
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    BoundColumn =1
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstSourceCompany"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SourceCompany, ClientId From RemarketingReportAllUnits WHERE [UNITId] is "
                        "not null  AND ([ClientId]=4046)   GROUP BY SourceCompany, ClientId ORDER BY Sour"
                        "ceCompany;"
                    ColumnWidths ="2016;0"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =120
                    LayoutCachedTop =960
                    LayoutCachedWidth =2136
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
                    Left =2175
                    Top =960
                    Width =720
                    Height =1800
                    ColumnOrder =13
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstModelYear"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitYr From RemarketingReportAllUnits WHERE [UNITId] is not null  AND ([C"
                        "lientId]=4046)   GROUP BY UnitYr ORDER BY UnitYr desc;"
                    ColumnWidths ="1440"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =2175
                    LayoutCachedTop =960
                    LayoutCachedWidth =2895
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
                    Left =2895
                    Top =960
                    Width =1170
                    Height =1800
                    ColumnOrder =14
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstUnitType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitType From RemarketingReportAllUnits WHERE [UNITId] is not null  AND ("
                        "[ClientId]=4046)   GROUP BY UnitType ORDER BY UnitType;"
                    ColumnWidths ="1296"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =2895
                    LayoutCachedTop =960
                    LayoutCachedWidth =4065
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
                    Left =4065
                    Top =960
                    Width =1380
                    Height =1800
                    ColumnOrder =15
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstMake"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitMake From RemarketingReportAllUnits WHERE [UNITId] is not null  AND ("
                        "[ClientId]=4046)   GROUP BY UnitMake ORDER BY UnitMake;"
                    ColumnWidths ="1296"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =4065
                    LayoutCachedTop =960
                    LayoutCachedWidth =5445
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
                    Visible = NotDefault
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =15600
                    Top =1440
                    Width =540
                    Height =600
                    ColumnOrder =16
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    BoundColumn =1
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstBuyer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BuyerName,BuyerId From RemarketingReportAllUnits WHERE [UNITId] is not nu"
                        "ll  AND ([ClientId]=4046)   GROUP BY BuyerName, BuyerId ORDER BY BuyerName;"
                    ColumnWidths ="144;0"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =15600
                    LayoutCachedTop =1440
                    LayoutCachedWidth =16140
                    LayoutCachedHeight =2040
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
                    OverlapFlags =247
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =10020
                    Top =1020
                    Width =1742
                    Height =1800
                    ColumnOrder =17
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstStatus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitStatus.UnitStatus, [UnitStatusDesc] & \" - \" & [UnitStatus] A"
                        "S [Desc] FROM LookUp_UnitStatus WHERE UnitStatus in ( SELECT UnitStatus FROM Rem"
                        "arketingReportAllUnits GROUP BY UnitStatus ) ORDER BY LookUp_UnitStatus.UnitStat"
                        "us;"
                    ColumnWidths ="0;944"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =10020
                    LayoutCachedTop =1020
                    LayoutCachedWidth =11762
                    LayoutCachedHeight =2820
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
                    Left =7980
                    Top =1080
                    Width =1553
                    Height =1140
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    Name ="cmdSetFilters"
                    Caption ="After Setting Filters Click HERE\015\012Run the Results"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7980
                    LayoutCachedTop =1080
                    LayoutCachedWidth =9533
                    LayoutCachedHeight =2220
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
                    Left =120
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearClients"
                    Caption ="□ Clear"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =120
                    LayoutCachedTop =780
                    LayoutCachedWidth =660
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =2175
                    Top =750
                    Width =144
                    Height =225
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearYear"
                    Caption ="□"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =2175
                    LayoutCachedTop =750
                    LayoutCachedWidth =2319
                    LayoutCachedHeight =975
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =2985
                    Top =750
                    Width =144
                    Height =225
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearUnitType"
                    Caption ="□"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =2985
                    LayoutCachedTop =750
                    LayoutCachedWidth =3129
                    LayoutCachedHeight =975
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =4140
                    Top =750
                    Width =144
                    Height =225
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearMake"
                    Caption ="□"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =4140
                    LayoutCachedTop =750
                    LayoutCachedWidth =4284
                    LayoutCachedHeight =975
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =10020
                    Top =780
                    Width =144
                    Height =225
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearStatus"
                    Caption ="□"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =10020
                    LayoutCachedTop =780
                    LayoutCachedWidth =10164
                    LayoutCachedHeight =1005
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =15600
                    Top =1260
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearBuyer"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =15600
                    LayoutCachedTop =1260
                    LayoutCachedWidth =16140
                    LayoutCachedHeight =1485
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =7980
                    Top =780
                    Width =1553
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="ClearAndRun"
                    Caption ="Clear Filters and Run"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7980
                    LayoutCachedTop =780
                    LayoutCachedWidth =9533
                    LayoutCachedHeight =1020
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
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =1
                    Left =12930
                    Top =780
                    Width =2370
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearAll"
                    Caption ="◄ Click to Clear All Selections  ►"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =12930
                    LayoutCachedTop =780
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =13620
                    Top =2796
                    Width =936
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label121"
                    Caption ="Odometer"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13620
                    LayoutCachedTop =2796
                    LayoutCachedWidth =14556
                    LayoutCachedHeight =3084
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Image
                    BackStyle =1
                    PictureType =2
                    Left =12750
                    Top =1560
                    Width =480
                    Height =420
                    Name ="Image161"
                    Picture ="1_Excel-365-Green"

                    LayoutCachedLeft =12750
                    LayoutCachedTop =1560
                    LayoutCachedWidth =13230
                    LayoutCachedHeight =1980
                    TabIndex =25
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =21975
                    Top =2280
                    Width =2175
                    Height =270
                    FontSize =8
                    Name ="Label162"
                    Caption ="┌───────  Pricing ───────┐"
                    LayoutCachedLeft =21975
                    LayoutCachedTop =2280
                    LayoutCachedWidth =24150
                    LayoutCachedHeight =2550
                    ForeThemeColorIndex =2
                End
                Begin Image
                    BackStyle =1
                    SizeMode =1
                    PictureType =2
                    Left =24300
                    Top =1080
                    Width =774
                    Height =954
                    Name ="Auto_Logo0"
                    Picture ="ATLAASGlobeLogo"

                    LayoutCachedLeft =24300
                    LayoutCachedTop =1080
                    LayoutCachedWidth =25074
                    LayoutCachedHeight =2034
                    TabIndex =24
                End
                Begin ListBox
                    OverlapFlags =215
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =13275
                    Top =1200
                    Width =2040
                    Height =1080
                    ColumnOrder =11
                    FontSize =8
                    FontWeight =700
                    TabIndex =14
                    BackColor =14277081
                    ForeColor =0
                    BorderColor =4210752
                    Name ="ReportsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LaunchReportsList.ReportPublicName, LaunchReportsList.ReportDesc, LaunchR"
                        "eportsList.Command, LaunchReportsList.FormCalled, LaunchReportsList.SecurityLeve"
                        "l, LaunchReportsList.Tags, LaunchReportsList.FormGroup FROM LaunchReportsList WH"
                        "ERE (((LaunchReportsList.SecurityLevel)<=1) AND ((LaunchReportsList.FormGroup)=\""
                        "RmktDashBoard\")) ORDER BY LaunchReportsList.ReportPublicName;"
                    ColumnWidths ="2016;0;0;0"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="ReportList"
                    ControlTipText ="Sold Report Report of the currently filtered Units"
                    GridlineColor =0

                    LayoutCachedLeft =13275
                    LayoutCachedTop =1200
                    LayoutCachedWidth =15315
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
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =21915
                    Top =960
                    Width =2220
                    Height =1320
                    ColumnOrder =10
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    ForeColor =0
                    BorderColor =4210752
                    Name ="TargetValueList"
                    RowSourceType ="Value List"
                    RowSource ="1;\"Target Set\";2;\"Target Not Set\";3;\"FMV Set\";4;\"FMV Not Set\";7;\"Has Pe"
                        "nalty\";8;\"No Penalty\""
                    ColumnWidths ="0;944"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"
                    ControlTipText ="Select the type of Unit Valuation to filter by, Results will run on click, singl"
                        "e selection"
                    GridlineColor =0

                    LayoutCachedLeft =21915
                    LayoutCachedTop =960
                    LayoutCachedWidth =24135
                    LayoutCachedHeight =2280
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
                    TextAlign =2
                    Left =21915
                    Top =750
                    Width =144
                    Height =225
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearTarget"
                    Caption ="□"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =21915
                    LayoutCachedTop =750
                    LayoutCachedWidth =22059
                    LayoutCachedHeight =975
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =21240
                    Top =960
                    Width =585
                    Height =1800
                    ColumnOrder =8
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    ForeColor =0
                    BorderColor =4210752
                    Name ="StateList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RemarketingReportAllUnits.PState FROM RemarketingReportAllUnits WHERE ((("
                        "RemarketingReportAllUnits.[UNITId]) Is Not Null) AND ((RemarketingReportAllUnits"
                        ".[ClientId])=4046)) GROUP BY RemarketingReportAllUnits.PState ORDER BY Remarketi"
                        "ngReportAllUnits.PState;"
                    ColumnWidths ="562"
                    Tag ="Group1"
                    ControlTipText ="Select the States to filter by"
                    GridlineColor =0

                    LayoutCachedLeft =21240
                    LayoutCachedTop =960
                    LayoutCachedWidth =21825
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
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =19755
                    Top =960
                    Height =1800
                    ColumnOrder =9
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    ForeColor =0
                    BorderColor =4210752
                    Name ="CityList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RemarketingReportAllUnits.PCity FROM RemarketingReportAllUnits WHERE (((R"
                        "emarketingReportAllUnits.[UNITId]) Is Not Null) AND ((RemarketingReportAllUnits."
                        "[ClientId])=4046)) GROUP BY RemarketingReportAllUnits.PCity, RemarketingReportAl"
                        "lUnits.PCity ORDER BY RemarketingReportAllUnits.PCity;"
                    ColumnWidths ="1440"
                    Tag ="Group1"
                    ControlTipText ="Select the City to filter by"
                    GridlineColor =0

                    LayoutCachedLeft =19755
                    LayoutCachedTop =960
                    LayoutCachedWidth =21195
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
                    TextAlign =2
                    Left =21240
                    Top =750
                    Width =144
                    Height =225
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearState"
                    Caption ="□"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =21240
                    LayoutCachedTop =750
                    LayoutCachedWidth =21384
                    LayoutCachedHeight =975
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =19755
                    Top =750
                    Width =144
                    Height =225
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearCity"
                    Caption ="□"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =19755
                    LayoutCachedTop =750
                    LayoutCachedWidth =19899
                    LayoutCachedHeight =975
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11833
                    Top =240
                    Width =689
                    Height =300
                    ColumnOrder =7
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="UnitsShown"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =11833
                    LayoutCachedTop =240
                    LayoutCachedWidth =12522
                    LayoutCachedHeight =540
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                    Begin
                        Begin Label
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
                            Caption ="Units"
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
                    Name ="Label31"
                    Caption ="Inventory Dashboard"
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
                    Caption ="Remarketing"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =150
                    LayoutCachedWidth =1410
                    LayoutCachedHeight =240
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =14556
                    Top =2796
                    Width =1128
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label306"
                    Caption ="Parked Date"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14556
                    LayoutCachedTop =2796
                    LayoutCachedWidth =15684
                    LayoutCachedHeight =3084
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =15684
                    Top =2796
                    Width =1116
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label310"
                    Caption ="Inspected"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15684
                    LayoutCachedTop =2796
                    LayoutCachedWidth =16800
                    LayoutCachedHeight =3084
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =3285
                    Top =90
                    Width =480
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label421"
                    Caption ="Sales"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =3285
                    LayoutCachedTop =90
                    LayoutCachedWidth =3765
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =23085
                    Top =2460
                    Width =1065
                    Height =270
                    ColumnOrder =6
                    FontSize =9
                    TabIndex =17
                    Name ="TargetLine"
                    ControlSource ="=Sum([TargetSalePrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =23085
                    LayoutCachedTop =2460
                    LayoutCachedWidth =24150
                    LayoutCachedHeight =2730
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =21899
                    Top =2460
                    Width =1126
                    Height =270
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =16
                    Name ="FMVLine"
                    ControlSource ="=Sum([FMV])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =21899
                    LayoutCachedTop =2460
                    LayoutCachedWidth =23025
                    LayoutCachedHeight =2730
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =7284
                    Top =2796
                    Width =2748
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label480"
                    Caption ="Lease Sch|Group"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7284
                    LayoutCachedTop =2796
                    LayoutCachedWidth =10032
                    LayoutCachedHeight =3084
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =1
                    Left =10032
                    Top =2796
                    Width =1200
                    Height =288
                    FontSize =10
                    FontWeight =700
                    ForeColor =13158
                    Name ="Label486"
                    Caption =" Unit Num"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10032
                    LayoutCachedTop =2796
                    LayoutCachedWidth =11232
                    LayoutCachedHeight =3084
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =21924
                    Top =2796
                    Width =1152
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label526"
                    Caption ="FMV"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =21924
                    LayoutCachedTop =2796
                    LayoutCachedWidth =23076
                    LayoutCachedHeight =3084
                    ColumnStart =14
                    ColumnEnd =14
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
                    Left =5445
                    Top =960
                    Width =1575
                    Height =1800
                    ColumnOrder =4
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstModel"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitMake From RemarketingReportAllUnits WHERE [UNITId] is not null  AND ("
                        "[ClientId]=4046)   GROUP BY UnitMake ORDER BY UnitMake;"
                    ColumnWidths ="1296"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =5445
                    LayoutCachedTop =960
                    LayoutCachedWidth =7020
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
                    TextAlign =2
                    Left =5520
                    Top =750
                    Width =144
                    Height =225
                    FontSize =10
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearModel"
                    Caption ="□"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =5520
                    LayoutCachedTop =750
                    LayoutCachedWidth =5664
                    LayoutCachedHeight =975
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =23136
                    Top =2796
                    Width =1068
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label14"
                    Caption ="Target Price"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =23136
                    LayoutCachedTop =2796
                    LayoutCachedWidth =24204
                    LayoutCachedHeight =3084
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
                    Left =17328
                    Top =2796
                    Width =2472
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label239"
                    Caption ="Address"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17328
                    LayoutCachedTop =2796
                    LayoutCachedWidth =19800
                    LayoutCachedHeight =3084
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
                    Left =16836
                    Top =2796
                    Width =456
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label292"
                    Caption ="Web"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16836
                    LayoutCachedTop =2796
                    LayoutCachedWidth =17292
                    LayoutCachedHeight =3084
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =11784
                    Top =2796
                    Width =1836
                    Height =288
                    FontSize =10
                    FontWeight =700
                    Name ="Label434"
                    Caption ="VIN"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11784
                    LayoutCachedTop =2796
                    LayoutCachedWidth =13620
                    LayoutCachedHeight =3084
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =14520
                    Top =2580
                    Width =1155
                    Height =285
                    FontSize =9
                    Name ="Label647"
                    Caption ="(Projected)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14520
                    LayoutCachedTop =2580
                    LayoutCachedWidth =15675
                    LayoutCachedHeight =2865
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =24300
                    Top =2790
                    Width =1073
                    Height =237
                    FontSize =8
                    FontWeight =700
                    TabIndex =18
                    ForeColor =0
                    Name ="FindAllABuyer"
                    Caption ="Find  ALL"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =24300
                    LayoutCachedTop =2790
                    LayoutCachedWidth =25373
                    LayoutCachedHeight =3027
                    ColumnStart =21
                    ColumnEnd =21
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =75.0
                    BorderColor =14211288
                    PressedColor =9211020
                    HoverForeColor =0
                    PressedForeColor =0
                    QuickStyle =32
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =20700
                    Top =120
                    Width =1800
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =19
                    ForeColor =8210719
                    Name ="SPDashboard"
                    Caption ="Salesperson  Dashboard"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =20700
                    LayoutCachedTop =120
                    LayoutCachedWidth =22500
                    LayoutCachedHeight =360
                    Alignment =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    HoverColor =13017476
                    PressedColor =14270637
                    PressedForeThemeColorIndex =0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin ComboBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =24300
                    Top =2460
                    Width =1073
                    Height =315
                    ColumnOrder =0
                    FontSize =8
                    FontWeight =700
                    TabIndex =20
                    BorderColor =4210752
                    Name ="Combo678"
                    RowSourceType ="Table/Query"

                    LayoutCachedLeft =24300
                    LayoutCachedTop =2460
                    LayoutCachedWidth =25373
                    LayoutCachedHeight =2775
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =24300
                            Top =2160
                            Width =1073
                            Height =255
                            FontSize =10
                            FontWeight =700
                            Name ="Label679"
                            Caption ="Range"
                            LayoutCachedLeft =24300
                            LayoutCachedTop =2160
                            LayoutCachedWidth =25373
                            LayoutCachedHeight =2415
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =1560
                    Top =780
                    Width =465
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    Name ="SetCompany"
                    Caption ="■ Set"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to SET filter for the selection below"
                    LayoutCachedLeft =1560
                    LayoutCachedTop =780
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =1005
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =2640
                    Top =750
                    Width =144
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    Name ="SetYear"
                    Caption ="■"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to SET filter for the selection below"
                    LayoutCachedLeft =2640
                    LayoutCachedTop =750
                    LayoutCachedWidth =2784
                    LayoutCachedHeight =975
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =3795
                    Top =750
                    Width =144
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    Name ="SetType"
                    Caption ="■"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to SET filter for the selection below"
                    LayoutCachedLeft =3795
                    LayoutCachedTop =750
                    LayoutCachedWidth =3939
                    LayoutCachedHeight =975
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =6780
                    Top =750
                    Width =144
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    Name ="SetModel"
                    Caption ="■"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to SET filter for the selection below"
                    LayoutCachedLeft =6780
                    LayoutCachedTop =750
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =975
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =5190
                    Top =750
                    Width =144
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    Name ="SetMake"
                    Caption ="■"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to SET filter for the selection below"
                    LayoutCachedLeft =5190
                    LayoutCachedTop =750
                    LayoutCachedWidth =5334
                    LayoutCachedHeight =975
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =12405
                    Top =750
                    Width =144
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    Name ="SetStatus"
                    Caption ="■"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to SET filter for the selection below"
                    LayoutCachedLeft =12405
                    LayoutCachedTop =750
                    LayoutCachedWidth =12549
                    LayoutCachedHeight =975
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =20940
                    Top =750
                    Width =144
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    Name ="SetCity"
                    Caption ="■"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to SET filter for the selection below"
                    LayoutCachedLeft =20940
                    LayoutCachedTop =750
                    LayoutCachedWidth =21084
                    LayoutCachedHeight =975
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =23925
                    Top =750
                    Width =144
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    Name ="SetTarget"
                    Caption ="■"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to SET filter for the selection below"
                    LayoutCachedLeft =23925
                    LayoutCachedTop =750
                    LayoutCachedWidth =24069
                    LayoutCachedHeight =975
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =21645
                    Top =750
                    Width =144
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    Name ="SetState"
                    Caption ="■"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to SET filter for the selection below"
                    LayoutCachedLeft =21645
                    LayoutCachedTop =750
                    LayoutCachedWidth =21789
                    LayoutCachedHeight =975
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7980
                    Top =2304
                    Width =1553
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =21
                    Name ="ExportToExcel"
                    Caption ="Export To Excel"
                    OnClick ="[Event Procedure]"
                    Tag ="UAT"
                    ControlTipText ="Export what is displayed to Excel"
                    LeftPadding =90
                    RightPadding =105
                    BottomPadding =165
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000375c18ef375c18ff375c18ff ,
                        0x375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ef000000000000000000000000375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff0000000000000000000000001b2e0cff1b2e0cff1b2e0cff ,
                        0x1b2e0cff1b2e0cff1b2e0cff1b2e0cff294512ff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff417c10ef417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff3f7710ff1b2e0cff375c18ff375c18ff375c18ff ,
                        0x375c18ff375c18ff417c10ff417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ffacc696ffffffffff598c2eff719d4cff ,
                        0xffffffffa0be88ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff4d841fffe7efe1ffd0dec3ffd0dec3ff ,
                        0xe7efe1ff4d841fff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff417c10ff7ca55bffffffffffffffffff ,
                        0x7ca55bff417c10ff417c10ff417c10ff203e08ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff417c10ff719d4cffffffffffffffffff ,
                        0x719d4cff417c10ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff417c10ffe7efe1ffe7efe1ffd0dec3ff ,
                        0xe7efe1ff417c10ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff94b579ffffffffff88ad6aff65953dff ,
                        0xffffffffa0be88ff417c10ff417c10ff335110ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ff417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff417c10ff365611ff66a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff417c10ef417c10ff417c10ff417c10ff417c10ff417c10ff ,
                        0x417c10ff417c10ff417c10ff437e11ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff00000000000000000000000066a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff00000000000000000000000066a321ff66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ff00000000000000000000000066a321ef66a321ff66a321ff ,
                        0x66a321ff66a321ff66a321ff66a321ff81c433ff81c433ff81c433ff81c433ff ,
                        0x81c433ff81c433ef
                    End

                    LayoutCachedLeft =7980
                    LayoutCachedTop =2304
                    LayoutCachedWidth =9533
                    LayoutCachedHeight =2544
                    PictureCaptionArrangement =5
                    ColumnStart =20
                    ColumnEnd =20
                    ForeThemeColorIndex =2
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =223
                    PictureType =1
                    Left =23775
                    Top =120
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =22
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =23775
                    LayoutCachedTop =120
                    LayoutCachedWidth =24525
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
                Begin CommandButton
                    OverlapFlags =215
                    Left =23220
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =23
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
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

                    LayoutCachedLeft =23220
                    LayoutCachedTop =120
                    LayoutCachedWidth =23775
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
            End
        End
        Begin Section
            Height =328
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =21288
                    Top =36
                    Width =576
                    Height =252
                    FontSize =10
                    TabIndex =16
                    BorderColor =0
                    Name ="LState"
                    ControlSource ="PState"
                    GroupTable =1
                    ShowDatePicker =0

                    LayoutCachedLeft =21288
                    LayoutCachedTop =36
                    LayoutCachedWidth =21864
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6540
                    Top =36
                    Width =708
                    Height =252
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =5
                    BorderColor =0
                    ForeColor =4754549
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =6540
                    LayoutCachedTop =36
                    LayoutCachedWidth =7248
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2184
                    Top =36
                    Width =4320
                    Height =252
                    ColumnOrder =14
                    FontSize =10
                    TabIndex =4
                    BorderColor =0
                    Name ="UnitDesc"
                    ControlSource ="=\"<B>\" & [UnitYr] & \"</B> \" & [UnitType] & \" \" & [UnitMake] & \" \" & [Uni"
                        "tModel] & IIf(IsNull([UnitConfiguration]),Null,\" \" & [UnitConfiguration]) & II"
                        "f(IsNull([UnitUOM]),Null,\" \" & [UnitUOM])"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    TextFormat =1
                    ShowDatePicker =0

                    LayoutCachedLeft =2184
                    LayoutCachedTop =36
                    LayoutCachedWidth =6504
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14556
                    Top =36
                    Width =1128
                    Height =252
                    FontSize =10
                    TabIndex =11
                    BorderColor =0
                    Name ="FMVBase"
                    ControlSource ="ParkDateProjected"
                    Format ="Short Date"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =14556
                    LayoutCachedTop =36
                    LayoutCachedWidth =15684
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =1
                    IMESentenceMode =3
                    Left =19800
                    Top =36
                    Width =1452
                    Height =252
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =15
                    BorderColor =0
                    Name ="LocCity"
                    ControlSource ="PCity"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =19800
                    LayoutCachedTop =36
                    LayoutCachedWidth =21252
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
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
                    Left =132
                    Top =36
                    Width =2016
                    Height =252
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =3
                    BorderColor =14211288
                    ForeColor =10040879
                    Name ="SourceCompany"
                    ControlSource ="SourceCompany"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =132
                    LayoutCachedTop =36
                    LayoutCachedWidth =2148
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =15
                    Top =30
                    Width =45
                    Height =210
                    FontSize =8
                    Name ="ClientID"
                    ControlSource ="ClientID"

                    LayoutCachedLeft =15
                    LayoutCachedTop =30
                    LayoutCachedWidth =60
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =75
                    Top =30
                    Width =45
                    Height =210
                    FontSize =8
                    TabIndex =1
                    Name ="BuyerID"
                    ControlSource ="BuyerID"

                    LayoutCachedLeft =75
                    LayoutCachedTop =30
                    LayoutCachedWidth =120
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11244
                    Top =36
                    Width =540
                    Height =252
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =8
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =11244
                    LayoutCachedTop =36
                    LayoutCachedWidth =11784
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
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
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13620
                    Top =36
                    Width =936
                    Height =252
                    FontSize =10
                    TabIndex =10
                    BorderColor =0
                    Name ="Odometer"
                    ControlSource ="Odometer"
                    Format ="Standard"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =13620
                    LayoutCachedTop =36
                    LayoutCachedWidth =14556
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15684
                    Top =36
                    Width =1116
                    Height =252
                    FontSize =10
                    TabIndex =12
                    BorderColor =0
                    Name ="Text308"
                    ControlSource ="InspectedDate"
                    Format ="Short Date"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =15684
                    LayoutCachedTop =36
                    LayoutCachedWidth =16800
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Top =270
                    Width =25920
                    Height =58
                    TabIndex =2
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End

                    LayoutCachedTop =270
                    LayoutCachedWidth =25920
                    LayoutCachedHeight =328
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7284
                    Top =36
                    Width =2748
                    Height =252
                    FontSize =8
                    TabIndex =6
                    BorderColor =0
                    Name ="Text478"
                    ControlSource ="Sch|Grp"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =7284
                    LayoutCachedTop =36
                    LayoutCachedWidth =10032
                    LayoutCachedHeight =288
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
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10032
                    Top =36
                    Width =1200
                    Height =252
                    FontSize =10
                    TabIndex =7
                    BorderColor =0
                    Name ="Text487"
                    ControlSource ="UnitNum"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =10032
                    LayoutCachedTop =36
                    LayoutCachedWidth =11232
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =21924
                    Top =36
                    Width =1152
                    Height =252
                    FontSize =10
                    TabIndex =17
                    BorderColor =0
                    Name ="Text520"
                    ControlSource ="FMV"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    ShowDatePicker =0

                    LayoutCachedLeft =21924
                    LayoutCachedTop =36
                    LayoutCachedWidth =23076
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =23136
                    Top =36
                    Width =1068
                    Height =252
                    FontSize =10
                    TabIndex =18
                    BorderColor =0
                    Name ="TargetSalePrice"
                    ControlSource ="TargetSalePrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    ShowDatePicker =0

                    LayoutCachedLeft =23136
                    LayoutCachedTop =36
                    LayoutCachedWidth =24204
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =17328
                    Top =36
                    Width =2472
                    Height =252
                    FontSize =8
                    TabIndex =14
                    BorderColor =0
                    Name ="LocationAddress"
                    ControlSource ="ParkedAddress"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =17328
                    LayoutCachedTop =36
                    LayoutCachedWidth =19800
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =16836
                    Top =36
                    Width =456
                    Height =252
                    FontSize =10
                    TabIndex =13
                    BorderColor =0
                    Name ="WebLink"
                    ControlSource ="InspectionLinkWeb"
                    GroupTable =1
                    ShowDatePicker =0

                    LayoutCachedLeft =16836
                    LayoutCachedTop =36
                    LayoutCachedWidth =17292
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =10
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
                    Left =11784
                    Top =36
                    Width =1836
                    Height =252
                    FontSize =9
                    TabIndex =9
                    BorderColor =0
                    Name ="Text432"
                    ControlSource ="UnitVIN"
                    Format ="Standard"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =11784
                    LayoutCachedTop =36
                    LayoutCachedWidth =13620
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
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
' See "RemarketingDashboardINVSales.cls"
