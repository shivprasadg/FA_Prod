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
    Width =28440
    DatasheetFontHeight =11
    ItemSuffix =710
    Left =26610
    Top =5235
    Right =-10486
    Bottom =23535
    ShortcutMenuBar ="SortClip"
    OrderBy ="[RemarketingDashboardAdmin].[Odometer] DESC"
    RecSrcDt = Begin
        0x3781fc0730bce440
    End
    RecordSource ="SELECT * FROM RemarketingReportAllUnits WHERE [UNITId] is not null  And InvoiceP"
        "aid Between #1/1/2014# AND #12/31/2014# Order By SourceCompany;"
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
            Height =3195
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Width =28440
                    Height =540
                    BorderColor =14211288
                    Name ="Box49"
                    HorizontalAnchor =2
                    LayoutCachedWidth =28440
                    LayoutCachedHeight =540
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =3000
                    Width =28440
                    Height =195
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"

                    LayoutCachedTop =3000
                    LayoutCachedWidth =28440
                    LayoutCachedHeight =3195
                    TabIndex =50
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =12255
                    Top =1260
                    Width =2280
                    Height =1275
                    BorderColor =10855845
                    Name ="Box179"
                    LayoutCachedLeft =12255
                    LayoutCachedTop =1260
                    LayoutCachedWidth =14535
                    LayoutCachedHeight =2535
                End
                Begin Line
                    OverlapFlags =93
                    Left =658
                    Top =885
                    Width =25229
                    BorderColor =12349952
                    Name ="Line88"
                    LayoutCachedLeft =658
                    LayoutCachedTop =885
                    LayoutCachedWidth =25887
                    LayoutCachedHeight =885
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =6060
                    Top =2820
                    Width =705
                    Height =285
                    FontSize =10
                    FontWeight =700
                    ForeColor =13158
                    Name ="Label6"
                    Caption ="Asset"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6060
                    LayoutCachedTop =2820
                    LayoutCachedWidth =6765
                    LayoutCachedHeight =3105
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
                    TextAlign =2
                    Left =15585
                    Top =2820
                    Width =1455
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label10"
                    Caption ="City"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15585
                    LayoutCachedTop =2820
                    LayoutCachedWidth =17040
                    LayoutCachedHeight =3105
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =27270
                    Top =2820
                    Width =975
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label15"
                    Caption ="PayOff Back"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =27270
                    LayoutCachedTop =2820
                    LayoutCachedWidth =28245
                    LayoutCachedHeight =3105
                    ColumnStart =23
                    ColumnEnd =23
                    LayoutGroup =1
                    BackThemeColorIndex =7
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =105
                    Top =2820
                    Width =1851
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label18"
                    Caption ="Account Name"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =105
                    LayoutCachedTop =2820
                    LayoutCachedWidth =1956
                    LayoutCachedHeight =3105
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =186
                    TextAlign =2
                    Left =10320
                    Top =2820
                    Width =315
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label21"
                    Caption ="Sts"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10320
                    LayoutCachedTop =2820
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =3105
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
                    ColumnOrder =28
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
                    ColumnOrder =29
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
                    Left =17040
                    Top =2820
                    Width =570
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label61"
                    Caption ="State"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17040
                    LayoutCachedTop =2820
                    LayoutCachedWidth =17610
                    LayoutCachedHeight =3105
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ListBox
                    OverlapFlags =223
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =120
                    Top =960
                    Width =1866
                    Height =1800
                    ColumnOrder =22
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    BoundColumn =1
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstSourceCompany"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SourceCompany, ClientId From RemarketingReportAllUnits WHERE [UNITId] is "
                        "not null  And InvoicePaid Between #1/1/2014# AND #12/31/2014#  GROUP BY SourceCo"
                        "mpany, ClientId ORDER BY SourceCompany;"
                    ColumnWidths ="2016;0"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =120
                    LayoutCachedTop =960
                    LayoutCachedWidth =1986
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
                    OverlapFlags =255
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =1950
                    Top =960
                    Width =720
                    Height =1800
                    ColumnOrder =23
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstModelYear"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitYr From RemarketingReportAllUnits WHERE [UNITId] is not null  And Inv"
                        "oicePaid Between #1/1/2014# AND #12/31/2014#  GROUP BY UnitYr ORDER BY UnitYr de"
                        "sc;"
                    ColumnWidths ="1440"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =1950
                    LayoutCachedTop =960
                    LayoutCachedWidth =2670
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
                    OverlapFlags =255
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =2670
                    Top =960
                    Width =1350
                    Height =1800
                    ColumnOrder =24
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstUnitType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitType From RemarketingReportAllUnits WHERE [UNITId] is not null  And I"
                        "nvoicePaid Between #1/1/2014# AND #12/31/2014#  GROUP BY UnitType ORDER BY UnitT"
                        "ype;"
                    ColumnWidths ="1296"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =2670
                    LayoutCachedTop =960
                    LayoutCachedWidth =4020
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
                    OverlapFlags =255
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =4020
                    Top =960
                    Width =1380
                    Height =1800
                    ColumnOrder =25
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstMake"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitMake From RemarketingReportAllUnits WHERE [UNITId] is not null  And I"
                        "nvoicePaid Between #1/1/2014# AND #12/31/2014#  GROUP BY UnitMake ORDER BY UnitM"
                        "ake;"
                    ColumnWidths ="1296"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =4020
                    LayoutCachedTop =960
                    LayoutCachedWidth =5400
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
                    ColumnCount =2
                    Left =21555
                    Top =960
                    Width =1530
                    Height =1800
                    ColumnOrder =26
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    BoundColumn =1
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstBuyer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BuyerName,BuyerId From RemarketingReportAllUnits WHERE [UNITId] is not nu"
                        "ll  And InvoicePaid Between #1/1/2014# AND #12/31/2014#  GROUP BY BuyerName, Buy"
                        "erId ORDER BY BuyerName;"
                    ColumnWidths ="144;0"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =21555
                    LayoutCachedTop =960
                    LayoutCachedWidth =23085
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
                    OverlapFlags =247
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =10320
                    Top =960
                    Width =1382
                    Height =1800
                    ColumnOrder =27
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstStatus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitStatus.UnitStatus, [UnitStatusDesc] & \" - \" & [UnitStatus] A"
                        "S [Desc] FROM LookUp_UnitStatus WHERE (((LookUp_UnitStatus.[UnitStatus]) In (SEL"
                        "ECT UnitStatus FROM RemarketingReportAllUnits GROUP BY UnitStatus ))) ORDER BY ["
                        "UnitStatusDesc] & \" - \" & [UnitStatus], LookUp_UnitStatus.[UnitStatus];"
                    ColumnWidths ="0;944"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =10320
                    LayoutCachedTop =960
                    LayoutCachedWidth =11702
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
                    OverlapFlags =85
                    Left =7485
                    Top =1080
                    Width =1553
                    Height =1140
                    FontSize =8
                    FontWeight =700
                    TabIndex =16
                    Name ="cmdSetFilters"
                    Caption ="After Setting Filters Click HERE\015\012Run the Results"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7485
                    LayoutCachedTop =1080
                    LayoutCachedWidth =9038
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
                    OverlapFlags =95
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
                    Caption ="Clear▼"
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
                    TextAlign =1
                    Left =1950
                    Top =780
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
                    LayoutCachedLeft =1950
                    LayoutCachedTop =780
                    LayoutCachedWidth =2490
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =2670
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearUnitType"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =2670
                    LayoutCachedTop =780
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =4020
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearMake"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =4020
                    LayoutCachedTop =780
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =10320
                    Top =780
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
                    LayoutCachedLeft =10320
                    LayoutCachedTop =780
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =21555
                    Top =780
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
                    LayoutCachedLeft =21555
                    LayoutCachedTop =780
                    LayoutCachedWidth =22095
                    LayoutCachedHeight =1005
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =7485
                    Top =780
                    Width =1553
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="ClearAndRun"
                    Caption ="Clear Filters and Run"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7485
                    LayoutCachedTop =780
                    LayoutCachedWidth =9038
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
                    Left =12075
                    Top =780
                    Width =2355
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearAll"
                    Caption ="◄ Click to clear All Selections  ►"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =12075
                    LayoutCachedTop =780
                    LayoutCachedWidth =14430
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =12450
                    Top =2820
                    Width =945
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label121"
                    Caption ="Odometer"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12450
                    LayoutCachedTop =2820
                    LayoutCachedWidth =13395
                    LayoutCachedHeight =3105
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Image
                    BackStyle =1
                    PictureType =2
                    Left =11970
                    Top =1710
                    Width =480
                    Height =420
                    Name ="Image161"
                    Picture ="1_Excel-365-Green"

                    LayoutCachedLeft =11970
                    LayoutCachedTop =1710
                    LayoutCachedWidth =12450
                    LayoutCachedHeight =2130
                    TabIndex =49
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =18900
                    Top =2280
                    Width =1725
                    Height =240
                    FontSize =8
                    Name ="Label162"
                    Caption ="┌─────  Pricing ─────┐"
                    LayoutCachedLeft =18900
                    LayoutCachedTop =2280
                    LayoutCachedWidth =20625
                    LayoutCachedHeight =2520
                    ForeThemeColorIndex =2
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =18720
                    Top =960
                    Width =2025
                    Height =1320
                    ColumnOrder =20
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    ForeColor =0
                    BorderColor =4210752
                    Name ="TargetValueList"
                    RowSourceType ="Value List"
                    RowSource ="1;\"Target Set\";2;\"Target Not Set\";3;\"Payoff Set\";4;\"Payoff Not Set\";5;\""
                        "Payoff Back\";6;\"Payoff Not Back\";7;\"Has Penalty\";8;\"No Penalty\""
                    ColumnWidths ="0;944"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"
                    ControlTipText ="Select the type of Unit Valuation to filter by, Results will run on click, singl"
                        "e selection"
                    GridlineColor =0

                    LayoutCachedLeft =18720
                    LayoutCachedTop =960
                    LayoutCachedWidth =20745
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
                    TextAlign =1
                    Left =18720
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearTarget"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =18720
                    LayoutCachedTop =780
                    LayoutCachedWidth =19260
                    LayoutCachedHeight =1005
                End
                Begin ListBox
                    OverlapFlags =223
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =17010
                    Top =960
                    Width =585
                    Height =1800
                    ColumnOrder =18
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    ForeColor =0
                    BorderColor =4210752
                    Name ="StateList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RemarketingReportAllUnits.PState FROM RemarketingReportAllUnits WHERE ((("
                        "RemarketingReportAllUnits.[UNITId]) Is Not Null) AND ((RemarketingReportAllUnits"
                        ".[InvoicePaid]) Between #1/1/2014# And #12/31/2014#)) GROUP BY RemarketingReport"
                        "AllUnits.PState ORDER BY RemarketingReportAllUnits.PState;"
                    ColumnWidths ="562"
                    Tag ="Group1"
                    ControlTipText ="Select the States to filter by"
                    GridlineColor =0

                    LayoutCachedLeft =17010
                    LayoutCachedTop =960
                    LayoutCachedWidth =17595
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
                    Left =15570
                    Top =960
                    Height =1800
                    ColumnOrder =19
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    ForeColor =0
                    BorderColor =4210752
                    Name ="CityList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RemarketingReportAllUnits.PCity FROM RemarketingReportAllUnits WHERE (((R"
                        "emarketingReportAllUnits.[UNITId]) Is Not Null) AND ((RemarketingReportAllUnits."
                        "[InvoicePaid]) Between #1/1/2014# And #12/31/2014#)) GROUP BY RemarketingReportA"
                        "llUnits.PCity ORDER BY RemarketingReportAllUnits.PCity;"
                    ColumnWidths ="1440"
                    Tag ="Group1"
                    ControlTipText ="Select the City to filter by"
                    GridlineColor =0

                    LayoutCachedLeft =15570
                    LayoutCachedTop =960
                    LayoutCachedWidth =17010
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
                    Left =17010
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearState"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =17010
                    LayoutCachedTop =780
                    LayoutCachedWidth =17550
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =15570
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearCity"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =15570
                    LayoutCachedTop =780
                    LayoutCachedWidth =16110
                    LayoutCachedHeight =1005
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11833
                    Top =240
                    Width =689
                    Height =300
                    ColumnOrder =17
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
                    Caption ="Admin Dashboard"
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
                    Left =13905
                    Top =2820
                    Width =825
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label306"
                    Caption ="Parked"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13905
                    LayoutCachedTop =2820
                    LayoutCachedWidth =14730
                    LayoutCachedHeight =3105
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =14730
                    Top =2820
                    Width =855
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label310"
                    Caption ="Inspected"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14730
                    LayoutCachedTop =2820
                    LayoutCachedWidth =15585
                    LayoutCachedHeight =3105
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =13395
                    Top =2820
                    Width =510
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label5"
                    Caption ="Days"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13395
                    LayoutCachedTop =2820
                    LayoutCachedWidth =13905
                    LayoutCachedHeight =3105
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =17610
                    Top =2820
                    Width =1095
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label376"
                    Caption ="Buy Out"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17610
                    LayoutCachedTop =2820
                    LayoutCachedWidth =18705
                    LayoutCachedHeight =3105
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18705
                    Top =2520
                    Width =1101
                    Height =270
                    ColumnOrder =16
                    FontSize =9
                    TabIndex =39
                    Name ="TargetLine"
                    ControlSource ="=Sum([TargetSalePrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =18705
                    LayoutCachedTop =2520
                    LayoutCachedWidth =19806
                    LayoutCachedHeight =2790
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17595
                    Top =2520
                    Width =1101
                    Height =270
                    ColumnOrder =15
                    FontSize =9
                    TabIndex =38
                    Name ="Text422"
                    ControlSource ="=Sum([BuyOut])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =17595
                    LayoutCachedTop =2520
                    LayoutCachedWidth =18696
                    LayoutCachedHeight =2790
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =6765
                    Top =2820
                    Width =2835
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label480"
                    Caption ="Lease Sch |Group"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6765
                    LayoutCachedTop =2820
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =3105
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =9600
                    Top =2820
                    Width =720
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label486"
                    Caption ="Unit#"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9600
                    LayoutCachedTop =2820
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =3105
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =1950
                    Top =2820
                    Width =4110
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label506"
                    Caption ="Model"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1950
                    LayoutCachedTop =2820
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =3105
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =18705
                    Top =2820
                    Width =1095
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label526"
                    Caption ="Target"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18705
                    LayoutCachedTop =2820
                    LayoutCachedWidth =19800
                    LayoutCachedHeight =3105
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ListBox
                    OverlapFlags =247
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =5400
                    Top =960
                    Width =1575
                    Height =1800
                    ColumnOrder =14
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstModel"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitModel From RemarketingReportAllUnits WHERE [UNITId] is not null  And "
                        "InvoicePaid Between #1/1/2014# AND #12/31/2014#  GROUP BY UnitModel ORDER BY Uni"
                        "tModel;"
                    ColumnWidths ="1296"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =5400
                    LayoutCachedTop =960
                    LayoutCachedWidth =6975
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
                    Left =5400
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearModel"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =780
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =13856
                    Top =2625
                    Width =885
                    Height =285
                    FontSize =8
                    Name ="Label647"
                    Caption ="(Projected)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13856
                    LayoutCachedTop =2625
                    LayoutCachedWidth =14741
                    LayoutCachedHeight =2910
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =19800
                    Top =2820
                    Width =1095
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label605"
                    Caption ="Sale Price"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19800
                    LayoutCachedTop =2820
                    LayoutCachedWidth =20895
                    LayoutCachedHeight =3105
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
                    Left =20895
                    Top =2820
                    Width =660
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label608"
                    Caption ="Bank"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =20895
                    LayoutCachedTop =2820
                    LayoutCachedWidth =21555
                    LayoutCachedHeight =3105
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =21555
                    Top =2820
                    Width =1515
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label616"
                    Caption ="Buyer"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =21555
                    LayoutCachedTop =2820
                    LayoutCachedWidth =23070
                    LayoutCachedHeight =3105
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10635
                    Top =2820
                    Width =1815
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label624"
                    Caption ="VIN"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10635
                    LayoutCachedTop =2820
                    LayoutCachedWidth =12450
                    LayoutCachedHeight =3105
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =26295
                    Top =2820
                    Width =975
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label35"
                    Caption ="PayOff Sent"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =26295
                    LayoutCachedTop =2820
                    LayoutCachedWidth =27270
                    LayoutCachedHeight =3105
                    ColumnStart =22
                    ColumnEnd =22
                    LayoutGroup =1
                    BackThemeColorIndex =7
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =24345
                    Top =2820
                    Width =975
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label36"
                    Caption ="BuyOut"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =24345
                    LayoutCachedTop =2820
                    LayoutCachedWidth =25320
                    LayoutCachedHeight =3105
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =23370
                    Top =2820
                    Width =975
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label37"
                    Caption ="SUF Recv"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =23370
                    LayoutCachedTop =2820
                    LayoutCachedWidth =24345
                    LayoutCachedHeight =3105
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =25320
                    Top =2820
                    Width =975
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label659"
                    Caption ="Invoice Paid"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =25320
                    LayoutCachedTop =2820
                    LayoutCachedWidth =26295
                    LayoutCachedHeight =3105
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Line
                    OverlapFlags =93
                    Left =25887
                    Top =900
                    Width =0
                    Height =792
                    BorderColor =12349952
                    Name ="Line203"
                    LayoutCachedLeft =25887
                    LayoutCachedTop =900
                    LayoutCachedWidth =25887
                    LayoutCachedHeight =1692
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =23933
                    Top =1680
                    Width =3915
                    Height =240
                    FontSize =8
                    ForeColor =12349952
                    Name ="Label44"
                    Caption ="┌───vv─────── Quick Set Date Rages ────────────┐"
                    LayoutCachedLeft =23933
                    LayoutCachedTop =1680
                    LayoutCachedWidth =27848
                    LayoutCachedHeight =1920
                    BackThemeColorIndex =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =26295
                    Top =2490
                    Width =975
                    Height =270
                    ColumnOrder =9
                    FontSize =8
                    FontWeight =700
                    TabIndex =36
                    BorderColor =4210752
                    Name ="PayOffSentEnd"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =26295
                    LayoutCachedTop =2490
                    LayoutCachedWidth =27270
                    LayoutCachedHeight =2760
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =26295
                    Top =2160
                    Width =975
                    Height =270
                    ColumnOrder =6
                    FontSize =8
                    FontWeight =700
                    TabIndex =32
                    BorderColor =4210752
                    Name ="PayOffSentStart"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =26295
                    LayoutCachedTop =2160
                    LayoutCachedWidth =27270
                    LayoutCachedHeight =2430
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =23370
                    Top =2490
                    Width =975
                    Height =270
                    ColumnOrder =10
                    FontSize =8
                    FontWeight =700
                    TabIndex =34
                    BorderColor =4210752
                    Name ="SUFRecvEnd"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =23370
                    LayoutCachedTop =2490
                    LayoutCachedWidth =24345
                    LayoutCachedHeight =2760
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =23370
                    Top =2160
                    Width =975
                    Height =270
                    ColumnOrder =7
                    FontSize =8
                    FontWeight =700
                    TabIndex =30
                    BorderColor =4210752
                    Name ="SUFRecvStart"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =23370
                    LayoutCachedTop =2160
                    LayoutCachedWidth =24345
                    LayoutCachedHeight =2430
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =24345
                    Top =2490
                    Width =975
                    Height =270
                    ColumnOrder =8
                    FontSize =8
                    FontWeight =700
                    TabIndex =35
                    BorderColor =4210752
                    Name ="BuyOutEnd"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =24345
                    LayoutCachedTop =2490
                    LayoutCachedWidth =25320
                    LayoutCachedHeight =2760
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =24345
                    Top =2160
                    Width =975
                    Height =270
                    ColumnOrder =5
                    FontSize =8
                    FontWeight =700
                    TabIndex =31
                    BorderColor =4210752
                    Name ="BuyOutStart"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =24345
                    LayoutCachedTop =2160
                    LayoutCachedWidth =25320
                    LayoutCachedHeight =2430
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =23460
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =18
                    ForeColor =12349952
                    Name ="SUF0"
                    Caption ="16"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =23460
                    LayoutCachedTop =1920
                    LayoutCachedWidth =23745
                    LayoutCachedHeight =2145
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =4
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =87
                    Left =23760
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =19
                    ForeColor =12349952
                    Name ="SUF1"
                    Caption ="15"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =23760
                    LayoutCachedTop =1920
                    LayoutCachedWidth =24045
                    LayoutCachedHeight =2145
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
                    GroupTable =4
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =87
                    Left =26340
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =24
                    ForeColor =12349952
                    Name ="Sent0"
                    Caption ="16"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =26340
                    LayoutCachedTop =1920
                    LayoutCachedWidth =26625
                    LayoutCachedHeight =2145
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =5
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =87
                    Left =26640
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =25
                    ForeColor =12349952
                    Name ="Sent1"
                    Caption ="15"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =26640
                    LayoutCachedTop =1920
                    LayoutCachedWidth =26925
                    LayoutCachedHeight =2145
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =5
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =87
                    Left =24390
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =21
                    ForeColor =12349952
                    Name ="Buy0"
                    Caption ="16"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =24390
                    LayoutCachedTop =1920
                    LayoutCachedWidth =24675
                    LayoutCachedHeight =2145
                    LayoutGroup =4
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
                    OverlapFlags =87
                    Left =24690
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =22
                    ForeColor =12349952
                    Name ="Buy1"
                    Caption ="15"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =24690
                    LayoutCachedTop =1920
                    LayoutCachedWidth =24975
                    LayoutCachedHeight =2145
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
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
                    OverlapFlags =87
                    Left =24060
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =20
                    ForeColor =12349952
                    Name ="SUF2"
                    Caption ="14"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =24060
                    LayoutCachedTop =1920
                    LayoutCachedWidth =24345
                    LayoutCachedHeight =2145
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
                    GroupTable =4
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =87
                    Left =24990
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =23
                    ForeColor =12349952
                    Name ="Buy2"
                    Caption ="14"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =24990
                    LayoutCachedTop =1920
                    LayoutCachedWidth =25275
                    LayoutCachedHeight =2145
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
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
                    OverlapFlags =87
                    Left =26940
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =26
                    ForeColor =12349952
                    Name ="Sent2"
                    Caption ="14"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =26940
                    LayoutCachedTop =1920
                    LayoutCachedWidth =27225
                    LayoutCachedHeight =2145
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =5
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19800
                    Top =2520
                    Width =1101
                    Height =270
                    ColumnOrder =11
                    FontSize =9
                    TabIndex =40
                    Name ="Text671"
                    ControlSource ="=Sum([UnitSalePrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =19800
                    LayoutCachedTop =2520
                    LayoutCachedWidth =20901
                    LayoutCachedHeight =2790
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =25320
                    Top =2490
                    Width =975
                    Height =270
                    ColumnOrder =12
                    FontSize =8
                    FontWeight =700
                    TabIndex =37
                    BorderColor =4210752
                    Name ="SoldDateEnd"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =25320
                    LayoutCachedTop =2490
                    LayoutCachedWidth =26295
                    LayoutCachedHeight =2760
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =25320
                    Top =2160
                    Width =975
                    Height =270
                    ColumnOrder =13
                    FontSize =8
                    FontWeight =700
                    TabIndex =33
                    BorderColor =4210752
                    Name ="SoldDateStart"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =25320
                    LayoutCachedTop =2160
                    LayoutCachedWidth =26295
                    LayoutCachedHeight =2430
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =87
                    Left =25380
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =27
                    ForeColor =12349952
                    Name ="Sold0"
                    Caption ="16"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =25380
                    LayoutCachedTop =1920
                    LayoutCachedWidth =25665
                    LayoutCachedHeight =2145
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =7
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =87
                    Left =25680
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =28
                    ForeColor =12349952
                    Name ="Sold1"
                    Caption ="15"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =25680
                    LayoutCachedTop =1920
                    LayoutCachedWidth =25965
                    LayoutCachedHeight =2145
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =7
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =87
                    Left =25980
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =29
                    ForeColor =12349952
                    Name ="Sold2"
                    Caption ="14"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =25980
                    LayoutCachedTop =1920
                    LayoutCachedWidth =26265
                    LayoutCachedHeight =2145
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =7
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =26212
                    Top =780
                    Width =2048
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="OpenMassMod"
                    Caption ="Open Mass Mod"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =26212
                    LayoutCachedTop =780
                    LayoutCachedWidth =28260
                    LayoutCachedHeight =1020
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =5
                    HoverColor =0
                    PressedThemeColorIndex =5
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =32
                    QuickStyleMask =-629
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =23070
                    Top =2820
                    Width =300
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label678"
                    Caption ="SP"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =23070
                    LayoutCachedTop =2820
                    LayoutCachedWidth =23370
                    LayoutCachedHeight =3105
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =9600
                    Top =2625
                    Width =720
                    Height =285
                    FontSize =8
                    Name ="Label686"
                    Caption ="Cust"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9600
                    LayoutCachedTop =2625
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =2910
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =20895
                    Top =780
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =14211288
                    ForeColor =12349952
                    Name ="ClearBank"
                    Caption ="Clear▼"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =20895
                    LayoutCachedTop =780
                    LayoutCachedWidth =21435
                    LayoutCachedHeight =1005
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =20895
                    Top =960
                    Width =660
                    Height =1800
                    ColumnOrder =4
                    FontSize =8
                    FontWeight =700
                    TabIndex =14
                    ForeColor =0
                    BorderColor =4210752
                    Name ="AssigneeList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BankPartner FROM RemarketingReportAllUnits WHERE UNITID Is Not Null GROUP"
                        " BY BankPartner;"
                    ColumnWidths ="1440"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =20895
                    LayoutCachedTop =960
                    LayoutCachedWidth =21555
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
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =12495
                    Top =1200
                    Width =2040
                    Height =1395
                    ColumnOrder =21
                    FontSize =8
                    FontWeight =700
                    TabIndex =17
                    BackColor =14277081
                    ForeColor =0
                    BorderColor =4210752
                    Name ="ReportsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LaunchReportsList.ReportPublicName, LaunchReportsList.ReportDesc, LaunchR"
                        "eportsList.Command, LaunchReportsList.FormCalled, LaunchReportsList.SecurityLeve"
                        "l, LaunchReportsList.Tags, LaunchReportsList.FormGroup FROM LaunchReportsList WH"
                        "ERE (((LaunchReportsList.SecurityLevel)<=5) AND ((LaunchReportsList.FormGroup)=\""
                        "RmktDashBoard\")) ORDER BY LaunchReportsList.ListOrder;"
                    ColumnWidths ="2016;0;0;0"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="ReportList"
                    ControlTipText ="Admin - Report of the currently filtered Units"
                    GridlineColor =0

                    LayoutCachedLeft =12495
                    LayoutCachedTop =1200
                    LayoutCachedWidth =14535
                    LayoutCachedHeight =2595
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
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =27270
                    Top =2490
                    Width =975
                    Height =270
                    ColumnOrder =0
                    FontSize =8
                    FontWeight =700
                    TabIndex =41
                    BorderColor =4210752
                    Name ="PayOffBackEnd"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =27270
                    LayoutCachedTop =2490
                    LayoutCachedWidth =28245
                    LayoutCachedHeight =2760
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =27270
                    Top =2160
                    Width =975
                    Height =270
                    ColumnOrder =3
                    FontSize =8
                    FontWeight =700
                    TabIndex =42
                    BorderColor =4210752
                    Name ="PayOffBackStart"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =27270
                    LayoutCachedTop =2160
                    LayoutCachedWidth =28245
                    LayoutCachedHeight =2430
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =87
                    Left =27300
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =43
                    ForeColor =12349952
                    Name ="Back0"
                    Caption ="16"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =27300
                    LayoutCachedTop =1920
                    LayoutCachedWidth =27585
                    LayoutCachedHeight =2145
                    LayoutGroup =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =8
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =87
                    Left =27600
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =44
                    ForeColor =12349952
                    Name ="Back1"
                    Caption ="15"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =27600
                    LayoutCachedTop =1920
                    LayoutCachedWidth =27885
                    LayoutCachedHeight =2145
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =8
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =27900
                    Top =1920
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =45
                    ForeColor =12349952
                    Name ="Back2"
                    Caption ="14"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =27900
                    LayoutCachedTop =1920
                    LayoutCachedWidth =28185
                    LayoutCachedHeight =2145
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =8
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7485
                    Top =2340
                    Width =1553
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =46
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

                    LayoutCachedLeft =7485
                    LayoutCachedTop =2340
                    LayoutCachedWidth =9038
                    LayoutCachedHeight =2580
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
                    Left =24300
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =47
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

                    LayoutCachedLeft =24300
                    LayoutCachedTop =120
                    LayoutCachedWidth =24855
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
                    Left =24867
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =48
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1

                    LayoutCachedLeft =24867
                    LayoutCachedTop =120
                    LayoutCachedWidth =25422
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
            Height =313
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13395
                    Width =510
                    Height =255
                    FontSize =10
                    TabIndex =11
                    BorderColor =4210752
                    Name ="DaysInInventory"
                    ControlSource ="[DaysInInventory]"
                    StatusBarText ="Short Name for Reporting"
                    ControlTipText ="Days after 'Parked' in INV (before Sold Date, if sold). A neg value reflect days"
                        " until parked."
                    ConditionalFormat = Begin
                        0x010000007e000000030000000000000000000000000000000200000001000000 ,
                        0x00000000d7edda00000000000000000005000000080000000100000000000000 ,
                        0xffffcc0000000000060000000b0000000e00000001000000ffffff0099330000 ,
                        0x300000003400350000003400360000003700350000003700360000000000
                    End
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =13395
                    LayoutCachedWidth =13905
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x01000300000000000000000000000100000000000000d7edda00010000003000 ,
                        0x0200000034003500000000000000000000000000000000000000000000000000 ,
                        0x000100000000000000ffffcc0002000000340036000200000037003500000000 ,
                        0x0000000000000000000000000000000000000600000001000000ffffff009933 ,
                        0x00000200000037003600000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1950
                    Width =4110
                    Height =255
                    FontSize =8
                    TabIndex =4
                    BorderColor =0
                    Name ="UnitDesc"
                    ControlSource ="=\"<B>\" & [UnitYr] & \"</B> \" & [UnitType] & \" \" & [UnitMake] & \" \" & [Uni"
                        "tModel]"
                    StatusBarText ="unit make"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1
                    ShowDatePicker =0

                    LayoutCachedLeft =1950
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =255
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
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =17040
                    Width =570
                    Height =255
                    FontSize =9
                    TabIndex =15
                    BorderColor =0
                    Name ="LState"
                    ControlSource ="PState"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =17040
                    LayoutCachedWidth =17610
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6060
                    Width =705
                    Height =255
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =5
                    BorderColor =0
                    ForeColor =4754549
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =6060
                    LayoutCachedWidth =6765
                    LayoutCachedHeight =255
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
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13905
                    Width =825
                    Height =255
                    FontSize =8
                    TabIndex =12
                    BorderColor =0
                    Name ="FMVBase"
                    ControlSource ="ParkDateProjected"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =13905
                    LayoutCachedWidth =14730
                    LayoutCachedHeight =255
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
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15585
                    Width =1455
                    Height =255
                    ColumnOrder =9
                    FontSize =8
                    TabIndex =14
                    BorderColor =0
                    Name ="LocCity"
                    ControlSource ="PCity"
                    Format ="$#,##0.00;($#,##0.00)"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =15585
                    LayoutCachedWidth =17040
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =27270
                    Width =975
                    Height =255
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =26
                    BorderColor =0
                    Name ="PayOffReceived"
                    ControlSource ="PayOffReceived"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =27270
                    LayoutCachedWidth =28245
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =23
                    ColumnEnd =23
                    LayoutGroup =1
                    BackThemeColorIndex =7
                    BackTint =40.0
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =105
                    Width =1851
                    Height =255
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =3
                    BorderColor =14211288
                    ForeColor =10040879
                    Name ="SourceCompany"
                    ControlSource ="SourceCompany"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =105
                    LayoutCachedWidth =1956
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =15
                    Top =30
                    Width =45
                    Height =210
                    FontSize =9
                    Name ="ClientID"
                    ControlSource ="ClientID"

                    LayoutCachedLeft =15
                    LayoutCachedTop =30
                    LayoutCachedWidth =60
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =75
                    Top =30
                    Width =45
                    Height =210
                    FontSize =9
                    TabIndex =1
                    Name ="BuyerID"
                    ControlSource ="BuyerID"

                    LayoutCachedLeft =75
                    LayoutCachedTop =30
                    LayoutCachedWidth =120
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10320
                    Width =315
                    Height =255
                    ColumnOrder =1
                    FontSize =9
                    FontWeight =700
                    TabIndex =8
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10320
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =255
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
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12450
                    Width =945
                    Height =255
                    FontSize =8
                    TabIndex =10
                    BorderColor =0
                    Name ="Odometer"
                    ControlSource ="Odometer"
                    Format ="Standard"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =12450
                    LayoutCachedWidth =13395
                    LayoutCachedHeight =255
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
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14730
                    Width =855
                    Height =255
                    FontSize =8
                    TabIndex =13
                    BorderColor =0
                    Name ="Text308"
                    ControlSource ="InspectedDate"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =14730
                    LayoutCachedWidth =15585
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Top =255
                    Width =28440
                    Height =58
                    FontSize =9
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

                    LayoutCachedTop =255
                    LayoutCachedWidth =28440
                    LayoutCachedHeight =313
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17610
                    Width =1095
                    Height =255
                    FontSize =9
                    TabIndex =16
                    BorderColor =0
                    Name ="BuyOut"
                    ControlSource ="BuyOut"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =17610
                    LayoutCachedWidth =18705
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6765
                    Width =2835
                    Height =255
                    FontSize =8
                    TabIndex =6
                    BorderColor =0
                    Name ="Text478"
                    ControlSource ="Sch|Grp"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =6765
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =255
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
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9600
                    Width =720
                    Height =255
                    FontSize =8
                    TabIndex =7
                    BorderColor =0
                    Name ="Text487"
                    ControlSource ="UnitNum"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =9600
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =4
                    ForeShade =50.0
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
                    Left =18705
                    Width =1095
                    Height =255
                    FontSize =9
                    TabIndex =17
                    BorderColor =0
                    Name ="Text520"
                    ControlSource ="TargetSalePrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =18705
                    LayoutCachedWidth =19800
                    LayoutCachedHeight =255
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
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =19800
                    Width =1095
                    Height =255
                    FontSize =9
                    TabIndex =18
                    BorderColor =0
                    Name ="Text599"
                    ControlSource ="UnitSalePrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =19800
                    LayoutCachedWidth =20895
                    LayoutCachedHeight =255
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
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =20895
                    Width =660
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =19
                    BorderColor =0
                    Name ="BankPartnerx"
                    ControlSource ="=[BankPartner]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =20895
                    LayoutCachedWidth =21555
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =21555
                    Width =1515
                    Height =255
                    FontSize =9
                    TabIndex =20
                    BorderColor =0
                    Name ="Buyer"
                    ControlSource ="Buyer"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =21555
                    LayoutCachedWidth =23070
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
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
                    Left =10635
                    Width =1815
                    Height =255
                    FontSize =8
                    TabIndex =9
                    BorderColor =0
                    Name ="Text618"
                    ControlSource ="UnitVIN"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =10635
                    LayoutCachedWidth =12450
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
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
                    Left =26295
                    Width =975
                    Height =255
                    FontSize =9
                    TabIndex =25
                    BorderColor =0
                    Name ="PayOffRequested"
                    ControlSource ="PayOffRequested"
                    Format ="mm/dd/yy"
                    StatusBarText ="date unit parked"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =26295
                    LayoutCachedWidth =27270
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =22
                    ColumnEnd =22
                    LayoutGroup =1
                    BackThemeColorIndex =7
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =24345
                    Width =975
                    Height =255
                    FontSize =9
                    TabIndex =23
                    BorderColor =0
                    Name ="BuyOutDate"
                    ControlSource ="BuyOutDate"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =24345
                    LayoutCachedWidth =25320
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =20
                    ColumnEnd =20
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
                    Left =23370
                    Width =975
                    Height =255
                    FontSize =9
                    TabIndex =22
                    BorderColor =0
                    Name ="SUF Returned"
                    ControlSource ="SUF Returned"
                    Format ="mm/dd/yy"
                    EventProcPrefix ="SUF_Returned"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =23370
                    LayoutCachedWidth =24345
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =19
                    ColumnEnd =19
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
                    Left =25320
                    Width =975
                    Height =255
                    FontSize =9
                    TabIndex =24
                    BorderColor =0
                    Name ="InvoicePaid"
                    ControlSource ="InvoicePaid"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =25320
                    LayoutCachedWidth =26295
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =40.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =23070
                    Width =300
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =21
                    BorderColor =0
                    Name ="Text679"
                    ControlSource ="SalesPerson"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =23070
                    LayoutCachedWidth =23370
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
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
' See "RemarketingDashboardAdmin.cls"
