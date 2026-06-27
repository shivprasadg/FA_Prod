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
    Width =27648
    DatasheetFontHeight =11
    ItemSuffix =675
    Left =4845
    Top =3045
    Right =22395
    Bottom =19125
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xc2d0fdff8fdde440
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
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =23295
                    Top =2325
                    Width =2160
                    Height =270
                    FontSize =8
                    Name ="Label162"
                    Caption ="┌──  Projected ──┐"
                    LayoutCachedLeft =23295
                    LayoutCachedTop =2325
                    LayoutCachedWidth =25455
                    LayoutCachedHeight =2595
                    ForeThemeColorIndex =2
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Width =27648
                    Height =540
                    BorderColor =14211288
                    Name ="Box49"
                    HorizontalAnchor =2
                    LayoutCachedWidth =27648
                    LayoutCachedHeight =540
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =3000
                    Width =27648
                    Height =195
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =3000
                    LayoutCachedWidth =27648
                    LayoutCachedHeight =3195
                    TabIndex =29
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =13125
                    Top =1200
                    Width =2280
                    Height =1080
                    BorderColor =10855845
                    Name ="Box179"
                    LayoutCachedLeft =13125
                    LayoutCachedTop =1200
                    LayoutCachedWidth =15405
                    LayoutCachedHeight =2280
                End
                Begin Line
                    OverlapFlags =93
                    Left =645
                    Top =885
                    Width =19440
                    BorderColor =12349952
                    Name ="Line88"
                    LayoutCachedLeft =645
                    LayoutCachedTop =885
                    LayoutCachedWidth =20085
                    LayoutCachedHeight =885
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =6495
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
                    LayoutCachedLeft =6495
                    LayoutCachedTop =2820
                    LayoutCachedWidth =7200
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
                    TextAlign =1
                    Left =2175
                    Top =2820
                    Width =4320
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Unit Description"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2175
                    LayoutCachedTop =2820
                    LayoutCachedWidth =6495
                    LayoutCachedHeight =3105
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =15960
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
                    LayoutCachedLeft =15960
                    LayoutCachedTop =2820
                    LayoutCachedWidth =17415
                    LayoutCachedHeight =3105
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =135
                    Top =2820
                    Width =2010
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label18"
                    Caption ="Account Name"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =135
                    LayoutCachedTop =2820
                    LayoutCachedWidth =2145
                    LayoutCachedHeight =3105
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =11160
                    Top =2820
                    Width =435
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
                    LayoutCachedLeft =11160
                    LayoutCachedTop =2820
                    LayoutCachedWidth =11595
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
                    ColumnOrder =23
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
                    ColumnOrder =24
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
                    Left =17415
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
                    LayoutCachedLeft =17415
                    LayoutCachedTop =2820
                    LayoutCachedWidth =17985
                    LayoutCachedHeight =3105
                    ColumnStart =12
                    ColumnEnd =12
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
                    ColumnOrder =17
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
                    ColumnOrder =18
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
                    ColumnOrder =19
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
                    ColumnOrder =20
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
                    Left =18540
                    Top =960
                    Width =1980
                    Height =1500
                    ColumnOrder =21
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

                    LayoutCachedLeft =18540
                    LayoutCachedTop =960
                    LayoutCachedWidth =20520
                    LayoutCachedHeight =2460
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
                    Left =10140
                    Top =960
                    Width =2462
                    Height =1800
                    ColumnOrder =22
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
                    ColumnWidths ="0;1008"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =10140
                    LayoutCachedTop =960
                    LayoutCachedWidth =12602
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
                    Left =2175
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
                    LayoutCachedLeft =2175
                    LayoutCachedTop =780
                    LayoutCachedWidth =2715
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =2895
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
                    LayoutCachedLeft =2895
                    LayoutCachedTop =780
                    LayoutCachedWidth =3435
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =4065
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
                    LayoutCachedLeft =4065
                    LayoutCachedTop =780
                    LayoutCachedWidth =4605
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =11160
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
                    LayoutCachedLeft =11160
                    LayoutCachedTop =780
                    LayoutCachedWidth =11700
                    LayoutCachedHeight =1005
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =18540
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
                    LayoutCachedLeft =18540
                    LayoutCachedTop =780
                    LayoutCachedWidth =19080
                    LayoutCachedHeight =1005
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
                    OverlapFlags =223
                    TextAlign =1
                    Left =12300
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
                    LayoutCachedLeft =12300
                    LayoutCachedTop =780
                    LayoutCachedWidth =14655
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =11595
                    Top =2820
                    Width =1005
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
                    LayoutCachedLeft =11595
                    LayoutCachedTop =2820
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =3105
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Image
                    BackStyle =1
                    PictureType =2
                    Left =12840
                    Top =1560
                    Width =480
                    Height =420
                    Name ="Image161"
                    Picture ="1_Excel-365-Green"

                    LayoutCachedLeft =12840
                    LayoutCachedTop =1560
                    LayoutCachedWidth =13320
                    LayoutCachedHeight =1980
                    TabIndex =28
                    BackThemeColorIndex =1
                    BackShade =85.0
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
                    TabIndex =27
                End
                Begin ListBox
                    OverlapFlags =215
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =13365
                    Top =1200
                    Width =2040
                    Height =1365
                    ColumnOrder =16
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
                        "ERE (((LaunchReportsList.SecurityLevel)<=9) AND ((LaunchReportsList.FormGroup)=\""
                        "RmktDashboard\")) ORDER BY LaunchReportsList.ReportPublicName;"
                    ColumnWidths ="2016;0;0;0"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="ReportList"
                    ControlTipText ="Sold Report Report of the currently filtered Units"
                    GridlineColor =0

                    LayoutCachedLeft =13365
                    LayoutCachedTop =1200
                    LayoutCachedWidth =15405
                    LayoutCachedHeight =2565
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
                    Left =21300
                    Top =1020
                    Width =2025
                    Height =1320
                    ColumnOrder =15
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

                    LayoutCachedLeft =21300
                    LayoutCachedTop =1020
                    LayoutCachedWidth =23325
                    LayoutCachedHeight =2340
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
                    OverlapFlags =93
                    TextAlign =1
                    Left =21300
                    Top =840
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
                    LayoutCachedLeft =21300
                    LayoutCachedTop =840
                    LayoutCachedWidth =21840
                    LayoutCachedHeight =1065
                End
                Begin ListBox
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =17400
                    Top =960
                    Width =585
                    Height =1800
                    ColumnOrder =13
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

                    LayoutCachedLeft =17400
                    LayoutCachedTop =960
                    LayoutCachedWidth =17985
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
                    Left =15945
                    Top =960
                    Height =1800
                    ColumnOrder =14
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    ForeColor =0
                    BorderColor =4210752
                    Name ="CityList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RemarketingReportAllUnits.PCity FROM RemarketingReportAllUnits WHERE (((R"
                        "emarketingReportAllUnits.[UNITId]) Is Not Null) AND ((RemarketingReportAllUnits."
                        "[ClientId])=4046)) GROUP BY RemarketingReportAllUnits.PCity ORDER BY Remarketing"
                        "ReportAllUnits.PCity;"
                    ColumnWidths ="1440"
                    Tag ="Group1"
                    ControlTipText ="Select the City to filter by"
                    GridlineColor =0

                    LayoutCachedLeft =15945
                    LayoutCachedTop =960
                    LayoutCachedWidth =17385
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
                    Left =17400
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
                    LayoutCachedLeft =17400
                    LayoutCachedTop =780
                    LayoutCachedWidth =17940
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =15945
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
                    LayoutCachedLeft =15945
                    LayoutCachedTop =780
                    LayoutCachedWidth =16485
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
                    ColumnOrder =12
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
                    Left =13110
                    Top =2820
                    Width =870
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label306"
                    Caption ="Parked"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13110
                    LayoutCachedTop =2820
                    LayoutCachedWidth =13980
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
                    Left =14850
                    Top =2820
                    Width =1110
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label310"
                    Caption ="Inspected"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14850
                    LayoutCachedTop =2820
                    LayoutCachedWidth =15960
                    LayoutCachedHeight =3105
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =12600
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
                    LayoutCachedLeft =12600
                    LayoutCachedTop =2820
                    LayoutCachedWidth =13110
                    LayoutCachedHeight =3105
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =23430
                    Top =2820
                    Width =1035
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BackColor =13017476
                    Name ="Label323"
                    Caption ="PnL Target"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =23430
                    LayoutCachedTop =2820
                    LayoutCachedWidth =24465
                    LayoutCachedHeight =3105
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =17985
                    Top =2820
                    Width =1125
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
                    LayoutCachedLeft =17985
                    LayoutCachedTop =2820
                    LayoutCachedWidth =19110
                    LayoutCachedHeight =3105
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =26055
                    Top =2820
                    Width =1125
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label419"
                    Caption ="┌Bank Share┐"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =26055
                    LayoutCachedTop =2820
                    LayoutCachedWidth =27180
                    LayoutCachedHeight =3105
                    ColumnStart =21
                    ColumnEnd =22
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =60.0
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =2505
                    Top =105
                    Width =1260
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label421"
                    Caption ="Manager"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =2505
                    LayoutCachedTop =105
                    LayoutCachedWidth =3765
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    BackStyle =0
                    IMESentenceMode =3
                    Left =22320
                    Top =2520
                    Width =1080
                    Height =270
                    ColumnOrder =11
                    FontSize =9
                    TabIndex =20
                    Name ="TargetLine"
                    ControlSource ="=Sum([TargetSalePrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =22320
                    LayoutCachedTop =2520
                    LayoutCachedWidth =23400
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
                    Left =17985
                    Top =2520
                    Width =1126
                    Height =270
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =16
                    Name ="Text422"
                    ControlSource ="=Sum([BuyOut])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =17985
                    LayoutCachedTop =2520
                    LayoutCachedWidth =19111
                    LayoutCachedHeight =2790
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    BackStyle =0
                    IMESentenceMode =3
                    Left =24495
                    Top =2520
                    Width =1050
                    Height =270
                    ColumnOrder =10
                    FontSize =8
                    TabIndex =22
                    Name ="Text444"
                    ControlSource ="=Sum([GainLoss])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =24495
                    LayoutCachedTop =2520
                    LayoutCachedWidth =25545
                    LayoutCachedHeight =2790
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =24465
                    Top =2820
                    Width =915
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BackColor =12566463
                    Name ="Label442"
                    Caption ="Gain Loss"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =24465
                    LayoutCachedTop =2820
                    LayoutCachedWidth =25380
                    LayoutCachedHeight =3105
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
                    Left =7200
                    Top =2820
                    Width =2940
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
                    LayoutCachedLeft =7200
                    LayoutCachedTop =2820
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =3105
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10140
                    Top =2820
                    Width =1020
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
                    LayoutCachedLeft =10140
                    LayoutCachedTop =2820
                    LayoutCachedWidth =11160
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
                    TextAlign =2
                    Left =22350
                    Top =2820
                    Width =1080
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
                    LayoutCachedLeft =22350
                    LayoutCachedTop =2820
                    LayoutCachedWidth =23430
                    LayoutCachedHeight =3105
                    ColumnStart =17
                    ColumnEnd =17
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
                    ColumnOrder =8
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
                    TextAlign =1
                    Left =5445
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
                    LayoutCachedLeft =5445
                    LayoutCachedTop =780
                    LayoutCachedWidth =5985
                    LayoutCachedHeight =1005
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =23443
                    Top =2520
                    Width =1037
                    Height =270
                    ColumnOrder =6
                    FontSize =8
                    TabIndex =21
                    Name ="Text612"
                    ControlSource ="=Sum([PnLTarget])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =23443
                    LayoutCachedTop =2520
                    LayoutCachedWidth =24480
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
                    Left =20220
                    Top =2520
                    Width =1037
                    Height =270
                    ColumnOrder =7
                    FontSize =8
                    TabIndex =18
                    Name ="Text613"
                    ControlSource ="=Sum([BookValue])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =20220
                    LayoutCachedTop =2520
                    LayoutCachedWidth =21257
                    LayoutCachedHeight =2790
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =19110
                    Top =2820
                    Width =1125
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label709"
                    Caption ="Term Penalty"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19110
                    LayoutCachedTop =2820
                    LayoutCachedWidth =20235
                    LayoutCachedHeight =3105
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19095
                    Top =2520
                    Width =1125
                    Height =270
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =17
                    Name ="Text725"
                    ControlSource ="=Sum([TermPenalty])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =19095
                    LayoutCachedTop =2520
                    LayoutCachedWidth =20220
                    LayoutCachedHeight =2790
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =20235
                    Top =2820
                    Width =1035
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label711"
                    Caption ="Book Value"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =20235
                    LayoutCachedTop =2820
                    LayoutCachedWidth =21270
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
                    Left =21270
                    Top =2820
                    Width =1080
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label639"
                    Caption ="Residual"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =21270
                    LayoutCachedTop =2820
                    LayoutCachedWidth =22350
                    LayoutCachedHeight =3105
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =21255
                    Top =2520
                    Width =1080
                    Height =270
                    ColumnOrder =4
                    FontSize =8
                    TabIndex =19
                    Name ="Text643"
                    ControlSource ="=Sum([Residual])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =21255
                    LayoutCachedTop =2520
                    LayoutCachedWidth =22335
                    LayoutCachedHeight =2790
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =13980
                    Top =2820
                    Width =870
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label650"
                    Caption ="EOL"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13980
                    LayoutCachedTop =2820
                    LayoutCachedWidth =14850
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
                    Left =12555
                    Top =2640
                    Width =2175
                    Height =270
                    FontSize =8
                    Name ="Label656"
                    Caption ="┌──  Projected Dates ──┐"
                    LayoutCachedLeft =12555
                    LayoutCachedTop =2640
                    LayoutCachedWidth =14730
                    LayoutCachedHeight =2910
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =25380
                    Top =2820
                    Width =675
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label667"
                    Caption ="Bank"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =25380
                    LayoutCachedTop =2820
                    LayoutCachedWidth =26055
                    LayoutCachedHeight =3105
                    ColumnStart =20
                    ColumnEnd =20
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
                    Left =25380
                    Top =1020
                    Width =810
                    Height =1800
                    ColumnOrder =0
                    FontSize =8
                    FontWeight =700
                    TabIndex =23
                    ForeColor =0
                    BorderColor =4210752
                    Name ="AssigneeList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BankPartner FROM RemarketingReportAllUnits WHERE UNITID Is Not Null GROUP"
                        " BY BankPartner;"
                    ColumnWidths ="1440"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =25380
                    LayoutCachedTop =1020
                    LayoutCachedWidth =26190
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
                    Top =2340
                    Width =1553
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =24
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
                    LayoutCachedTop =2340
                    LayoutCachedWidth =9533
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
                    PictureType =1
                    Left =26055
                    Top =120
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =25
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =26055
                    LayoutCachedTop =120
                    LayoutCachedWidth =26805
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
                    Left =25500
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =26
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

                    LayoutCachedLeft =25500
                    LayoutCachedTop =120
                    LayoutCachedWidth =26055
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
            Height =298
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =17415
                    Width =570
                    Height =255
                    FontSize =9
                    TabIndex =17
                    BorderColor =0
                    Name ="LState"
                    ControlSource ="PState"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =17415
                    LayoutCachedWidth =17985
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
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6495
                    Width =705
                    Height =255
                    ColumnOrder =2
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
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

                    LayoutCachedLeft =6495
                    LayoutCachedWidth =7200
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
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2175
                    Width =4320
                    Height =255
                    ColumnOrder =14
                    FontSize =10
                    TabIndex =6
                    BorderColor =0
                    Name ="UnitDesc"
                    ControlSource ="=\"<B>\" & [UnitYr] & \"</B> \" & [UnitType] & \" \" & [UnitMake] & \" \" & [Uni"
                        "tModel] & IIf(IsNull([UnitConfiguration]),Null,\" \" & [UnitConfiguration]) & II"
                        "f(IsNull([UnitUOM]),Null,\" \" & [UnitUOM])"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1
                    ShowDatePicker =0

                    LayoutCachedLeft =2175
                    LayoutCachedWidth =6495
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
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13110
                    Width =870
                    Height =255
                    FontSize =9
                    TabIndex =13
                    BorderColor =0
                    Name ="ParkDateProjected"
                    ControlSource ="ParkDateProjected"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =13110
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =255
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
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15960
                    Width =1455
                    Height =255
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =16
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

                    LayoutCachedLeft =15960
                    LayoutCachedWidth =17415
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
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
                    Left =135
                    Width =2010
                    Height =255
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =5
                    BorderColor =14211288
                    ForeColor =10040879
                    Name ="SourceCompany"
                    ControlSource ="SourceCompany"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =135
                    LayoutCachedWidth =2145
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =15
                    Top =30
                    Width =45
                    Height =210
                    FontSize =8
                    TabIndex =2
                    Name ="ClientID"
                    ControlSource ="ClientID"

                    LayoutCachedLeft =15
                    LayoutCachedTop =30
                    LayoutCachedWidth =60
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =75
                    Top =30
                    Width =45
                    Height =210
                    FontSize =8
                    TabIndex =3
                    Name ="BuyerID"
                    ControlSource ="BuyerID"

                    LayoutCachedLeft =75
                    LayoutCachedTop =30
                    LayoutCachedWidth =120
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11160
                    Width =435
                    Height =255
                    ColumnOrder =1
                    FontSize =9
                    FontWeight =700
                    TabIndex =10
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11160
                    LayoutCachedWidth =11595
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
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11595
                    Width =1005
                    Height =255
                    FontSize =10
                    TabIndex =11
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

                    LayoutCachedLeft =11595
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =255
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
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14850
                    Width =1110
                    Height =255
                    FontSize =9
                    TabIndex =15
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

                    LayoutCachedLeft =14850
                    LayoutCachedWidth =15960
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12600
                    Width =510
                    Height =255
                    FontSize =10
                    TabIndex =12
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

                    LayoutCachedLeft =12600
                    LayoutCachedWidth =13110
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
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
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =23430
                    Width =1035
                    Height =255
                    FontSize =10
                    TabIndex =23
                    BorderColor =0
                    Name ="PnLTarget"
                    ControlSource ="PnLTarget"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Target Price - Buy Out Cost"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =23430
                    LayoutCachedWidth =24465
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
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =127
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Top =240
                    Width =27648
                    Height =58
                    TabIndex =4
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End

                    LayoutCachedTop =240
                    LayoutCachedWidth =27648
                    LayoutCachedHeight =298
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
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17985
                    Width =1125
                    Height =255
                    FontSize =10
                    TabIndex =18
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

                    LayoutCachedLeft =17985
                    LayoutCachedWidth =19110
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
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =25380
                    Width =675
                    Height =255
                    FontSize =9
                    TabIndex =25
                    BorderColor =0
                    Name ="Text400"
                    ControlSource ="=[BankPartner]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =25380
                    LayoutCachedWidth =26055
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =26055
                    Width =270
                    Height =255
                    FontSize =8
                    TabIndex =26
                    BorderColor =0
                    Name ="BankShareYN"
                    ControlSource ="BankShareYN"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =26055
                    LayoutCachedWidth =26325
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =21
                    ColumnEnd =21
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ForeThemeColorIndex =0
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
                    Left =26325
                    Width =855
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =27
                    BorderColor =0
                    Name ="Margin"
                    ControlSource ="BankShareNet"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =26325
                    LayoutCachedWidth =27180
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =22
                    ColumnEnd =22
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ForeThemeColorIndex =0
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
                    Left =24465
                    Width =915
                    Height =255
                    FontSize =10
                    TabIndex =24
                    BorderColor =0
                    Name ="Text427"
                    ControlSource ="GainLoss"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =24465
                    LayoutCachedWidth =25380
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Visible = NotDefault
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =23520
                    Width =0
                    Height =255
                    FontSize =10
                    TabIndex =1
                    BorderColor =0
                    Name ="InspectionLinkWeb"
                    ControlSource ="InspectionLinkWeb"
                    GroupTable =4
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =23520
                    LayoutCachedWidth =23520
                    LayoutCachedHeight =255
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeThemeColorIndex =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin TextBox
                    Visible = NotDefault
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =23460
                    Width =0
                    Height =255
                    FontSize =10
                    BorderColor =0
                    Name ="InspectionLinkLocal"
                    ControlSource ="InspectionLinkLocal"
                    GroupTable =4
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =23460
                    LayoutCachedWidth =23460
                    LayoutCachedHeight =255
                    LayoutGroup =2
                    ForeThemeColorIndex =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =4
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7200
                    Width =2940
                    Height =255
                    FontSize =8
                    TabIndex =8
                    BorderColor =0
                    Name ="Text478"
                    ControlSource ="Sch|Grp"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =7200
                    LayoutCachedWidth =10140
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
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10140
                    Width =1020
                    Height =255
                    FontSize =8
                    TabIndex =9
                    BorderColor =0
                    Name ="Text487"
                    ControlSource ="UnitNum"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =10140
                    LayoutCachedWidth =11160
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
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =22350
                    Width =1080
                    Height =255
                    FontSize =10
                    TabIndex =22
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

                    LayoutCachedLeft =22350
                    LayoutCachedWidth =23430
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
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
                    Left =19110
                    Width =1125
                    Height =255
                    FontSize =10
                    TabIndex =19
                    BorderColor =0
                    Name ="Text703"
                    ControlSource ="TermPenalty"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =19110
                    LayoutCachedWidth =20235
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =20235
                    Width =1035
                    Height =255
                    FontSize =10
                    TabIndex =20
                    BorderColor =0
                    Name ="Text712"
                    ControlSource ="BookValue"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =20235
                    LayoutCachedWidth =21270
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
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =21270
                    Width =1080
                    Height =255
                    FontSize =10
                    TabIndex =21
                    BorderColor =0
                    Name ="Text633"
                    ControlSource ="Residual"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =21270
                    LayoutCachedWidth =22350
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
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
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13980
                    Width =870
                    Height =255
                    FontSize =9
                    TabIndex =14
                    BorderColor =0
                    Name ="Text644"
                    ControlSource ="ProjEOL"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =13980
                    LayoutCachedWidth =14850
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
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "RemarketingDashboardINVMgr.cls"
