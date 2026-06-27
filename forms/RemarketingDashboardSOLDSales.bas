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
    ItemSuffix =682
    Left =10395
    Top =6060
    Right =24300
    Bottom =18915
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x6ed5887b0fb8e440
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
                Begin Line
                    OverlapFlags =93
                    Left =645
                    Top =885
                    Width =18000
                    BorderColor =12349952
                    Name ="Line88"
                    LayoutCachedLeft =645
                    LayoutCachedTop =885
                    LayoutCachedWidth =18645
                    LayoutCachedHeight =885
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =7980
                    Top =780
                    Width =1553
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
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
                Begin ListBox
                    Visible = NotDefault
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =7500
                    Top =960
                    Width =180
                    Height =1800
                    ColumnOrder =10
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    ForeColor =0
                    BorderColor =4210752
                    Name ="CityList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LCity From RemarketingReportAllUnits WHERE [UNITId] is not null  AND ([Cl"
                        "ientId]=4046)   GROUP BY LCity ORDER BY LCity;"
                    ColumnWidths ="1440"
                    Tag ="Group1"
                    ControlTipText ="Select the City to filter by"
                    GridlineColor =0

                    LayoutCachedLeft =7500
                    LayoutCachedTop =960
                    LayoutCachedWidth =7680
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
                    OverlapFlags =85
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =7740
                    Top =960
                    Width =225
                    Height =1800
                    ColumnOrder =9
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    ForeColor =0
                    BorderColor =4210752
                    Name ="StateList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LState From RemarketingReportAllUnits WHERE [UNITId] is not null  AND ([C"
                        "lientId]=4046)   GROUP BY LState ORDER BY LState;"
                    ColumnWidths ="562"
                    Tag ="Group1"
                    ControlTipText ="Select the States to filter by"
                    GridlineColor =0

                    LayoutCachedLeft =7740
                    LayoutCachedTop =960
                    LayoutCachedWidth =7965
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
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =8280
                    Top =2400
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
                    LayoutCachedLeft =8280
                    LayoutCachedTop =2400
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =2625
                End
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

                    LayoutCachedTop =3000
                    LayoutCachedWidth =25920
                    LayoutCachedHeight =3195
                    TabIndex =27
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =12405
                    Top =1200
                    Width =2280
                    Height =1080
                    BorderColor =10855845
                    Name ="Box179"
                    LayoutCachedLeft =12405
                    LayoutCachedTop =1200
                    LayoutCachedWidth =14685
                    LayoutCachedHeight =2280
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =6525
                    Top =2820
                    Width =705
                    Height =285
                    FontSize =10
                    FontWeight =700
                    ForeColor =13158
                    Name ="Label6"
                    Caption ="Asset"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =6525
                    LayoutCachedTop =2820
                    LayoutCachedWidth =7230
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
                    TextAlign =2
                    Left =19590
                    Top =2820
                    Width =1155
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label15"
                    Caption ="Sale Price"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19590
                    LayoutCachedTop =2820
                    LayoutCachedWidth =20745
                    LayoutCachedHeight =3105
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =135
                    Top =2820
                    Width =2016
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
                    LayoutCachedWidth =2151
                    LayoutCachedHeight =3105
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =10575
                    Top =2820
                    Width =990
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label21"
                    Caption ="Status"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =10575
                    LayoutCachedTop =2820
                    LayoutCachedWidth =11565
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
                    ColumnOrder =19
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
                    ColumnOrder =20
                    FontSize =8
                    TabIndex =3
                    Name ="txtCurrent"

                    LayoutCachedLeft =9360
                    LayoutCachedTop =240
                    LayoutCachedWidth =10125
                    LayoutCachedHeight =450
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =22680
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
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
                    BackStyle =0

                    LayoutCachedLeft =22680
                    LayoutCachedTop =120
                    LayoutCachedWidth =23235
                    LayoutCachedHeight =348
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    PictureType =1
                    Left =23265
                    Top =120
                    Width =750
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1
                    BackStyle =0

                    LayoutCachedLeft =23265
                    LayoutCachedTop =120
                    LayoutCachedWidth =24015
                    LayoutCachedHeight =348
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
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
                    ColumnOrder =13
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
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
                    ColumnOrder =14
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
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
                    ColumnOrder =15
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
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
                    ColumnOrder =16
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
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
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =16140
                    Top =960
                    Width =1860
                    Height =1800
                    ColumnOrder =17
                    FontSize =8
                    FontWeight =700
                    TabIndex =14
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

                    LayoutCachedLeft =16140
                    LayoutCachedTop =960
                    LayoutCachedWidth =18000
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
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =10560
                    Top =960
                    Width =990
                    Height =1800
                    ColumnOrder =18
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstStatus"
                    RowSourceType ="Value List"
                    RowSource ="\"S\";\"Sold-S\";\"KS\";\"Sold-KS\";\"PS\";\"Pending-PS\""
                    ColumnWidths ="0;944"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =10560
                    LayoutCachedTop =960
                    LayoutCachedWidth =11550
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
                    TabIndex =16
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
                    Left =10560
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
                    LayoutCachedLeft =10560
                    LayoutCachedTop =780
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =1005
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =16140
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
                    LayoutCachedLeft =16140
                    LayoutCachedTop =780
                    LayoutCachedWidth =16680
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
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
                    OverlapFlags =247
                    TextAlign =2
                    Left =11640
                    Top =2820
                    Width =1005
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label121"
                    Caption ="Odometer"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =11640
                    LayoutCachedTop =2820
                    LayoutCachedWidth =12645
                    LayoutCachedHeight =3105
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Image
                    BackStyle =1
                    PictureType =2
                    Left =12120
                    Top =1560
                    Width =480
                    Height =420
                    Name ="Image161"
                    Picture ="1_Excel-365-Green"

                    LayoutCachedLeft =12120
                    LayoutCachedTop =1560
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =1980
                    TabIndex =26
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =18360
                    Top =2280
                    Width =2460
                    Height =270
                    FontSize =8
                    Name ="Label162"
                    Caption ="┌───────  Pricing ───────┐"
                    LayoutCachedLeft =18360
                    LayoutCachedTop =2280
                    LayoutCachedWidth =20820
                    LayoutCachedHeight =2550
                    ForeThemeColorIndex =2
                End
                Begin Image
                    BackStyle =1
                    SizeMode =1
                    PictureType =1
                    Left =21780
                    Top =1260
                    Width =774
                    Height =954
                    Name ="Auto_Logo0"

                    LayoutCachedLeft =21780
                    LayoutCachedTop =1260
                    LayoutCachedWidth =22554
                    LayoutCachedHeight =2214
                    TabIndex =25
                End
                Begin ListBox
                    OverlapFlags =215
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =12645
                    Top =1200
                    Width =2040
                    Height =1080
                    ColumnOrder =12
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
                        "ERE (((LaunchReportsList.SecurityLevel)<=1) AND ((LaunchReportsList.FormGroup)=\""
                        "RmktDashBoard\")) ORDER BY LaunchReportsList.ReportPublicName;"
                    ColumnWidths ="2016;0;0;0"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="ReportList"
                    ControlTipText ="Sold Report Report of the currently filtered Units"
                    GridlineColor =0

                    LayoutCachedLeft =12645
                    LayoutCachedTop =1200
                    LayoutCachedWidth =14685
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
                    Left =18510
                    Top =960
                    Width =2025
                    Height =1320
                    ColumnOrder =11
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    ForeColor =0
                    BorderColor =4210752
                    Name ="TargetValueList"
                    RowSourceType ="Value List"
                    RowSource ="1;\"Target Set\";2;\"Target Not Set\""
                    ColumnWidths ="0;944"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"
                    ControlTipText ="Select the type of Unit Valuation to filter by, Results will run on click, singl"
                        "e selection"
                    GridlineColor =0

                    LayoutCachedLeft =18510
                    LayoutCachedTop =960
                    LayoutCachedWidth =20535
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
                    Left =18510
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
                    LayoutCachedLeft =18510
                    LayoutCachedTop =780
                    LayoutCachedWidth =19050
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
                    ColumnOrder =8
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
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
                    ForeColor =4754549
                    Name ="Label31"
                    Caption ="Sold Dashboard"
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
                    OverlapFlags =247
                    TextAlign =2
                    Left =12720
                    Top =2820
                    Width =1110
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label310"
                    Caption ="Inspected"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12720
                    LayoutCachedTop =2820
                    LayoutCachedWidth =13830
                    LayoutCachedHeight =3105
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =21675
                    Top =2820
                    Width =945
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BackColor =13017476
                    Name ="Label323"
                    Caption ="Margin"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =21675
                    LayoutCachedTop =2820
                    LayoutCachedWidth =22620
                    LayoutCachedHeight =3105
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =20745
                    Top =2820
                    Width =930
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BackColor =13017476
                    Name ="Label325"
                    Caption ="PnL Sale"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =20745
                    LayoutCachedTop =2820
                    LayoutCachedWidth =21675
                    LayoutCachedHeight =3105
                    ColumnStart =15
                    ColumnEnd =15
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
                    Left =18465
                    Top =2820
                    Width =1125
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label376"
                    Caption ="Target Price"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18465
                    LayoutCachedTop =2820
                    LayoutCachedWidth =19590
                    LayoutCachedHeight =3105
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =2085
                    Top =105
                    Width =840
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label421"
                    Caption ="Sales"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =2085
                    LayoutCachedTop =105
                    LayoutCachedWidth =2925
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19637
                    Top =2460
                    Width =1126
                    Height =270
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =23
                    Name ="GLLine"
                    ControlSource ="=Sum([UnitSalePrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =19637
                    LayoutCachedTop =2460
                    LayoutCachedWidth =20763
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
                    Left =18509
                    Top =2460
                    Width =1126
                    Height =270
                    ColumnOrder =6
                    FontSize =9
                    TabIndex =22
                    Name ="Text422"
                    ControlSource ="=Sum([TargetSalePrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShowDatePicker =0

                    LayoutCachedLeft =18509
                    LayoutCachedTop =2460
                    LayoutCachedWidth =19635
                    LayoutCachedHeight =2730
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =7305
                    Top =2820
                    Width =1995
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label480"
                    Caption ="Lease Sch |Group"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =7305
                    LayoutCachedTop =2820
                    LayoutCachedWidth =9300
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
                    TextAlign =1
                    Left =9375
                    Top =2820
                    Width =1125
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label486"
                    Caption ="Cust Unit#"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =9375
                    LayoutCachedTop =2820
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =3105
                    ColumnStart =4
                    ColumnEnd =4
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
                    ColumnOrder =5
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
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
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =7140
                    Top =1740
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
                    LayoutCachedLeft =7140
                    LayoutCachedTop =1740
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =1965
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =16125
                    Top =2820
                    Width =1815
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label609"
                    Caption ="Buyer"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16125
                    LayoutCachedTop =2820
                    LayoutCachedWidth =17940
                    LayoutCachedHeight =3105
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    Left =14880
                    Top =2820
                    Width =1185
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label610"
                    Caption ="Disposition "
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14880
                    LayoutCachedTop =2820
                    LayoutCachedWidth =16065
                    LayoutCachedHeight =3105
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =1
                    Left =23415
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
                    LayoutCachedLeft =23415
                    LayoutCachedTop =2820
                    LayoutCachedWidth =24390
                    LayoutCachedHeight =3105
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =119
                    TextAlign =2
                    IMESentenceMode =3
                    Left =23430
                    Top =2544
                    Width =975
                    Height =270
                    ColumnOrder =4
                    FontSize =8
                    FontWeight =700
                    TabIndex =24
                    BorderColor =4210752
                    Name ="SoldDateEnd"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =23430
                    LayoutCachedTop =2544
                    LayoutCachedWidth =24405
                    LayoutCachedHeight =2814
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =23430
                    Top =2214
                    Width =975
                    Height =270
                    ColumnOrder =0
                    FontSize =8
                    FontWeight =700
                    TabIndex =21
                    BorderColor =4210752
                    Name ="SoldDateStart"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =23430
                    LayoutCachedTop =2214
                    LayoutCachedWidth =24405
                    LayoutCachedHeight =2484
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =23460
                    Top =1980
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =18
                    ForeColor =12349952
                    Name ="Sold0"
                    Caption ="16"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =23460
                    LayoutCachedTop =1980
                    LayoutCachedWidth =23745
                    LayoutCachedHeight =2205
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =23760
                    Top =1980
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =19
                    ForeColor =12349952
                    Name ="Sold1"
                    Caption ="15"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =23760
                    LayoutCachedTop =1980
                    LayoutCachedWidth =24045
                    LayoutCachedHeight =2205
                    ColumnStart =1
                    ColumnEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =24060
                    Top =1980
                    Width =285
                    Height =225
                    FontSize =7
                    TabIndex =20
                    ForeColor =12349952
                    Name ="Sold2"
                    Caption ="14"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =24060
                    LayoutCachedTop =1980
                    LayoutCachedWidth =24345
                    LayoutCachedHeight =2205
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =22755
                    Top =2820
                    Width =660
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label628"
                    Caption ="SP"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =22755
                    LayoutCachedTop =2820
                    LayoutCachedWidth =23415
                    LayoutCachedHeight =3105
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =22620
                    Top =2820
                    Width =135
                    Height =285
                    Name ="EmptyCell658"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =22620
                    LayoutCachedTop =2820
                    LayoutCachedWidth =22755
                    LayoutCachedHeight =3105
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =13890
                    Top =2820
                    Width =435
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label242"
                    Caption ="Local"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13890
                    LayoutCachedTop =2820
                    LayoutCachedWidth =14325
                    LayoutCachedHeight =3105
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =14385
                    Top =2820
                    Width =435
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label292"
                    Caption ="Web"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14385
                    LayoutCachedTop =2820
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =3105
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17970
                    Top =2820
                    Width =495
                    Height =285
                    Name ="EmptyCell666"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =17970
                    LayoutCachedTop =2820
                    LayoutCachedWidth =18465
                    LayoutCachedHeight =3105
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
            End
        End
        Begin Section
            Height =315
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10575
                    Top =30
                    Width =990
                    Height =255
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =10
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =1
                    RightPadding =38

                    LayoutCachedLeft =10575
                    LayoutCachedTop =30
                    LayoutCachedWidth =11565
                    LayoutCachedHeight =285
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
                    FontUnderline = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6525
                    Top =30
                    Width =705
                    Height =255
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =7
                    BorderColor =0
                    ForeColor =4754549
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =6525
                    LayoutCachedTop =30
                    LayoutCachedWidth =7230
                    LayoutCachedHeight =285
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
                    Top =30
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
                    RightPadding =0
                    TextFormat =1
                    ShowDatePicker =0

                    LayoutCachedLeft =2175
                    LayoutCachedTop =30
                    LayoutCachedWidth =6495
                    LayoutCachedHeight =285
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
                    TextFontCharSet =186
                    TextAlign =3
                    IMESentenceMode =3
                    Left =19590
                    Top =30
                    Width =1155
                    Height =255
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =18
                    BorderColor =0
                    Name ="FMVBase2"
                    ControlSource ="UnitSalePrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =19590
                    LayoutCachedTop =30
                    LayoutCachedWidth =20745
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                    CurrencySymbol ="$"
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
                    Top =30
                    Width =2016
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

                    LayoutCachedLeft =135
                    LayoutCachedTop =30
                    LayoutCachedWidth =2151
                    LayoutCachedHeight =285
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
                    OverlapFlags =85
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11640
                    Top =30
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
                    RightPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =11640
                    LayoutCachedTop =30
                    LayoutCachedWidth =12645
                    LayoutCachedHeight =285
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
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12720
                    Top =30
                    Width =1110
                    Height =255
                    FontSize =10
                    TabIndex =12
                    BorderColor =0
                    Name ="Text308"
                    ControlSource ="InspectedDate"
                    GroupTable =1
                    ShowDatePicker =0

                    LayoutCachedLeft =12720
                    LayoutCachedTop =30
                    LayoutCachedWidth =13830
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =21675
                    Top =30
                    Width =945
                    Height =255
                    FontSize =10
                    TabIndex =20
                    BorderColor =0
                    Name ="PnLTarget"
                    ControlSource ="Margin"
                    Format ="Percent"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =21675
                    LayoutCachedTop =30
                    LayoutCachedWidth =22620
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =20745
                    Top =30
                    Width =930
                    Height =255
                    FontSize =10
                    TabIndex =19
                    BorderColor =0
                    Name ="Text326"
                    ControlSource ="TotalProfit"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =20745
                    LayoutCachedTop =30
                    LayoutCachedWidth =21675
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
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
                    Height =29
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

                    LayoutCachedTop =270
                    LayoutCachedWidth =25920
                    LayoutCachedHeight =299
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
                    Left =18465
                    Top =30
                    Width =1125
                    Height =255
                    FontSize =10
                    TabIndex =17
                    BorderColor =0
                    Name ="TargetSalePrice"
                    ControlSource ="TargetSalePrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =18465
                    LayoutCachedTop =30
                    LayoutCachedWidth =19590
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
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
                    Left =7305
                    Top =30
                    Width =1995
                    Height =255
                    FontSize =9
                    TabIndex =8
                    BorderColor =0
                    Name ="Text478"
                    ControlSource ="Sch|Grp"
                    GroupTable =1
                    RightPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =7305
                    LayoutCachedTop =30
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =285
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
                    Left =9375
                    Top =30
                    Width =1125
                    Height =255
                    FontSize =10
                    TabIndex =9
                    BorderColor =0
                    Name ="Text487"
                    ControlSource ="UnitNum"
                    GroupTable =1
                    RightPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =9375
                    LayoutCachedTop =30
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =285
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
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14880
                    Top =30
                    Width =1185
                    Height =255
                    FontSize =10
                    TabIndex =15
                    BorderColor =0
                    Name ="Text605"
                    ControlSource ="BuyerType"
                    GroupTable =1
                    ShowDatePicker =0

                    LayoutCachedLeft =14880
                    LayoutCachedTop =30
                    LayoutCachedWidth =16065
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
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
                    Left =16125
                    Top =30
                    Width =1815
                    Height =255
                    FontSize =10
                    TabIndex =16
                    BorderColor =0
                    Name ="Text607"
                    ControlSource ="Buyer"
                    GroupTable =1
                    ShowDatePicker =0

                    LayoutCachedLeft =16125
                    LayoutCachedTop =30
                    LayoutCachedWidth =17940
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =23415
                    Top =30
                    Width =975
                    Height =255
                    FontSize =9
                    TabIndex =22
                    BorderColor =0
                    Name ="InvoicePaid"
                    ControlSource ="InvoicePaid"
                    Format ="Short Date"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =23415
                    LayoutCachedTop =30
                    LayoutCachedWidth =24390
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =22755
                    Top =30
                    Width =660
                    Height =255
                    FontSize =10
                    TabIndex =21
                    BorderColor =0
                    Name ="Text626"
                    ControlSource ="SalesPerson"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    ShowDatePicker =0

                    LayoutCachedLeft =22755
                    LayoutCachedTop =30
                    LayoutCachedWidth =23415
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =22620
                    Top =30
                    Width =135
                    Height =255
                    Name ="EmptyCell659"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =22620
                    LayoutCachedTop =30
                    LayoutCachedWidth =22755
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
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
                    Left =13890
                    Top =30
                    Width =435
                    Height =255
                    FontSize =10
                    TabIndex =13
                    BorderColor =0
                    Name ="LocalLink"
                    ControlSource ="[InspectionLinkLocal]"
                    GroupTable =1
                    ShowDatePicker =0

                    LayoutCachedLeft =13890
                    LayoutCachedTop =30
                    LayoutCachedWidth =14325
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =10
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
                    Left =14385
                    Top =30
                    Width =435
                    Height =255
                    FontSize =10
                    TabIndex =14
                    BorderColor =0
                    Name ="WebLink"
                    ControlSource ="InspectionLinkWeb"
                    GroupTable =1
                    ShowDatePicker =0

                    LayoutCachedLeft =14385
                    LayoutCachedTop =30
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =10
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =17970
                    Top =30
                    Width =495
                    Height =255
                    Name ="EmptyCell667"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =17970
                    LayoutCachedTop =30
                    LayoutCachedWidth =18465
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
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
' See "RemarketingDashboardSOLDSales.cls"
