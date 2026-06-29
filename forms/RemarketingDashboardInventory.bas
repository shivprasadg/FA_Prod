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
    Width =24180
    DatasheetFontHeight =11
    ItemSuffix =295
    Left =5550
    Top =3150
    Right =29730
    Bottom =15240
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x378941d0f5b6e440
    End
    RecordSource ="SELECT * FROM RemarketingReportAllUnits WHERE [UNITId] is not null  AND ([Client"
        "Id] is not null)  AND ([UnitYr] > 1)  AND ([UnitType] <> '')  AND ([UnitMake] <>"
        " '') AND ([UnitStatus] <> '')  Order By SourceCompany;"
    Caption ="Inventory"
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
                    Width =24120
                    Height =540
                    BorderColor =14211288
                    Name ="Box49"
                    HorizontalAnchor =2
                    LayoutCachedWidth =24120
                    LayoutCachedHeight =540
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Left =60
                    Top =3000
                    Width =24120
                    Height =195
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"

                    LayoutCachedLeft =60
                    LayoutCachedTop =3000
                    LayoutCachedWidth =24180
                    LayoutCachedHeight =3195
                    TabIndex =25
                End
                Begin Line
                    OverlapFlags =93
                    Left =21237
                    Top =885
                    Width =0
                    Height =1008
                    BorderColor =12349952
                    Name ="Line203"
                    LayoutCachedLeft =21237
                    LayoutCachedTop =885
                    LayoutCachedWidth =21237
                    LayoutCachedHeight =1893
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =17985
                    Top =1140
                    Width =2280
                    Height =1080
                    BorderColor =10855845
                    Name ="Box179"
                    LayoutCachedLeft =17985
                    LayoutCachedTop =1140
                    LayoutCachedWidth =20265
                    LayoutCachedHeight =2220
                End
                Begin Line
                    OverlapFlags =95
                    Left =645
                    Top =885
                    Width =20592
                    BorderColor =12349952
                    Name ="Line88"
                    LayoutCachedLeft =645
                    LayoutCachedTop =885
                    LayoutCachedWidth =21237
                    LayoutCachedHeight =885
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =3360
                    Top =2820
                    Width =705
                    Height =285
                    FontSize =10
                    FontWeight =700
                    ForeColor =13158
                    Name ="Label6"
                    Caption ="♦ Asset"
                    GroupTable =1
                    LayoutCachedLeft =3360
                    LayoutCachedTop =2820
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =3105
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =11715
                    Top =2580
                    Width =3435
                    Height =270
                    FontSize =8
                    Name ="Label20"
                    Caption ="┌───────────── Parked Location ─────────┐"
                    LayoutCachedLeft =11715
                    LayoutCachedTop =2580
                    LayoutCachedWidth =15150
                    LayoutCachedHeight =2850
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =2652
                    Top =2820
                    Width =648
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Year"
                    GroupTable =1
                    LayoutCachedLeft =2652
                    LayoutCachedTop =2820
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =3105
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =5364
                    Top =2820
                    Width =1380
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label3"
                    Caption ="Make"
                    GroupTable =1
                    LayoutCachedLeft =5364
                    LayoutCachedTop =2820
                    LayoutCachedWidth =6744
                    LayoutCachedHeight =3105
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =6804
                    Top =2820
                    Width =1596
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label4"
                    Caption ="♦ Buyer"
                    GroupTable =1
                    LayoutCachedLeft =6804
                    LayoutCachedTop =2820
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =3105
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =10980
                    Top =2820
                    Width =1392
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label9"
                    Caption ="Location"
                    GroupTable =1
                    LayoutCachedLeft =10980
                    LayoutCachedTop =2820
                    LayoutCachedWidth =12372
                    LayoutCachedHeight =3105
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =14484
                    Top =2820
                    Width =1452
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label10"
                    Caption ="City"
                    GroupTable =1
                    LayoutCachedLeft =14484
                    LayoutCachedTop =2820
                    LayoutCachedWidth =15936
                    LayoutCachedHeight =3105
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =186
                    TextAlign =2
                    Left =19248
                    Top =2820
                    Width =1248
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label14"
                    Caption ="Sold Price"
                    GroupTable =1
                    LayoutCachedLeft =19248
                    LayoutCachedTop =2820
                    LayoutCachedWidth =20496
                    LayoutCachedHeight =3105
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =16632
                    Top =2820
                    Width =1248
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label15"
                    Caption ="FMV Base"
                    GroupTable =1
                    LayoutCachedLeft =16632
                    LayoutCachedTop =2820
                    LayoutCachedWidth =17880
                    LayoutCachedHeight =3105
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =186
                    TextAlign =2
                    Left =17940
                    Top =2820
                    Width =1248
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label16"
                    Caption ="Target Sale"
                    GroupTable =1
                    LayoutCachedLeft =17940
                    LayoutCachedTop =2820
                    LayoutCachedWidth =19188
                    LayoutCachedHeight =3105
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =1
                    Left =132
                    Top =2820
                    Width =2460
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label18"
                    Caption ="♦ Source Company"
                    GroupTable =1
                    LayoutCachedLeft =132
                    LayoutCachedTop =2820
                    LayoutCachedWidth =2592
                    LayoutCachedHeight =3105
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =8472
                    Top =2820
                    Width =1152
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label21"
                    Caption ="Status"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =8472
                    LayoutCachedTop =2820
                    LayoutCachedWidth =9624
                    LayoutCachedHeight =3105
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =9780
                    Top =2340
                    Width =1553
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =20
                    Name ="ResetForm"
                    Caption ="Reset Form"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =9780
                    LayoutCachedTop =2340
                    LayoutCachedWidth =11333
                    LayoutCachedHeight =2580
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =2524406
                    BorderColor =2524406
                    HoverThemeColorIndex =9
                    HoverShade =50.0
                    PressedColor =2524406
                    QuickStyle =35
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    Left =20556
                    Top =2820
                    Width =1188
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label37"
                    Caption ="Sold Date"
                    GroupTable =1
                    LayoutCachedLeft =20556
                    LayoutCachedTop =2820
                    LayoutCachedWidth =21744
                    LayoutCachedHeight =3105
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =20535
                    Top =2490
                    Width =1155
                    Height =270
                    ColumnOrder =13
                    FontSize =10
                    TabIndex =21
                    BorderColor =4210752
                    Name ="SoldDateEnd"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =20535
                    LayoutCachedTop =2490
                    LayoutCachedWidth =21690
                    LayoutCachedHeight =2760
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =20535
                    Top =2160
                    Width =1155
                    Height =270
                    ColumnOrder =14
                    FontSize =10
                    TabIndex =19
                    BorderColor =4210752
                    Name ="SoldDateStart"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"

                    LayoutCachedLeft =20535
                    LayoutCachedTop =2160
                    LayoutCachedWidth =21690
                    LayoutCachedHeight =2430
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
                    ColumnOrder =15
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
                    ColumnOrder =16
                    FontSize =8
                    TabIndex =3
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
                    Left =15996
                    Top =2820
                    Width =576
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label61"
                    Caption ="State"
                    GroupTable =1
                    LayoutCachedLeft =15996
                    LayoutCachedTop =2820
                    LayoutCachedWidth =16572
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
                    Left =4128
                    Top =2820
                    Width =1176
                    Height =285
                    FontSize =10
                    FontWeight =700
                    ForeColor =13158
                    Name ="Label70"
                    Caption ="Unit Type"
                    GroupTable =1
                    LayoutCachedLeft =4128
                    LayoutCachedTop =2820
                    LayoutCachedWidth =5304
                    LayoutCachedHeight =3105
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
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
                    Width =2460
                    Height =1800
                    ColumnOrder =7
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    BoundColumn =1
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstSourceCompany"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SourceCompany, ClientId From RemarketingReportAllUnits WHERE [UNITId] is "
                        "not null  AND ([ClientId] is not null)  AND ([UnitYr] > 1)  AND ([UnitType] <> '"
                        "')  AND ([UnitMake] <> '') AND ([UnitStatus] <> '')   GROUP BY SourceCompany, Cl"
                        "ientId ORDER BY SourceCompany;"
                    ColumnWidths ="2016;0"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =120
                    LayoutCachedTop =960
                    LayoutCachedWidth =2580
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
                    Left =2640
                    Top =960
                    Width =705
                    Height =1800
                    ColumnOrder =8
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstModelYear"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitYr From RemarketingReportAllUnits WHERE [UNITId] is not null  AND ([C"
                        "lientId] is not null)  AND ([UnitYr] > 1)  AND ([UnitType] <> '')  AND ([UnitMak"
                        "e] <> '') AND ([UnitStatus] <> '')   GROUP BY UnitYr ORDER BY UnitYr desc;"
                    ColumnWidths ="1440"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =2640
                    LayoutCachedTop =960
                    LayoutCachedWidth =3345
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
                    Left =4110
                    Top =960
                    Width =1170
                    Height =1800
                    ColumnOrder =9
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstUnitType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitType From RemarketingReportAllUnits WHERE [UNITId] is not null  AND ("
                        "[ClientId] is not null)  AND ([UnitYr] > 1)  AND ([UnitType] <> '')  AND ([UnitM"
                        "ake] <> '') AND ([UnitStatus] <> '')   GROUP BY UnitType ORDER BY UnitType;"
                    ColumnWidths ="1296"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =4110
                    LayoutCachedTop =960
                    LayoutCachedWidth =5280
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
                    Left =5340
                    Top =960
                    Width =1380
                    Height =1800
                    ColumnOrder =10
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstMake"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UnitMake From RemarketingReportAllUnits WHERE [UNITId] is not null  AND ("
                        "[ClientId] is not null)  AND ([UnitYr] > 1)  AND ([UnitType] <> '')  AND ([UnitM"
                        "ake] <> '') AND ([UnitStatus] <> '')   GROUP BY UnitMake ORDER BY UnitMake;"
                    ColumnWidths ="1296"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =5340
                    LayoutCachedTop =960
                    LayoutCachedWidth =6720
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
                    Left =6780
                    Top =960
                    Width =1590
                    Height =1800
                    ColumnOrder =11
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    BoundColumn =1
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstBuyer"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT BuyerName,BuyerId From RemarketingReportAllUnits WHERE [UNITId] is not nu"
                        "ll  AND ([ClientId] is not null)  AND ([UnitYr] > 1)  AND ([UnitType] <> '')  AN"
                        "D ([UnitMake] <> '') AND ([UnitStatus] <> '')   GROUP BY BuyerName, BuyerId ORDE"
                        "R BY BuyerName;"
                    ColumnWidths ="144;0"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =6780
                    LayoutCachedTop =960
                    LayoutCachedWidth =8370
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
                    Left =9792
                    Top =1080
                    Width =1553
                    Height =1140
                    FontSize =8
                    FontWeight =700
                    TabIndex =14
                    Name ="cmdSetFilters"
                    Caption ="After Setting Filters Click HERE\015\012Run the Results"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9792
                    LayoutCachedTop =1080
                    LayoutCachedWidth =11345
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
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    MultiSelect =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8460
                    Top =960
                    Width =1140
                    Height =1800
                    ColumnOrder =12
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    ForeColor =0
                    BorderColor =4210752
                    Name ="lstStatus"
                    RowSourceType ="Value List"
                    RowSource ="\"S\";\"Sold-S\";\"KS\";\"Sold-KS\";\"K\";\"Inventory-K\";\"R\";\"Inventory-R\";"
                        "\"TR\";\"Inventory-TR\""
                    ColumnWidths ="0;944"
                    Tag ="Group1"
                    GridlineColor =0

                    LayoutCachedLeft =8460
                    LayoutCachedTop =960
                    LayoutCachedWidth =9600
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
                    Left =2640
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
                    LayoutCachedLeft =2640
                    LayoutCachedTop =780
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =4110
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
                    LayoutCachedLeft =4110
                    LayoutCachedTop =780
                    LayoutCachedWidth =4650
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =5340
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
                    LayoutCachedLeft =5340
                    LayoutCachedTop =780
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =6780
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
                    LayoutCachedLeft =6780
                    LayoutCachedTop =780
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =8460
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
                    LayoutCachedLeft =8460
                    LayoutCachedTop =780
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =1005
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =9792
                    Top =780
                    Width =1553
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="ClearAndRun"
                    Caption ="Clear Filters and Run"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9792
                    LayoutCachedTop =780
                    LayoutCachedWidth =11345
                    LayoutCachedHeight =1020
                    ColumnStart =1
                    ColumnEnd =1
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
                    OverlapFlags =215
                    TextAlign =1
                    Left =16740
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
                    LayoutCachedLeft =16740
                    LayoutCachedTop =780
                    LayoutCachedWidth =19095
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =9696
                    Top =2820
                    Width =1212
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label121"
                    Caption ="Odometer"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =9696
                    LayoutCachedTop =2820
                    LayoutCachedWidth =10908
                    LayoutCachedHeight =3105
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =21780
                    Top =2205
                    Width =720
                    Height =210
                    FontSize =8
                    Name ="Label159"
                    Caption ="Start"
                    LayoutCachedLeft =21780
                    LayoutCachedTop =2205
                    LayoutCachedWidth =22500
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =21780
                    Top =2535
                    Width =720
                    Height =210
                    FontSize =8
                    Name ="Label160"
                    Caption ="End"
                    LayoutCachedLeft =21780
                    LayoutCachedTop =2535
                    LayoutCachedWidth =22500
                    LayoutCachedHeight =2745
                    ForeThemeColorIndex =2
                End
                Begin Image
                    BackStyle =1
                    PictureType =2
                    Left =17700
                    Top =1500
                    Width =480
                    Height =420
                    Name ="Image161"
                    Picture ="1_Excel-365-Green"

                    LayoutCachedLeft =17700
                    LayoutCachedTop =1500
                    LayoutCachedWidth =18180
                    LayoutCachedHeight =1920
                    TabIndex =24
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =16785
                    Top =2580
                    Width =3225
                    Height =270
                    FontSize =8
                    Name ="Label162"
                    Caption ="┌─────────  Pricing ─────────┐"
                    LayoutCachedLeft =16785
                    LayoutCachedTop =2580
                    LayoutCachedWidth =20010
                    LayoutCachedHeight =2850
                    ForeThemeColorIndex =2
                End
                Begin Image
                    BackStyle =1
                    SizeMode =1
                    PictureType =1
                    Left =22920
                    Top =840
                    Width =774
                    Height =954
                    Name ="Auto_Logo0"

                    LayoutCachedLeft =22920
                    LayoutCachedTop =840
                    LayoutCachedWidth =23694
                    LayoutCachedHeight =1794
                    TabIndex =23
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =20532
                    Top =1908
                    Width =384
                    Height =228
                    FontSize =7
                    TabIndex =16
                    ForeColor =12349952
                    Name ="Sold0"
                    Caption ="2016"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =20532
                    LayoutCachedTop =1908
                    LayoutCachedWidth =20916
                    LayoutCachedHeight =2136
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderColor =14211288
                    HoverColor =10319446
                    PressedColor =15527148
                    PressedForeColor =11830108
                    GroupTable =2
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =20940
                    Top =1908
                    Width =372
                    Height =228
                    FontSize =7
                    TabIndex =17
                    ForeColor =12349952
                    Name ="Sold1"
                    Caption ="2015"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =20940
                    LayoutCachedTop =1908
                    LayoutCachedWidth =21312
                    LayoutCachedHeight =2136
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
                    GroupTable =2
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =21336
                    Top =1908
                    Width =372
                    Height =228
                    FontSize =7
                    TabIndex =18
                    ForeColor =12349952
                    Name ="Sold2"
                    Caption ="2014"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =21336
                    LayoutCachedTop =1908
                    LayoutCachedWidth =21708
                    LayoutCachedHeight =2136
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
                    GroupTable =2
                    Overlaps =1
                End
                Begin ListBox
                    OverlapFlags =215
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =18225
                    Top =1140
                    Width =2040
                    Height =1080
                    ColumnOrder =6
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    BackColor =14277081
                    ForeColor =0
                    BorderColor =4210752
                    Name ="ReportsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LaunchReportsList.ReportPublicName, LaunchReportsList.ReportDesc, LaunchR"
                        "eportsList.Command, LaunchReportsList.FormCalled, LaunchReportsList.SecurityLeve"
                        "l, LaunchReportsList.Tags FROM LaunchReportsList ORDER BY LaunchReportsList.Repo"
                        "rtPublicName;"
                    ColumnWidths ="2016;0;0;0"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="ReportList"
                    ControlTipText ="Sold Report Report of the currently filtered Units"
                    GridlineColor =0

                    LayoutCachedLeft =18225
                    LayoutCachedTop =1140
                    LayoutCachedWidth =20265
                    LayoutCachedHeight =2220
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
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11640
                    Top =1020
                    Width =510
                    Height =840
                    ColumnOrder =5
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    ForeColor =0
                    BorderColor =4210752
                    Name ="TargetValueList"
                    RowSourceType ="Value List"
                    RowSource ="0;All;1;\"Target Set\";2;\"Target Not Set\";3;\"FMV Set\";4;\"FMV Not Set\";5;\""
                        "Base Set\";6;\"Base Not Set\""
                    ColumnWidths ="0;944"
                    OnClick ="[Event Procedure]"
                    Tag ="Group1"
                    ControlTipText ="Select the type of Unit Valuation to filter by, Results will run on click, singl"
                        "e selection"
                    GridlineColor =0

                    LayoutCachedLeft =11640
                    LayoutCachedTop =1020
                    LayoutCachedWidth =12150
                    LayoutCachedHeight =1860
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
                    OverlapFlags =223
                    TextAlign =1
                    Left =11640
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
                    LayoutCachedLeft =11640
                    LayoutCachedTop =840
                    LayoutCachedWidth =12180
                    LayoutCachedHeight =1065
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =12432
                    Top =2820
                    Width =1992
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label239"
                    Caption ="Address"
                    GroupTable =1
                    LayoutCachedLeft =12432
                    LayoutCachedTop =2820
                    LayoutCachedWidth =14424
                    LayoutCachedHeight =3105
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =21804
                    Top =2820
                    Width =576
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label242"
                    Caption ="Local"
                    GroupTable =1
                    LayoutCachedLeft =21804
                    LayoutCachedTop =2820
                    LayoutCachedWidth =22380
                    LayoutCachedHeight =3105
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
                    Left =15990
                    Top =960
                    Width =630
                    Height =1800
                    ColumnOrder =0
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    ForeColor =0
                    BorderColor =4210752
                    Name ="StateList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RemarketingReportAllUnits.PState FROM RemarketingReportAllUnits WHERE ((("
                        "RemarketingReportAllUnits.[UNITId]) Is Not Null) AND ((RemarketingReportAllUnits"
                        ".[ClientId]) Is Not Null) AND ((RemarketingReportAllUnits.[UnitYr])>1) AND ((Rem"
                        "arketingReportAllUnits.[UnitType])<>'') AND ((RemarketingReportAllUnits.[UnitMak"
                        "e])<>'') AND ((RemarketingReportAllUnits.[UnitStatus])<>'')) GROUP BY Remarketin"
                        "gReportAllUnits.PState ORDER BY RemarketingReportAllUnits.PState;"
                    ColumnWidths ="562"
                    Tag ="Group1"
                    ControlTipText ="Select the States to filter by"
                    GridlineColor =0

                    LayoutCachedLeft =15990
                    LayoutCachedTop =960
                    LayoutCachedWidth =16620
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
                    Left =14475
                    Top =960
                    Height =1800
                    ColumnOrder =4
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    ForeColor =0
                    BorderColor =4210752
                    Name ="CityList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RemarketingReportAllUnits.PCity FROM RemarketingReportAllUnits WHERE (((R"
                        "emarketingReportAllUnits.[UNITId]) Is Not Null) AND ((RemarketingReportAllUnits."
                        "[ClientId]) Is Not Null) AND ((RemarketingReportAllUnits.[UnitYr])>1) AND ((Rema"
                        "rketingReportAllUnits.[UnitType])<>'') AND ((RemarketingReportAllUnits.[UnitMake"
                        "])<>'') AND ((RemarketingReportAllUnits.[UnitStatus])<>'')) GROUP BY Remarketing"
                        "ReportAllUnits.PCity ORDER BY RemarketingReportAllUnits.PCity;"
                    ColumnWidths ="1440"
                    Tag ="Group1"
                    ControlTipText ="Select the City to filter by"
                    GridlineColor =0

                    LayoutCachedLeft =14475
                    LayoutCachedTop =960
                    LayoutCachedWidth =15915
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
                    Left =15990
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
                    LayoutCachedLeft =15990
                    LayoutCachedTop =780
                    LayoutCachedWidth =16530
                    LayoutCachedHeight =1005
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    Left =14475
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
                    LayoutCachedLeft =14475
                    LayoutCachedTop =780
                    LayoutCachedWidth =15015
                    LayoutCachedHeight =1005
                End
                Begin EmptyCell
                    Left =23076
                    Top =2820
                    Width =288
                    Height =285
                    Name ="EmptyCell274"
                    GroupTable =1
                    LayoutCachedLeft =23076
                    LayoutCachedTop =2820
                    LayoutCachedWidth =23364
                    LayoutCachedHeight =3105
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =23424
                    Top =2820
                    Width =288
                    Height =285
                    Name ="EmptyCell277"
                    GroupTable =1
                    LayoutCachedLeft =23424
                    LayoutCachedTop =2820
                    LayoutCachedWidth =23712
                    LayoutCachedHeight =3105
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =22440
                    Top =2820
                    Width =576
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label292"
                    Caption ="Web"
                    GroupTable =1
                    LayoutCachedLeft =22440
                    LayoutCachedTop =2820
                    LayoutCachedWidth =23016
                    LayoutCachedHeight =3105
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11833
                    Top =240
                    Width =689
                    Height =300
                    FontSize =12
                    FontWeight =700
                    TabIndex =22
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
                    Width =7620
                    Height =555
                    FontSize =20
                    Name ="Label31"
                    Caption ="Inventory Dashboard"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =7740
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
            End
        End
        Begin Section
            Height =332
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4128
                    Top =36
                    Width =1176
                    Height =252
                    FontSize =9
                    TabIndex =6
                    BorderColor =0
                    Name ="UnitType"
                    ControlSource ="UnitType"
                    StatusBarText ="unit make"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =4128
                    LayoutCachedTop =36
                    LayoutCachedWidth =5304
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
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15996
                    Top =36
                    Width =576
                    Height =252
                    FontSize =10
                    TabIndex =14
                    BorderColor =0
                    Name ="LState"
                    ControlSource ="PState"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =15996
                    LayoutCachedTop =36
                    LayoutCachedWidth =16572
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3360
                    Top =36
                    Width =705
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
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =3360
                    LayoutCachedTop =36
                    LayoutCachedWidth =4065
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
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2652
                    Top =36
                    Width =648
                    Height =252
                    ColumnOrder =14
                    FontSize =10
                    TabIndex =4
                    BorderColor =0
                    Name ="UnitYr"
                    ControlSource ="UnitYr"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =2652
                    LayoutCachedTop =36
                    LayoutCachedWidth =3300
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
                    TextFontCharSet =186
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5364
                    Top =36
                    Width =1380
                    Height =252
                    ColumnOrder =13
                    FontSize =10
                    TabIndex =7
                    BorderColor =0
                    Name ="UnitMake"
                    ControlSource ="UnitMake"
                    StatusBarText ="unit make"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =5364
                    LayoutCachedTop =36
                    LayoutCachedWidth =6744
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
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
                    Left =10980
                    Top =36
                    Width =1392
                    Height =252
                    ColumnOrder =10
                    FontSize =10
                    TabIndex =11
                    BorderColor =0
                    Name ="ParkedLocation"
                    ControlSource ="ParkedLocation"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =10980
                    LayoutCachedTop =36
                    LayoutCachedWidth =12372
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
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
                    Left =19248
                    Top =36
                    Width =1248
                    Height =252
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =17
                    BorderColor =0
                    Name ="UnitSold"
                    ControlSource ="UnitSalePrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =19248
                    LayoutCachedTop =36
                    LayoutCachedWidth =20496
                    LayoutCachedHeight =288
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6804
                    Top =36
                    Width =1596
                    Height =252
                    ColumnOrder =8
                    FontSize =10
                    TabIndex =8
                    BorderColor =14211288
                    ForeColor =10040879
                    Name ="Buyer"
                    ControlSource ="Buyer"
                    StatusBarText ="Short Name for Reporting"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =6804
                    LayoutCachedTop =36
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
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
                    Left =14484
                    Top =36
                    Width =1452
                    Height =252
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =13
                    BorderColor =0
                    Name ="LocCity"
                    ControlSource ="PCity"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =14484
                    LayoutCachedTop =36
                    LayoutCachedWidth =15936
                    LayoutCachedHeight =288
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
                    TextFontCharSet =186
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16632
                    Top =36
                    Width =1248
                    Height =252
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =15
                    BorderColor =0
                    Name ="FMVBase2"
                    ControlSource ="FMVBase"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =16632
                    LayoutCachedTop =36
                    LayoutCachedWidth =17880
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17940
                    Top =36
                    Width =1248
                    Height =252
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =16
                    BorderColor =0
                    Name ="TargetSalePrice"
                    ControlSource ="TargetSalePrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =17940
                    LayoutCachedTop =36
                    LayoutCachedWidth =19188
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
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
                    Left =132
                    Top =36
                    Width =2460
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
                    BottomPadding =38

                    LayoutCachedLeft =132
                    LayoutCachedTop =36
                    LayoutCachedWidth =2592
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =3
                    IMESentenceMode =3
                    Left =20556
                    Top =36
                    Width =1188
                    Height =252
                    FontSize =10
                    TabIndex =18
                    BorderColor =0
                    Name ="InvoicePaid"
                    ControlSource ="InvoicePaid"
                    Format ="Short Date"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =20556
                    LayoutCachedTop =36
                    LayoutCachedWidth =21744
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =127
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =270
                    Width =24120
                    Height =29
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
                    LayoutCachedWidth =24120
                    LayoutCachedHeight =299
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
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
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8472
                    Top =36
                    Width =1152
                    Height =252
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =9
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =8472
                    LayoutCachedTop =36
                    LayoutCachedWidth =9624
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
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
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9696
                    Top =36
                    Width =1212
                    Height =252
                    FontSize =10
                    TabIndex =10
                    BorderColor =0
                    Name ="Odometer"
                    ControlSource ="Odometer"
                    Format ="Standard"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =9696
                    LayoutCachedTop =36
                    LayoutCachedWidth =10908
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
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12432
                    Top =36
                    Width =1992
                    Height =252
                    FontSize =10
                    TabIndex =12
                    BorderColor =0
                    Name ="ParkedAddress"
                    ControlSource ="ParkedAddress"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =12432
                    LayoutCachedTop =36
                    LayoutCachedWidth =14424
                    LayoutCachedHeight =288
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
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =22440
                    Top =36
                    Width =576
                    Height =252
                    FontSize =10
                    TabIndex =20
                    BorderColor =0
                    Name ="InspectionLinkWeb"
                    ControlSource ="InspectionLinkWeb"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =22440
                    LayoutCachedTop =36
                    LayoutCachedWidth =23016
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    ForeThemeColorIndex =10
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
                    Left =21804
                    Top =36
                    Width =576
                    Height =252
                    FontSize =10
                    TabIndex =19
                    BorderColor =0
                    Name ="InspectionLinkLocal"
                    ControlSource ="InspectionLinkLocal"
                    GroupTable =1
                    BottomPadding =38
                    ShowDatePicker =0

                    LayoutCachedLeft =21804
                    LayoutCachedTop =36
                    LayoutCachedWidth =22380
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    ForeThemeColorIndex =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =23076
                    Top =36
                    Width =288
                    Height =252
                    Name ="EmptyCell288"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =23076
                    LayoutCachedTop =36
                    LayoutCachedWidth =23364
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =23424
                    Top =36
                    Width =288
                    Height =252
                    Name ="EmptyCell289"
                    GroupTable =1
                    BottomPadding =38
                    LayoutCachedLeft =23424
                    LayoutCachedTop =36
                    LayoutCachedWidth =23712
                    LayoutCachedHeight =288
                    RowStart =1
                    RowEnd =1
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =1
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
' See "RemarketingDashboardInventory.cls"
