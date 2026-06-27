Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14400
    DatasheetFontHeight =11
    ItemSuffix =20
    Left =5670
    Top =4650
    Right =21855
    Bottom =14115
    RecSrcDt = Begin
        0x5079b08c37c5e440
    End
    DatasheetFontName ="Calibri"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =720
            BackColor =8210719
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    BackStyle =1
                    PictureType =2
                    Top =540
                    Width =14400
                    Height =180
                    BorderColor =6108695
                    Name ="Image142"
                    Picture ="1_DkBlue_Shadow"
                    GridlineColor =10921638
                    HorizontalAnchor =2

                    LayoutCachedTop =540
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =720
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =255
                    Left =960
                    Top =300
                    Width =3660
                    Height =360
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label7"
                    Caption ="Invoiced Units - Reverse Lookup"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedLeft =960
                    LayoutCachedTop =300
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =660
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =247
                    Left =960
                    Width =3660
                    Height =360
                    FontSize =14
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label8"
                    Caption ="Remarketing"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedLeft =960
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =360
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Image
                    SizeMode =0
                    PictureType =1
                    Left =120
                    Width =780
                    Height =540
                    BorderColor =10921638
                    Name ="EventImage"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    ControlSource ="ImageType"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =900
                    LayoutCachedHeight =540
                    TabIndex =1
                    RowEnd =1
                    LayoutGroup =2
                    GroupTable =2
                End
            End
        End
        Begin Section
            Height =6180
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =10
                    Left =300
                    Top =1462
                    Width =13830
                    Height =4605
                    FontSize =9
                    TabIndex =11
                    BoundColumn =2
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="InvoicedUnitsList"
                    RowSourceType ="Table/Query"
                    RowSource ="Select * From RmktInvoiceRequestReverseLookUp Where UnitDesc Like'*2010*' Order "
                        "By UnitDesc;"
                    ColumnWidths ="0;1440;576;1008;864;864;864;2160;4320"
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =1462
                    LayoutCachedWidth =14130
                    LayoutCachedHeight =6067
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4188
                    Top =1200
                    Width =864
                    FontSize =9
                    TabIndex =4
                    BackColor =13434879
                    BorderColor =967423
                    Name ="AssetSearch"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =4188
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5052
                    LayoutCachedHeight =1440
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5916
                    Top =1207
                    Width =2160
                    FontSize =9
                    TabIndex =9
                    BackColor =13434879
                    BorderColor =967423
                    Name ="VINSearch"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =5916
                    LayoutCachedTop =1207
                    LayoutCachedWidth =8076
                    LayoutCachedHeight =1447
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8076
                    Top =1207
                    Width =4320
                    FontSize =9
                    TabIndex =10
                    BackColor =13434879
                    BorderColor =967423
                    Name ="DescSearch"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =8076
                    LayoutCachedTop =1207
                    LayoutCachedWidth =12396
                    LayoutCachedHeight =1447
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5052
                    Top =1207
                    Width =864
                    FontSize =9
                    TabIndex =8
                    BackColor =13434879
                    BorderColor =967423
                    Name ="UnitNumSearch"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =5052
                    LayoutCachedTop =1207
                    LayoutCachedWidth =5916
                    LayoutCachedHeight =1447
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =13380
                    Top =1080
                    Width =540
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =15527148
                    ForeColor =12349952
                    Name ="ClearUnAssignedUnits"
                    Caption ="Clear▼"
                    ControlTipText ="Click to clear the selection below"
                    LayoutCachedLeft =13380
                    LayoutCachedTop =1080
                    LayoutCachedWidth =13920
                    LayoutCachedHeight =1305
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =300
                    Top =1207
                    FontSize =9
                    TabIndex =5
                    BackColor =13434879
                    BorderColor =967423
                    Name ="BuyerSearch"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =1207
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =1447
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1740
                    Top =1207
                    Width =576
                    FontSize =9
                    TabIndex =6
                    BackColor =13434879
                    BorderColor =967423
                    Name ="ReqSearch"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1740
                    LayoutCachedTop =1207
                    LayoutCachedWidth =2316
                    LayoutCachedHeight =1447
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2316
                    Top =1207
                    Width =1008
                    FontSize =9
                    TabIndex =7
                    BackColor =13434879
                    BorderColor =967423
                    Name ="InvNumSearch"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =2316
                    LayoutCachedTop =1207
                    LayoutCachedWidth =3324
                    LayoutCachedHeight =1447
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =600
                    Top =960
                    Width =8550
                    Height =225
                    FontSize =9
                    Name ="Label88"
                    Caption =" ┌──── Search   ───┬─────┬───────────────────┬─────────┬─────────────┬──────────"
                        "──────────────┐"
                    LayoutCachedLeft =600
                    LayoutCachedTop =960
                    LayoutCachedWidth =9150
                    LayoutCachedHeight =1185
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =330
                    Top =450
                    Width =1426
                    Height =270
                    FontSize =8
                    FontWeight =700
                    Name ="OpenBuyer"
                    Caption ="Go To Buyer"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    GridlineColor =10921638
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =330
                    LayoutCachedTop =450
                    LayoutCachedWidth =1756
                    LayoutCachedHeight =720
                    LayoutGroup =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =1
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1770
                    Top =450
                    Width =1570
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="OpenInvoice"
                    Caption ="Go To Invoice"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    GridlineColor =10921638
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1770
                    LayoutCachedTop =450
                    LayoutCachedWidth =3340
                    LayoutCachedHeight =720
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =1
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =3360
                    Top =450
                    Width =810
                    Height =270
                    Name ="EmptyCell12"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    GridlineColor =10921638
                    LayoutCachedLeft =3360
                    LayoutCachedTop =450
                    LayoutCachedWidth =4170
                    LayoutCachedHeight =720
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4185
                    Top =450
                    Width =3855
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="OpenUnit"
                    Caption ="Go To Unit"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    GridlineColor =10921638
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4185
                    LayoutCachedTop =450
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =720
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =1
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8055
                    Top =450
                    Width =3460
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="OpenSchedule"
                    Caption ="Go To Schedule Control"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    GridlineColor =10921638
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8055
                    LayoutCachedTop =450
                    LayoutCachedWidth =11515
                    LayoutCachedHeight =720
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =1
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12985
                    Top =480
                    Width =689
                    Height =300
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    TabIndex =12
                    BorderColor =14277081
                    Name ="UnitsShown"
                    ControlSource ="=[InvoicedUnitsList].[ListCount]"

                    LayoutCachedLeft =12985
                    LayoutCachedTop =480
                    LayoutCachedWidth =13674
                    LayoutCachedHeight =780
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =12960
                            Top =300
                            Width =750
                            Height =540
                            FontSize =7
                            BackColor =14211288
                            BorderColor =14211288
                            Name ="Label63"
                            Caption ="Units"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =12960
                            LayoutCachedTop =300
                            LayoutCachedWidth =13710
                            LayoutCachedHeight =840
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
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
' See "RmktInvoiceRequestReverseLookUpForm.cls"
