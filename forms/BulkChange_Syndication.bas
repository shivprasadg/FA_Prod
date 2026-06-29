Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
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
    Width =21846
    DatasheetFontHeight =11
    ItemSuffix =528
    Left =4680
    Top =3315
    Right =26475
    Bottom =7455
    ShortcutMenuBar ="SortClip"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0x89763c60f055e640
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Bulk_Update_SelectedUnits_SYND"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1188
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =7320
                    Top =60
                    Width =6300
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label419"
                    Caption ="┌───────────────────  Interim Funding ───────────────────┐"
                    LayoutCachedLeft =7320
                    LayoutCachedTop =60
                    LayoutCachedWidth =13620
                    LayoutCachedHeight =330
                    BackThemeColorIndex =4
                    BackTint =40.0
                    ForeThemeColorIndex =0
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Left =12
                    Top =996
                    Width =21744
                    Height =192
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedLeft =12
                    LayoutCachedTop =996
                    LayoutCachedWidth =21756
                    LayoutCachedHeight =1188
                    TabIndex =18
                    LayoutGroup =1
                    GroupTable =7
                End
                Begin Label
                    OverlapFlags =93
                    Left =120
                    Top =225
                    Width =6192
                    Height =510
                    FontSize =20
                    FontWeight =700
                    Name ="Label115"
                    Caption ="Interim Funding / Syndication Fields"
                    LayoutCachedLeft =120
                    LayoutCachedTop =225
                    LayoutCachedWidth =6312
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =0
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =7296
                    Top =288
                    Width =1692
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="IFBankSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =8
                    LeftPadding =0
                    RightPadding =0
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7296
                    LayoutCachedTop =288
                    LayoutCachedWidth =8988
                    LayoutCachedHeight =516
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =8
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =15
                    Width =765
                    Height =210
                    ColumnOrder =2
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedLeft =15
                    LayoutCachedWidth =780
                    LayoutCachedHeight =210
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =948
                    Top =804
                    Width =1992
                    Height =276
                    FontSize =9
                    Name ="Label181"
                    Caption ="VIN"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =948
                    LayoutCachedTop =804
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =120
                    Top =804
                    Width =765
                    Height =276
                    FontSize =9
                    Name ="Label180"
                    Caption ="♦AssetID"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =804
                    LayoutCachedWidth =885
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =2
                    Left =11556
                    Top =804
                    Width =1128
                    Height =276
                    FontSize =9
                    Name ="LabelIRP"
                    Caption ="Pay Off Date"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11556
                    LayoutCachedTop =804
                    LayoutCachedWidth =12684
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =9024
                    Top =804
                    Width =1260
                    Height =276
                    FontSize =9
                    Name ="Label175"
                    Caption ="Amount"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9024
                    LayoutCachedTop =804
                    LayoutCachedWidth =10284
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =2
                    Left =10356
                    Top =804
                    Width =1128
                    Height =276
                    FontSize =9
                    Name ="Label176"
                    Caption ="Fund Date"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10356
                    LayoutCachedTop =804
                    LayoutCachedWidth =11484
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =10356
                    Top =288
                    Width =1128
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="IFDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =8
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10356
                    LayoutCachedTop =288
                    LayoutCachedWidth =11484
                    LayoutCachedHeight =516
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =8
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =119
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =10356
                    Top =552
                    Width =1128
                    Height =252
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =11
                    Name ="IFDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10356
                    LayoutCachedTop =552
                    LayoutCachedWidth =11484
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =9024
                            Top =552
                            Width =1260
                            Height =252
                            FontSize =9
                            Name ="Label306"
                            Caption ="Interim All"
                            GroupTable =8
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =9024
                            LayoutCachedTop =552
                            LayoutCachedWidth =10284
                            LayoutCachedHeight =804
                            RowStart =1
                            RowEnd =1
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =2
                            BackThemeColorIndex =2
                            BackTint =60.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =8
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =9024
                    Top =288
                    Width =1260
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="IFAmountSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =8
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9024
                    LayoutCachedTop =288
                    LayoutCachedWidth =10284
                    LayoutCachedHeight =516
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =8
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =2
                    Left =7296
                    Top =804
                    Width =1692
                    Height =276
                    FontSize =9
                    Name ="Label220"
                    Caption ="Bank"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7296
                    LayoutCachedTop =804
                    LayoutCachedWidth =8988
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin ComboBox
                    OverlapFlags =119
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6840
                    Left =7296
                    Top =552
                    Width =1692
                    Height =252
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =10
                    Name ="IFBankCBO"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankID, Banks.BankShortName, Banks.BankName FROM Banks ORDER BY Ban"
                        "ks.BankShortName;"
                    ColumnWidths ="0;1800;5040"
                    StatusBarText ="bank used"
                    GroupTable =8
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =7296
                    LayoutCachedTop =552
                    LayoutCachedWidth =8988
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =11556
                    Top =288
                    Width =1128
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="PayOffDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =8
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11556
                    LayoutCachedTop =288
                    LayoutCachedWidth =12684
                    LayoutCachedHeight =516
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverColor =11830108
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =8
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =119
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11556
                    Top =552
                    Width =1128
                    Height =252
                    ColumnOrder =6
                    FontSize =9
                    TabIndex =12
                    Name ="PayOffDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11556
                    LayoutCachedTop =552
                    LayoutCachedWidth =12684
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =14805
                    Top =60
                    Width =3225
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label984"
                    Caption ="┌───────  Syndication Values ─────────┐"
                    LayoutCachedLeft =14805
                    LayoutCachedTop =60
                    LayoutCachedWidth =18030
                    LayoutCachedHeight =330
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14460
                    Top =552
                    Width =1068
                    Height =252
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =14
                    Name ="SyndRentCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14460
                    LayoutCachedTop =552
                    LayoutCachedWidth =15528
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =2
                    BackThemeColorIndex =9
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =14460
                    Top =804
                    Width =1068
                    Height =276
                    FontSize =9
                    BackColor =12566463
                    Name ="Label825"
                    Caption ="Unit Rent"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14460
                    LayoutCachedTop =804
                    LayoutCachedWidth =15528
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =120
                    Top =552
                    Width =765
                    Height =252
                    Name ="EmptyCell339"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =552
                    LayoutCachedWidth =885
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =948
                    Top =552
                    Width =1992
                    Height =252
                    Name ="EmptyCell340"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =948
                    LayoutCachedTop =552
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =120
                    Top =288
                    Width =765
                    Height =228
                    Name ="EmptyCell353"
                    GroupTable =8
                    LayoutCachedLeft =120
                    LayoutCachedTop =288
                    LayoutCachedWidth =885
                    LayoutCachedHeight =516
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =948
                    Top =288
                    Width =1992
                    Height =228
                    Name ="EmptyCell354"
                    GroupTable =8
                    LayoutCachedLeft =948
                    LayoutCachedTop =288
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =516
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =14460
                    Top =288
                    Width =1068
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="SyndRentSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =8
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14460
                    LayoutCachedTop =288
                    LayoutCachedWidth =15528
                    LayoutCachedHeight =516
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1643706
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedColor =1643706
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =8
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =15600
                    Top =804
                    Width =1068
                    Height =276
                    FontSize =9
                    BackColor =14151142
                    Name ="Label992"
                    Caption ="Residual"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15600
                    LayoutCachedTop =804
                    LayoutCachedWidth =16668
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =119
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15600
                    Top =552
                    Width =1068
                    Height =252
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =15
                    BackColor =14151142
                    Name ="SynResidCbo"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15600
                    LayoutCachedTop =552
                    LayoutCachedWidth =16668
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =15600
                    Top =288
                    Width =1068
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="SynResidSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =8
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15600
                    LayoutCachedTop =288
                    LayoutCachedWidth =16668
                    LayoutCachedHeight =516
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =5
                    HoverShade =75.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =8
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =13752
                    Top =804
                    Width =636
                    Height =276
                    FontSize =8
                    Name ="Label521"
                    Caption ="MSO"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13752
                    LayoutCachedTop =804
                    LayoutCachedWidth =14388
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =3000
                    Top =288
                    Width =4260
                    Height =228
                    Name ="EmptyCell401"
                    GroupTable =8
                    LayoutCachedLeft =3000
                    LayoutCachedTop =288
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =516
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =3000
                    Top =552
                    Width =4260
                    Height =252
                    Name ="EmptyCell402"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3000
                    LayoutCachedTop =552
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =247
                    Left =3000
                    Top =804
                    Width =4260
                    Height =276
                    FontSize =9
                    Name ="Label405"
                    Caption ="Unit Description"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3000
                    LayoutCachedTop =804
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =18600
                    Top =804
                    Width =1032
                    Height =276
                    FontSize =9
                    BackColor =14151142
                    Name ="Label163"
                    Caption ="Unit Cost"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18600
                    LayoutCachedTop =804
                    LayoutCachedWidth =19632
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =18600
                    Top =288
                    Width =1032
                    Height =228
                    Name ="EmptyCell442"
                    GroupTable =8
                    LayoutCachedLeft =18600
                    LayoutCachedTop =288
                    LayoutCachedWidth =19632
                    LayoutCachedHeight =516
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =13752
                    Top =288
                    Width =636
                    Height =228
                    Name ="EmptyCell446"
                    GroupTable =8
                    LayoutCachedLeft =13752
                    LayoutCachedTop =288
                    LayoutCachedWidth =14388
                    LayoutCachedHeight =516
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =13752
                    Top =552
                    Width =636
                    Height =252
                    Name ="EmptyCell450"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13752
                    LayoutCachedTop =552
                    LayoutCachedWidth =14388
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =18600
                    Top =552
                    Width =1032
                    Height =252
                    Name ="EmptyCell451"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =18600
                    LayoutCachedTop =552
                    LayoutCachedWidth =19632
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =19692
                    Top =288
                    Width =1032
                    Height =228
                    Name ="EmptyCell454"
                    GroupTable =8
                    LayoutCachedLeft =19692
                    LayoutCachedTop =288
                    LayoutCachedWidth =20724
                    LayoutCachedHeight =516
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =19692
                    Top =552
                    Width =1032
                    Height =252
                    Name ="EmptyCell455"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19692
                    LayoutCachedTop =552
                    LayoutCachedWidth =20724
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =19692
                    Top =804
                    Width =1032
                    Height =276
                    FontSize =9
                    BackColor =14151142
                    Name ="Label458"
                    Caption ="Child Cost"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19692
                    LayoutCachedTop =804
                    LayoutCachedWidth =20724
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =20784
                    Top =288
                    Width =1032
                    Height =228
                    Name ="EmptyCell462"
                    GroupTable =8
                    LayoutCachedLeft =20784
                    LayoutCachedTop =288
                    LayoutCachedWidth =21816
                    LayoutCachedHeight =516
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin EmptyCell
                    Left =20784
                    Top =552
                    Width =1032
                    Height =252
                    Name ="EmptyCell463"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =20784
                    LayoutCachedTop =552
                    LayoutCachedWidth =21816
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =2
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =20784
                    Top =804
                    Width =1032
                    Height =276
                    FontSize =9
                    BackColor =14151142
                    Name ="Label466"
                    Caption ="Total Cost"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =20784
                    LayoutCachedTop =804
                    LayoutCachedWidth =21816
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =16740
                    Top =804
                    Width =864
                    Height =276
                    FontSize =9
                    Name ="Label493"
                    Caption ="FA Guar"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16740
                    LayoutCachedTop =804
                    LayoutCachedWidth =17604
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =119
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16740
                    Top =552
                    Width =864
                    Height =252
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =16
                    BackColor =11193849
                    Name ="RVGuarCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="Sets the RV Gap (Guarantee) for selected Units"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16740
                    LayoutCachedTop =552
                    LayoutCachedWidth =17604
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =16740
                    Top =288
                    Width =864
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="RVGuarSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sets the RV Gap (Guarantee) for selected Units"
                    GroupTable =8
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =16740
                    LayoutCachedTop =288
                    LayoutCachedWidth =17604
                    LayoutCachedHeight =516
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =489889
                    BorderThemeColorIndex =1
                    HoverColor =489889
                    PressedColor =489889
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =8
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =1
                    Left =12756
                    Top =804
                    Width =936
                    Height =276
                    FontSize =9
                    Name ="Label203"
                    Caption ="DRAW"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12756
                    LayoutCachedTop =804
                    LayoutCachedWidth =13692
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =12756
                    Top =288
                    Width =936
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="DrawDownSubmit"
                    Caption ="▼ Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =8
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12756
                    LayoutCachedTop =288
                    LayoutCachedWidth =13692
                    LayoutCachedHeight =516
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =8
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =119
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =4320
                    Left =12756
                    Top =552
                    Width =936
                    Height =252
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =13
                    BackColor =15590879
                    Name ="DrawDownCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DrawDowns.DrawDownId, DrawDowns.DrawNumber AS Draw, DrawDowns.DrawName, B"
                        "anks.BankShortName AS Bank, Banks.BankID FROM DrawDowns INNER JOIN Banks ON Draw"
                        "Downs.DrawBankId = Banks.bankID WHERE (((DrawDowns.GroupID)=[Forms]![BulkChange_"
                        "Form]![GrpID]));"
                    ColumnWidths ="0;432;1296;864;0"
                    Tag ="Medit"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12756
                    LayoutCachedTop =552
                    LayoutCachedWidth =13692
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17676
                    Top =552
                    Width =864
                    Height =252
                    FontSize =9
                    TabIndex =17
                    BackColor =15788753
                    Name ="RVGuarOEMCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="Sets the RV Gap (Guarantee) for selected Units"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =17676
                    LayoutCachedTop =552
                    LayoutCachedWidth =18540
                    LayoutCachedHeight =804
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =17676
                    Top =804
                    Width =864
                    Height =276
                    FontSize =9
                    Name ="Label519"
                    Caption ="OEM Guar"
                    GroupTable =8
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17676
                    LayoutCachedTop =804
                    LayoutCachedWidth =18540
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =17676
                    Top =288
                    Width =864
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="RVGuarOEMSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sets the RV Gap (Guarantee) for selected Units"
                    GroupTable =8
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17676
                    LayoutCachedTop =288
                    LayoutCachedWidth =18540
                    LayoutCachedHeight =516
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =9732411
                    BorderThemeColorIndex =1
                    HoverColor =9732411
                    PressedColor =9732411
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =8
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =386
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11556
                    Top =48
                    Width =1128
                    Height =276
                    FontSize =9
                    TabIndex =6
                    Name ="UnitPayOff"
                    ControlSource ="UnitPayOff"
                    Format ="Short Date"
                    StatusBarText ="yes unit is IRP'ed"
                    GroupTable =8

                    LayoutCachedLeft =11556
                    LayoutCachedTop =48
                    LayoutCachedWidth =12684
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9024
                    Top =48
                    Width =1260
                    Height =276
                    FontSize =9
                    TabIndex =4
                    Name ="UnitIFAmt"
                    ControlSource ="UnitIFAmt"
                    StatusBarText ="IRP Acct Number"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =8

                    LayoutCachedLeft =9024
                    LayoutCachedTop =48
                    LayoutCachedWidth =10284
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10356
                    Top =48
                    Width =1128
                    Height =276
                    FontSize =9
                    TabIndex =5
                    Name ="UnitIFDate"
                    ControlSource ="UnitIFDate"
                    Format ="Short Date"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =8

                    LayoutCachedLeft =10356
                    LayoutCachedTop =48
                    LayoutCachedWidth =11484
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =120
                    Top =48
                    Width =765
                    Height =276
                    FontSize =9
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =8

                    LayoutCachedLeft =120
                    LayoutCachedTop =48
                    LayoutCachedWidth =885
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =948
                    Top =48
                    Width =1992
                    Height =276
                    FontSize =9
                    TabIndex =1
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =8

                    LayoutCachedLeft =948
                    LayoutCachedTop =48
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =314
                    Width =21744
                    Height =72
                    TabIndex =15
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End
                    HorizontalAnchor =2

                    LayoutCachedTop =314
                    LayoutCachedWidth =21744
                    LayoutCachedHeight =386
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =87
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6840
                    Left =7296
                    Top =48
                    Width =1692
                    Height =276
                    FontSize =9
                    TabIndex =3
                    Name ="UnitIFBank"
                    ControlSource ="UnitIFBankID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankID, Banks.BankShortName, Banks.BankName FROM Banks ORDER BY Ban"
                        "ks.BankShortName;"
                    ColumnWidths ="0;1800;5040"
                    StatusBarText ="bank used"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =8
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =7296
                    LayoutCachedTop =48
                    LayoutCachedWidth =8988
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14460
                    Top =48
                    Width =1068
                    Height =276
                    FontSize =9
                    TabIndex =9
                    Name ="UNITRENT"
                    ControlSource ="UNITRENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="MONTHLY RENT"
                    GroupTable =8

                    LayoutCachedLeft =14460
                    LayoutCachedTop =48
                    LayoutCachedWidth =15528
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =2
                    BackThemeColorIndex =9
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15600
                    Top =48
                    Width =1068
                    Height =276
                    FontSize =9
                    TabIndex =10
                    BackColor =14151142
                    Name ="UnitSynResid"
                    ControlSource ="UnitSynResid"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =8

                    LayoutCachedLeft =15600
                    LayoutCachedTop =48
                    LayoutCachedWidth =16668
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13752
                    Top =48
                    Width =636
                    Height =276
                    FontSize =9
                    TabIndex =8
                    Name ="MSO"
                    ControlSource ="=IIf([UnitMSORecd] & [UnitMSOCopy]<>\"\",\"Y\",\"\")"
                    Format ="Short Date"
                    GroupTable =8

                    LayoutCachedLeft =13752
                    LayoutCachedTop =48
                    LayoutCachedWidth =14388
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3000
                    Top =48
                    Width =4260
                    Height =276
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="Text397"
                    ControlSource ="UnitDesc"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =8

                    LayoutCachedLeft =3000
                    LayoutCachedTop =48
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =18600
                    Top =48
                    Width =1032
                    Height =276
                    FontSize =9
                    TabIndex =13
                    Name ="VendorCost"
                    ControlSource ="VendorCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =8

                    LayoutCachedLeft =18600
                    LayoutCachedTop =48
                    LayoutCachedWidth =19632
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =20784
                    Top =48
                    Width =1032
                    Height =276
                    FontSize =9
                    TabIndex =16
                    Name ="Text460"
                    ControlSource ="UnitOEM"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =8

                    LayoutCachedLeft =20784
                    LayoutCachedTop =48
                    LayoutCachedWidth =21816
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =19692
                    Top =48
                    Width =1032
                    Height =276
                    FontSize =9
                    TabIndex =14
                    Name ="ChildPartCost"
                    ControlSource ="ChildPartCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =8

                    LayoutCachedLeft =19692
                    LayoutCachedTop =48
                    LayoutCachedWidth =20724
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16740
                    Top =48
                    Width =864
                    Height =276
                    FontSize =9
                    TabIndex =11
                    BackColor =11193849
                    Name ="RVGuar"
                    ControlSource ="RVGuar"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Set The RV Gap (Guarantee)"
                    GroupTable =8

                    LayoutCachedLeft =16740
                    LayoutCachedTop =48
                    LayoutCachedWidth =17604
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =4320
                    Left =12756
                    Top =48
                    Width =936
                    Height =276
                    FontSize =9
                    TabIndex =7
                    BackColor =15590879
                    Name ="DrawDownId"
                    ControlSource ="DrawDownId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DrawDowns.DrawDownId, DrawDowns.DrawNumber AS Draw, DrawDowns.DrawName, B"
                        "anks.BankShortName AS Bank, Banks.BankID FROM DrawDowns INNER JOIN Banks ON Draw"
                        "Downs.DrawBankId = Banks.bankID WHERE (((DrawDowns.GroupID)=[Forms]![BulkChange_"
                        "Form]![GrpID]));"
                    ColumnWidths ="0;432;1296;864;0"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    AfterUpdate ="[Event Procedure]"
                    Format ="Standard"
                    GroupTable =8

                    LayoutCachedLeft =12756
                    LayoutCachedTop =48
                    LayoutCachedWidth =13692
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17676
                    Top =48
                    Width =864
                    Height =276
                    FontSize =9
                    TabIndex =12
                    BackColor =15788753
                    Name ="RVGuarOem"
                    ControlSource ="RVGuarOem"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Set The RV Gap (Guarantee)"
                    GroupTable =8

                    LayoutCachedLeft =17676
                    LayoutCachedTop =48
                    LayoutCachedWidth =18540
                    LayoutCachedHeight =324
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =8
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
' See "BulkChange_Syndication.cls"
