Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    RecordLocks =2
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18000
    DatasheetFontHeight =11
    ItemSuffix =320
    Left =4110
    Top =7170
    Right =19380
    Bottom =15315
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x57028b18579ee440
    End
    RecordSource ="ExitOptionsQry"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    FilterOnLoad =255
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1002
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =645
                    Width =17295
                    Height =345
                    BackColor =15527148
                    Name ="Box48"
                    LayoutCachedTop =645
                    LayoutCachedWidth =17295
                    LayoutCachedHeight =990
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =345
                    Width =17295
                    Height =285
                    BorderColor =1643706
                    Name ="Box47"
                    LayoutCachedTop =345
                    LayoutCachedWidth =17295
                    LayoutCachedHeight =630
                    BackThemeColorIndex =4
                    BackShade =50.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =1032
                    Top =696
                    Width =1140
                    Height =276
                    FontSize =10
                    Name ="Label1"
                    Caption ="Unit Number"
                    GroupTable =1
                    LeftPadding =60
                    RightPadding =75
                    LayoutCachedLeft =1032
                    LayoutCachedTop =696
                    LayoutCachedWidth =2172
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =9660
                    Top =15
                    Width =29
                    Height =270
                    ColumnOrder =3
                    FontWeight =700
                    TabIndex =12
                    Name ="SchNo"
                    ControlSource ="SchNo"

                    LayoutCachedLeft =9660
                    LayoutCachedTop =15
                    LayoutCachedWidth =9689
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =9540
                    Top =15
                    Width =29
                    Height =270
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =11
                    ForeColor =12566463
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =9540
                    LayoutCachedTop =15
                    LayoutCachedWidth =9569
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =108
                    Top =696
                    Width =828
                    Height =276
                    FontSize =9
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label20"
                    Caption ="♦Asset"
                    GroupTable =1
                    LayoutCachedLeft =108
                    LayoutCachedTop =696
                    LayoutCachedWidth =936
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =36
                    Top =696
                    Width =0
                    Height =276
                    FontSize =9
                    ForeColor =5026082
                    Name ="Label21"
                    Caption ="SGrpSchID"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =36
                    LayoutCachedTop =696
                    LayoutCachedWidth =36
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9390
                    Top =15
                    Width =29
                    Height =270
                    ColumnOrder =1
                    FontWeight =700
                    TabIndex =10
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =9390
                    LayoutCachedTop =15
                    LayoutCachedWidth =9419
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5652
                    Top =696
                    Width =1080
                    Height =276
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =4
                    Name ="BuyOutDateCbo"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =5652
                    LayoutCachedTop =696
                    LayoutCachedWidth =6732
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10104
                    Top =696
                    Width =1212
                    Height =276
                    ColumnOrder =7
                    FontSize =10
                    TabIndex =7
                    Name ="IFinDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =10104
                    LayoutCachedTop =696
                    LayoutCachedWidth =11316
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =12792
                    Top =12
                    Width =1272
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="PayOffDateSubmit"
                    Caption ="▼Submit"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12792
                    LayoutCachedTop =12
                    LayoutCachedWidth =14064
                    LayoutCachedHeight =252
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =2
                    PressedThemeColorIndex =2
                    GroupTable =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6804
                    Top =696
                    Width =1080
                    Height =276
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =5
                    Name ="BuyOutPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =6804
                    LayoutCachedTop =696
                    LayoutCachedWidth =7884
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12792
                    Top =696
                    Width =1272
                    Height =276
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =9
                    Name ="IFAmountCbo"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =12792
                    LayoutCachedTop =696
                    LayoutCachedWidth =14064
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6840
                    Left =8736
                    Top =696
                    Width =1296
                    Height =276
                    ColumnOrder =6
                    FontSize =10
                    TabIndex =6
                    Name ="IFBank"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankId, Banks.BankShortName, Banks.BankName FROM Banks;"
                    ColumnWidths ="0;1800;5040"
                    StatusBarText ="bank used"
                    Tag ="Medit"
                    GroupTable =1
                    AllowValueListEdits =0

                    LayoutCachedLeft =8736
                    LayoutCachedTop =696
                    LayoutCachedWidth =10032
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =7956
                            Top =696
                            Width =708
                            Height =276
                            FontSize =10
                            Name ="Label6"
                            Caption =" MTHS"
                            GroupTable =1
                            LayoutCachedLeft =7956
                            LayoutCachedTop =696
                            LayoutCachedWidth =8664
                            LayoutCachedHeight =972
                            RowStart =2
                            RowEnd =2
                            ColumnStart =7
                            ColumnEnd =7
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Width =765
                    Height =210
                    ColumnOrder =0
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedWidth =765
                    LayoutCachedHeight =210
                End
                Begin CommandButton
                    OverlapFlags =87
                    TextFontCharSet =238
                    Left =1032
                    Top =12
                    Width =1140
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="send2Excel"
                    Caption ="Excel"
                    GroupTable =1
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =1032
                    LayoutCachedTop =12
                    LayoutCachedWidth =2172
                    LayoutCachedHeight =252
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =4754549
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    GroupTable =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =8736
                    Top =372
                    Width =1296
                    Height =288
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label46"
                    Caption ="Assignee"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8736
                    LayoutCachedTop =372
                    LayoutCachedWidth =10032
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =36
                    Top =372
                    Width =0
                    Height =288
                    Name ="EmptyCell68"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =36
                    LayoutCachedTop =372
                    LayoutCachedWidth =36
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =108
                    Top =372
                    Width =828
                    Height =288
                    Name ="EmptyCell69"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =108
                    LayoutCachedTop =372
                    LayoutCachedWidth =936
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =1032
                    Top =372
                    Width =1140
                    Height =288
                    Name ="EmptyCell70"
                    GroupTable =1
                    LeftPadding =60
                    TopPadding =0
                    RightPadding =75
                    BottomPadding =0
                    LayoutCachedLeft =1032
                    LayoutCachedTop =372
                    LayoutCachedWidth =2172
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2280
                    Top =372
                    Width =2148
                    Height =288
                    Name ="EmptyCell71"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2280
                    LayoutCachedTop =372
                    LayoutCachedWidth =4428
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7956
                    Top =372
                    Width =708
                    Height =288
                    Name ="EmptyCell75"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7956
                    LayoutCachedTop =372
                    LayoutCachedWidth =8664
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =6804
                    Top =372
                    Width =1080
                    Height =288
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label150"
                    Caption ="ExchangeIT"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6804
                    LayoutCachedTop =372
                    LayoutCachedWidth =7884
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =36
                    Top =12
                    Width =0
                    Name ="EmptyCell96"
                    GroupTable =1
                    TopPadding =15
                    RightPadding =38
                    BottomPadding =120
                    LayoutCachedLeft =36
                    LayoutCachedTop =12
                    LayoutCachedWidth =36
                    LayoutCachedHeight =252
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =108
                    Top =12
                    Width =828
                    Name ="EmptyCell97"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =108
                    LayoutCachedTop =12
                    LayoutCachedWidth =936
                    LayoutCachedHeight =252
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2280
                    Top =12
                    Width =2148
                    Name ="EmptyCell99"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =2280
                    LayoutCachedTop =12
                    LayoutCachedWidth =4428
                    LayoutCachedHeight =252
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =5652
                    Top =372
                    Width =1080
                    Height =288
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label187"
                    Caption ="Termation"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5652
                    LayoutCachedTop =372
                    LayoutCachedWidth =6732
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7956
                    Top =12
                    Width =708
                    Name ="EmptyCell103"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =7956
                    LayoutCachedTop =12
                    LayoutCachedWidth =8664
                    LayoutCachedHeight =252
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =12792
                    Top =372
                    Width =1272
                    Height =288
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label44"
                    Caption ="Rent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12792
                    LayoutCachedTop =372
                    LayoutCachedWidth =14064
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =10104
                    Top =372
                    Width =1212
                    Height =288
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label35"
                    Caption ="Assigned Date"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10104
                    LayoutCachedTop =372
                    LayoutCachedWidth =11316
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =11388
                    Top =372
                    Width =1332
                    Height =288
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label203"
                    Caption ="Notice Date"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11388
                    LayoutCachedTop =372
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =660
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
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11388
                    Top =696
                    Width =1332
                    Height =276
                    ColumnOrder =8
                    FontSize =10
                    TabIndex =8
                    Name ="IFAccountCBO"
                    GroupTable =1

                    LayoutCachedLeft =11388
                    LayoutCachedTop =696
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4500
                    Top =696
                    Width =1080
                    Height =276
                    FontSize =10
                    TabIndex =3
                    Name ="SellerInvCbo"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =4500
                    LayoutCachedTop =696
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =972
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =2280
                            Top =696
                            Width =2148
                            Height =276
                            FontSize =10
                            Name ="Label2"
                            Caption ="♦VIN"
                            GroupTable =1
                            LayoutCachedLeft =2280
                            LayoutCachedTop =696
                            LayoutCachedWidth =4428
                            LayoutCachedHeight =972
                            RowStart =2
                            RowEnd =2
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4500
                    Top =372
                    Width =1080
                    Height =288
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label271"
                    Caption ="BLCD"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4500
                    LayoutCachedTop =372
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =11388
                    Top =12
                    Width =1332
                    Name ="EmptyCell295"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =11388
                    LayoutCachedTop =12
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =252
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =10104
                    Top =12
                    Width =1212
                    Name ="EmptyCell296"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =10104
                    LayoutCachedTop =12
                    LayoutCachedWidth =11316
                    LayoutCachedHeight =252
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8736
                    Top =12
                    Width =1296
                    Name ="EmptyCell297"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =8736
                    LayoutCachedTop =12
                    LayoutCachedWidth =10032
                    LayoutCachedHeight =252
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6804
                    Top =12
                    Width =1080
                    Name ="EmptyCell298"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =6804
                    LayoutCachedTop =12
                    LayoutCachedWidth =7884
                    LayoutCachedHeight =252
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5652
                    Top =12
                    Width =1080
                    Name ="EmptyCell299"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =5652
                    LayoutCachedTop =12
                    LayoutCachedWidth =6732
                    LayoutCachedHeight =252
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4500
                    Top =12
                    Width =1080
                    Name ="EmptyCell300"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =4500
                    LayoutCachedTop =12
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =252
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =374
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =36
                    Top =36
                    Width =0
                    Height =276
                    FontSize =9
                    ForeColor =5026082
                    Name ="SGrpSchID"
                    ControlSource ="SchID"
                    GroupTable =1
                    RightPadding =38

                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =36
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1032
                    Top =36
                    Width =1140
                    Height =276
                    FontSize =10
                    TabIndex =2
                    Name ="UnitNum"
                    GroupTable =1
                    LeftPadding =60
                    RightPadding =75

                    LayoutCachedLeft =1032
                    LayoutCachedTop =36
                    LayoutCachedWidth =2172
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2280
                    Top =36
                    Width =2148
                    Height =276
                    FontSize =10
                    TabIndex =3
                    Name ="UnitVIN"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =2280
                    LayoutCachedTop =36
                    LayoutCachedWidth =4428
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5652
                    Top =36
                    Width =1080
                    Height =276
                    FontSize =10
                    TabIndex =5
                    BackColor =13237503
                    Name ="BODT"
                    ControlSource ="StartOrPurposedDate"
                    ValidationText ="BO Date is required"
                    GroupTable =1

                    LayoutCachedLeft =5652
                    LayoutCachedTop =36
                    LayoutCachedWidth =6732
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6804
                    Top =36
                    Width =1080
                    Height =276
                    FontSize =10
                    TabIndex =6
                    BackColor =13237503
                    Name ="BuyOutCost"
                    ValidationRule =">=0"
                    ValidationText ="Buyout Price can not be a neg number"
                    DefaultValue ="0"
                    GroupTable =1

                    LayoutCachedLeft =6804
                    LayoutCachedTop =36
                    LayoutCachedWidth =7884
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11388
                    Top =36
                    Width =1332
                    Height =276
                    FontSize =10
                    TabIndex =9
                    BackColor =16381933
                    Name ="NoticeDate"
                    ControlSource ="AssigneeNotified"
                    DefaultValue ="0"
                    GroupTable =1

                    LayoutCachedLeft =11388
                    LayoutCachedTop =36
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7956
                    Top =36
                    Width =708
                    Height =276
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    Name ="UnitStatus"
                    ControlSource ="TermMonths"
                    GroupTable =1

                    LayoutCachedLeft =7956
                    LayoutCachedTop =36
                    LayoutCachedWidth =8664
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =108
                    Top =36
                    Width =828
                    Height =276
                    FontSize =9
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =108
                    LayoutCachedTop =36
                    LayoutCachedWidth =936
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10104
                    Top =36
                    Width =1212
                    Height =276
                    FontSize =10
                    TabIndex =10
                    BackColor =16317438
                    Name ="IFDate"
                    GroupTable =1

                    LayoutCachedLeft =10104
                    LayoutCachedTop =36
                    LayoutCachedWidth =11316
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12792
                    Top =36
                    Width =1272
                    Height =276
                    FontSize =10
                    TabIndex =11
                    BackColor =16317438
                    Name ="Rent"
                    ControlSource ="Rent"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1

                    LayoutCachedLeft =12792
                    LayoutCachedTop =36
                    LayoutCachedWidth =14064
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =315
                    Width =18000
                    Height =59
                    TabIndex =12
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End

                    LayoutCachedTop =315
                    LayoutCachedWidth =18000
                    LayoutCachedHeight =374
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4500
                    Top =36
                    Width =1080
                    Height =276
                    FontSize =9
                    TabIndex =4
                    Name ="SellerInvoice"
                    ControlSource ="OrginalBLCD"
                    GroupTable =1

                    LayoutCachedLeft =4500
                    LayoutCachedTop =36
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8736
                    Top =36
                    Width =1296
                    Height =276
                    FontSize =10
                    TabIndex =8
                    BackColor =16317438
                    Name ="UnitIFBank"
                    StatusBarText ="bank used"
                    GroupTable =1

                    LayoutCachedLeft =8736
                    LayoutCachedTop =36
                    LayoutCachedWidth =10032
                    LayoutCachedHeight =312
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
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
' See "ExitOptions_subform.cls"
