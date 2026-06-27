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
    ItemSuffix =294
    Left =11415
    Top =5760
    Right =26475
    Bottom =18375
    Filter ="(((ReMarketingBuyOutUnits.UnitStatus=\"S\"))) And (ReMarketingBuyOutUnits.BuyDat"
        "eEffective Is Null)"
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x946eaeab510be540
    End
    RecordSource ="ReMarketingBuyOutUnits"
    OnCurrent ="[Event Procedure]"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1005
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =645
                    Width =17985
                    Height =345
                    BackColor =15527148
                    Name ="Box48"
                    LayoutCachedTop =645
                    LayoutCachedWidth =17985
                    LayoutCachedHeight =990
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =345
                    Width =18000
                    Height =285
                    BackColor =1643706
                    BorderColor =1643706
                    Name ="Box47"
                    LayoutCachedTop =345
                    LayoutCachedWidth =18000
                    LayoutCachedHeight =630
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =990
                    Top =690
                    Width =1140
                    Height =285
                    FontSize =10
                    Name ="Label1"
                    Caption ="Unit Number"
                    GroupTable =1
                    LayoutCachedLeft =990
                    LayoutCachedTop =690
                    LayoutCachedWidth =2130
                    LayoutCachedHeight =975
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
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9660
                    Top =15
                    Width =29
                    Height =270
                    ColumnOrder =4
                    FontWeight =700
                    TabIndex =20
                    Name ="SchNo"
                    ControlSource ="SchNo"

                    LayoutCachedLeft =9660
                    LayoutCachedTop =15
                    LayoutCachedWidth =9689
                    LayoutCachedHeight =285
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
                    ColumnOrder =3
                    FontWeight =700
                    TabIndex =19
                    ForeColor =12566463
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =9540
                    LayoutCachedTop =15
                    LayoutCachedWidth =9569
                    LayoutCachedHeight =285
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =105
                    Top =690
                    Width =825
                    Height =285
                    FontSize =9
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label20"
                    Caption ="♦Asset"
                    GroupTable =1
                    LayoutCachedLeft =105
                    LayoutCachedTop =690
                    LayoutCachedWidth =930
                    LayoutCachedHeight =975
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
                    Left =30
                    Top =690
                    Width =0
                    Height =285
                    FontSize =9
                    ForeColor =5026082
                    Name ="Label21"
                    Caption ="SGrpSchID"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =30
                    LayoutCachedTop =690
                    LayoutCachedWidth =30
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =5520
                    Top =15
                    Width =1185
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="BuyOutDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5520
                    LayoutCachedTop =15
                    LayoutCachedWidth =6705
                    LayoutCachedHeight =255
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =5
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =5
                    HoverTint =80.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9390
                    Top =15
                    Width =29
                    Height =270
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =18
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =9390
                    LayoutCachedTop =15
                    LayoutCachedWidth =9419
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5520
                    Top =690
                    Width =1185
                    Height =285
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =11
                    Name ="BuyOutDateCbo"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =5520
                    LayoutCachedTop =690
                    LayoutCachedWidth =6705
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =10845
                    Top =15
                    Width =1185
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="IFinDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10845
                    LayoutCachedTop =15
                    LayoutCachedWidth =12030
                    LayoutCachedHeight =255
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =5
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =5
                    HoverTint =80.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10845
                    Top =690
                    Width =1185
                    Height =285
                    ColumnOrder =8
                    FontSize =10
                    TabIndex =14
                    Name ="IFinDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =10845
                    LayoutCachedTop =690
                    LayoutCachedWidth =12030
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =15135
                    Top =15
                    Width =1185
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="PayOffDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15135
                    LayoutCachedTop =15
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =255
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =5
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =5
                    HoverTint =80.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15135
                    Top =690
                    Width =1185
                    Height =285
                    ColumnOrder =11
                    FontSize =10
                    TabIndex =17
                    Name ="IFPaidDateCbo"
                    Format ="Short Date"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =15135
                    LayoutCachedTop =690
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =6765
                    Top =15
                    Width =1215
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="BuyOutSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6765
                    LayoutCachedTop =15
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =255
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =5
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =5
                    HoverTint =80.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6765
                    Top =690
                    Width =1215
                    Height =285
                    ColumnOrder =6
                    FontSize =10
                    TabIndex =12
                    Name ="BuyOutPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =6765
                    LayoutCachedTop =690
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =13830
                    Top =15
                    Width =1245
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="IFAmountSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =13830
                    LayoutCachedTop =15
                    LayoutCachedWidth =15075
                    LayoutCachedHeight =255
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =5
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =5
                    HoverTint =80.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13830
                    Top =690
                    Width =1245
                    Height =285
                    ColumnOrder =10
                    FontSize =10
                    TabIndex =16
                    Name ="IFAmountCbo"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =13830
                    LayoutCachedTop =690
                    LayoutCachedWidth =15075
                    LayoutCachedHeight =975
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
                    Left =8925
                    Top =690
                    Width =1860
                    Height =285
                    ColumnOrder =7
                    FontSize =10
                    TabIndex =13
                    Name ="IFBank"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankId, Banks.BankShortName, Banks.BankName FROM Banks;"
                    ColumnWidths ="0;1800;5040"
                    StatusBarText ="bank used"
                    Tag ="Medit"
                    GroupTable =1
                    AllowValueListEdits =0

                    LayoutCachedLeft =8925
                    LayoutCachedTop =690
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =975
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
                            Left =8040
                            Top =690
                            Width =825
                            Height =285
                            FontSize =10
                            Name ="Label6"
                            Caption ="Status"
                            GroupTable =1
                            LayoutCachedLeft =8040
                            LayoutCachedTop =690
                            LayoutCachedWidth =8865
                            LayoutCachedHeight =975
                            RowStart =2
                            RowEnd =2
                            ColumnStart =7
                            ColumnEnd =7
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =8925
                    Top =15
                    Width =1860
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="IFBankSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8925
                    LayoutCachedTop =15
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =255
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =5
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =5
                    HoverTint =80.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Width =765
                    Height =210
                    ColumnOrder =1
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedWidth =765
                    LayoutCachedHeight =210
                End
                Begin CommandButton
                    OverlapFlags =95
                    TextFontCharSet =238
                    Left =990
                    Top =15
                    Width =1140
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="send2Excel"
                    Caption ="Excel"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120

                    LayoutCachedLeft =990
                    LayoutCachedTop =15
                    LayoutCachedWidth =2130
                    LayoutCachedHeight =255
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
                    PressedColor =4754549
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =8925
                    Top =375
                    Width =1860
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label46"
                    Caption =" Interim  Bank Name"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8925
                    LayoutCachedTop =375
                    LayoutCachedWidth =10785
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
                    Left =30
                    Top =375
                    Width =0
                    Height =285
                    Name ="EmptyCell68"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =30
                    LayoutCachedTop =375
                    LayoutCachedWidth =30
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =105
                    Top =375
                    Width =825
                    Height =285
                    Name ="EmptyCell69"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =105
                    LayoutCachedTop =375
                    LayoutCachedWidth =930
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =990
                    Top =375
                    Width =1140
                    Height =285
                    Name ="EmptyCell70"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =990
                    LayoutCachedTop =375
                    LayoutCachedWidth =2130
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2190
                    Top =375
                    Width =2070
                    Height =285
                    Name ="EmptyCell71"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2190
                    LayoutCachedTop =375
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8040
                    Top =375
                    Width =825
                    Height =285
                    Name ="EmptyCell75"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8040
                    LayoutCachedTop =375
                    LayoutCachedWidth =8865
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
                    Left =6765
                    Top =375
                    Width =1215
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label150"
                    Caption ="Buy Out Cost"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6765
                    LayoutCachedTop =375
                    LayoutCachedWidth =7980
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
                    Left =30
                    Top =15
                    Width =0
                    Name ="EmptyCell96"
                    GroupTable =1
                    TopPadding =15
                    RightPadding =38
                    BottomPadding =120
                    LayoutCachedLeft =30
                    LayoutCachedTop =15
                    LayoutCachedWidth =30
                    LayoutCachedHeight =255
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =105
                    Top =15
                    Width =825
                    Name ="EmptyCell97"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =105
                    LayoutCachedTop =15
                    LayoutCachedWidth =930
                    LayoutCachedHeight =255
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2190
                    Top =15
                    Width =2070
                    Name ="EmptyCell99"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =2190
                    LayoutCachedTop =15
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =255
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =5520
                    Top =375
                    Width =1185
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label187"
                    Caption ="Buy Out Date"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5520
                    LayoutCachedTop =375
                    LayoutCachedWidth =6705
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
                    Left =8040
                    Top =15
                    Width =825
                    Name ="EmptyCell103"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    LayoutCachedLeft =8040
                    LayoutCachedTop =15
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =255
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =15135
                    Top =375
                    Width =1185
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label41"
                    Caption ="Interim Paid"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15135
                    LayoutCachedTop =375
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =13830
                    Top =375
                    Width =1245
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label44"
                    Caption =" Interim Amt"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13830
                    LayoutCachedTop =375
                    LayoutCachedWidth =15075
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
                    Left =10845
                    Top =375
                    Width =1185
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label35"
                    Caption =" Interim Date"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10845
                    LayoutCachedTop =375
                    LayoutCachedWidth =12030
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
                    Left =12090
                    Top =375
                    Width =1680
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label203"
                    Caption ="Loan Number"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12090
                    LayoutCachedTop =375
                    LayoutCachedWidth =13770
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
                Begin CommandButton
                    OverlapFlags =87
                    Left =12090
                    Top =15
                    Width =1680
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="IFAcctNumberSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12090
                    LayoutCachedTop =15
                    LayoutCachedWidth =13770
                    LayoutCachedHeight =255
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =5
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =5
                    HoverTint =80.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-65
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12090
                    Top =690
                    Width =1680
                    Height =285
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =15
                    Name ="IFAccountCBO"
                    GroupTable =1

                    LayoutCachedLeft =12090
                    LayoutCachedTop =690
                    LayoutCachedWidth =13770
                    LayoutCachedHeight =975
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
                    Left =4320
                    Top =690
                    Width =1140
                    Height =285
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =10
                    Name ="BuyDateEffectiveCbo"
                    Format ="Short Date"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =4320
                    LayoutCachedTop =690
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =975
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
                            Left =2190
                            Top =690
                            Width =2070
                            Height =285
                            FontSize =10
                            Name ="Label2"
                            Caption ="VIN"
                            GroupTable =1
                            LayoutCachedLeft =2190
                            LayoutCachedTop =690
                            LayoutCachedWidth =4260
                            LayoutCachedHeight =975
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
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =4320
                    Top =375
                    Width =1140
                    Height =285
                    FontSize =10
                    BackColor =3312630
                    ForeColor =16777215
                    Name ="Label271"
                    Caption ="Effective Buy"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4320
                    LayoutCachedTop =375
                    LayoutCachedWidth =5460
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
                Begin CommandButton
                    OverlapFlags =87
                    Left =4320
                    Top =15
                    Width =1140
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="BuyDateEffectiveSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4320
                    LayoutCachedTop =15
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =255
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =3312630
                    BorderThemeColorIndex =1
                    HoverColor =3312630
                    PressedColor =3312630
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    Shadow =-1
                    QuickStyle =31
                    QuickStyleMask =-113
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =374
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =30
                    Top =30
                    Width =0
                    Height =255
                    FontSize =9
                    ForeColor =5026082
                    Name ="SGrpSchID"
                    ControlSource ="SchID"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =30
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =990
                    Top =30
                    Width =1140
                    Height =255
                    FontSize =10
                    TabIndex =2
                    Name ="UnitNum"
                    ControlSource ="UnitNum"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =990
                    LayoutCachedTop =30
                    LayoutCachedWidth =2130
                    LayoutCachedHeight =285
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
                    Left =2190
                    Top =30
                    Width =2070
                    Height =255
                    FontSize =10
                    TabIndex =3
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =2190
                    LayoutCachedTop =30
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =285
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
                    Left =5520
                    Top =30
                    Width =1185
                    Height =255
                    FontSize =10
                    TabIndex =5
                    BackColor =13237503
                    Name ="BODT"
                    ControlSource ="BuyOutDate"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =5520
                    LayoutCachedTop =30
                    LayoutCachedWidth =6705
                    LayoutCachedHeight =285
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
                    Left =6765
                    Top =30
                    Width =1215
                    Height =255
                    FontSize =10
                    TabIndex =6
                    BackColor =13237503
                    Name ="BuyOutCost"
                    ControlSource ="BuyOutCost"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =6765
                    LayoutCachedTop =30
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12090
                    Top =30
                    Width =1680
                    Height =255
                    FontSize =10
                    TabIndex =10
                    BackColor =16381933
                    Name ="IFLoanNumber"
                    ControlSource ="IFLoanNumber"
                    DefaultValue ="0"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =12090
                    LayoutCachedTop =30
                    LayoutCachedWidth =13770
                    LayoutCachedHeight =285
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
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8040
                    Top =30
                    Width =825
                    Height =255
                    FontSize =10
                    TabIndex =7
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =8040
                    LayoutCachedTop =30
                    LayoutCachedWidth =8865
                    LayoutCachedHeight =285
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
                    Left =105
                    Top =30
                    Width =825
                    Height =255
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
                    BottomPadding =38

                    LayoutCachedLeft =105
                    LayoutCachedTop =30
                    LayoutCachedWidth =930
                    LayoutCachedHeight =285
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
                    Left =10845
                    Top =30
                    Width =1185
                    Height =255
                    FontSize =10
                    TabIndex =9
                    BackColor =16317438
                    Name ="IFDate"
                    ControlSource ="IFDate"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =10845
                    LayoutCachedTop =30
                    LayoutCachedWidth =12030
                    LayoutCachedHeight =285
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
                    Left =13830
                    Top =30
                    Width =1245
                    Height =255
                    FontSize =10
                    TabIndex =11
                    BackColor =16317438
                    Name ="IFAmount"
                    ControlSource ="IFAmount"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =13830
                    LayoutCachedTop =30
                    LayoutCachedWidth =15075
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15135
                    Top =30
                    Width =1185
                    Height =255
                    FontSize =10
                    TabIndex =12
                    BackColor =16317438
                    Name ="IFPaidDate"
                    ControlSource ="IFPaidDate"
                    Format ="Short Date"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =15135
                    LayoutCachedTop =30
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6840
                    Left =8925
                    Top =30
                    Width =1860
                    Height =255
                    FontSize =10
                    TabIndex =8
                    BackColor =16317438
                    Name ="UnitIFBank"
                    ControlSource ="IFBankID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankId, Banks.BankShortName, Banks.BankName FROM Banks;"
                    ColumnWidths ="0;1800;5040"
                    StatusBarText ="bank used"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =38
                    AllowValueListEdits =0

                    LayoutCachedLeft =8925
                    LayoutCachedTop =30
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
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
                    TabIndex =13
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
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4320
                    Top =30
                    Width =1140
                    Height =255
                    FontSize =9
                    TabIndex =4
                    Name ="SellerInvoice"
                    ControlSource ="BuyDateEffective"
                    Format ="Short Date"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =4320
                    LayoutCachedTop =30
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin FormFooter
            Height =660
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6795
                    Top =30
                    Width =1350
                    Height =255
                    FontSize =10
                    BackColor =13237503
                    Name ="Text14"
                    ControlSource ="=Sum([BuyOutCost])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =6795
                    LayoutCachedTop =30
                    LayoutCachedWidth =8145
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =6015
                    Top =30
                    Width =750
                    Height =255
                    FontSize =10
                    FontWeight =700
                    Name ="Label15"
                    Caption ="Totals:"
                    LayoutCachedLeft =6015
                    LayoutCachedTop =30
                    LayoutCachedWidth =6765
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13875
                    Width =1245
                    Height =255
                    FontSize =10
                    TabIndex =1
                    BackColor =16317438
                    Name ="Text225"
                    ControlSource ="=Sum([IFAmount])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =13875
                    LayoutCachedWidth =15120
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
    End
End
CodeBehindForm
' See "ReMarketingBuyOut_subform.cls"
