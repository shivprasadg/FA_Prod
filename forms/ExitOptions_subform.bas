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
            Height =990
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
                    Left =1020
                    Top =690
                    Width =1140
                    Height =270
                    FontSize =10
                    Name ="Label1"
                    Caption ="Unit Number"
                    GroupTable =1
                    LeftPadding =60
                    RightPadding =75
                    LayoutCachedLeft =1020
                    LayoutCachedTop =690
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =960
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =2265
                    Top =690
                    Width =2145
                    Height =270
                    FontSize =10
                    Name ="Label2"
                    Caption ="♦VIN"
                    GroupTable =1
                    LayoutCachedLeft =2265
                    LayoutCachedTop =690
                    LayoutCachedWidth =4410
                    LayoutCachedHeight =960
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7890
                    Top =690
                    Width =705
                    Height =270
                    FontSize =10
                    Name ="Label6"
                    Caption =" MTHS"
                    GroupTable =1
                    LayoutCachedLeft =7890
                    LayoutCachedTop =690
                    LayoutCachedWidth =8595
                    LayoutCachedHeight =960
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
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
                    Left =105
                    Top =690
                    Width =825
                    Height =270
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
                    LayoutCachedHeight =960
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
                    Height =270
                    FontSize =9
                    ForeColor =5026082
                    Name ="Label21"
                    Caption ="SGrpSchID"
                    GroupTable =1
                    RightPadding =38
                    LayoutCachedLeft =30
                    LayoutCachedTop =690
                    LayoutCachedWidth =30
                    LayoutCachedHeight =960
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
                    Left =5610
                    Top =690
                    Width =1080
                    Height =270
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =4
                    Name ="BuyOutDateCbo"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =5610
                    LayoutCachedTop =690
                    LayoutCachedWidth =6690
                    LayoutCachedHeight =960
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
                    Left =10005
                    Top =690
                    Width =1215
                    Height =270
                    ColumnOrder =7
                    FontSize =10
                    TabIndex =7
                    Name ="IFinDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =10005
                    LayoutCachedTop =690
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =960
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =12675
                    Top =15
                    Width =1275
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

                    LayoutCachedLeft =12675
                    LayoutCachedTop =15
                    LayoutCachedWidth =13950
                    LayoutCachedHeight =255
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
                    Left =6750
                    Top =690
                    Width =1080
                    Height =270
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =5
                    Name ="BuyOutPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =6750
                    LayoutCachedTop =690
                    LayoutCachedWidth =7830
                    LayoutCachedHeight =960
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
                    Left =12675
                    Top =690
                    Width =1275
                    Height =270
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =9
                    Name ="IFAmountCbo"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =12675
                    LayoutCachedTop =690
                    LayoutCachedWidth =13950
                    LayoutCachedHeight =960
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
                    Left =8655
                    Top =690
                    Width =1290
                    Height =270
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

                    LayoutCachedLeft =8655
                    LayoutCachedTop =690
                    LayoutCachedWidth =9945
                    LayoutCachedHeight =960
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
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
                    Left =1020
                    Top =15
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

                    LayoutCachedLeft =1020
                    LayoutCachedTop =15
                    LayoutCachedWidth =2160
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
                    Left =8655
                    Top =375
                    Width =1290
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label46"
                    Caption ="Assignee"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8655
                    LayoutCachedTop =375
                    LayoutCachedWidth =9945
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
                    GridlineColor =10921638
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
                    GridlineColor =10921638
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
                    Left =1020
                    Top =375
                    Width =1140
                    Height =285
                    Name ="EmptyCell70"
                    GroupTable =1
                    LeftPadding =60
                    TopPadding =0
                    RightPadding =75
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =1020
                    LayoutCachedTop =375
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =2265
                    Top =375
                    Width =2145
                    Height =285
                    Name ="EmptyCell71"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2265
                    LayoutCachedTop =375
                    LayoutCachedWidth =4410
                    LayoutCachedHeight =660
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7890
                    Top =375
                    Width =705
                    Height =285
                    Name ="EmptyCell75"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =7890
                    LayoutCachedTop =375
                    LayoutCachedWidth =8595
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
                    Left =6750
                    Top =375
                    Width =1080
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label150"
                    Caption ="ExchangeIT"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6750
                    LayoutCachedTop =375
                    LayoutCachedWidth =7830
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
                    GridlineColor =10921638
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
                    GridlineColor =10921638
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
                    Left =2265
                    Top =15
                    Width =2145
                    Name ="EmptyCell99"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    GridlineColor =10921638
                    LayoutCachedLeft =2265
                    LayoutCachedTop =15
                    LayoutCachedWidth =4410
                    LayoutCachedHeight =255
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =5610
                    Top =375
                    Width =1080
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label187"
                    Caption ="Termation"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5610
                    LayoutCachedTop =375
                    LayoutCachedWidth =6690
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
                    Left =7890
                    Top =15
                    Width =705
                    Name ="EmptyCell103"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    GridlineColor =10921638
                    LayoutCachedLeft =7890
                    LayoutCachedTop =15
                    LayoutCachedWidth =8595
                    LayoutCachedHeight =255
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =12675
                    Top =375
                    Width =1275
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label44"
                    Caption ="Rent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12675
                    LayoutCachedTop =375
                    LayoutCachedWidth =13950
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
                    Left =10005
                    Top =375
                    Width =1215
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label35"
                    Caption ="Assigned Date"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10005
                    LayoutCachedTop =375
                    LayoutCachedWidth =11220
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
                    Left =11280
                    Top =375
                    Width =1335
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label203"
                    Caption ="Notice Date"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11280
                    LayoutCachedTop =375
                    LayoutCachedWidth =12615
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
                    Left =11280
                    Top =690
                    Width =1335
                    Height =270
                    ColumnOrder =8
                    FontSize =10
                    TabIndex =8
                    Name ="IFAccountCBO"
                    GroupTable =1

                    LayoutCachedLeft =11280
                    LayoutCachedTop =690
                    LayoutCachedWidth =12615
                    LayoutCachedHeight =960
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
                    Left =4470
                    Top =690
                    Width =1080
                    Height =270
                    FontSize =10
                    TabIndex =3
                    Name ="SellerInvCbo"
                    Tag ="Medit"
                    GroupTable =1

                    LayoutCachedLeft =4470
                    LayoutCachedTop =690
                    LayoutCachedWidth =5550
                    LayoutCachedHeight =960
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4470
                    Top =375
                    Width =1080
                    Height =285
                    FontSize =10
                    BackColor =1643706
                    ForeColor =16777215
                    Name ="Label271"
                    Caption ="BLCD"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4470
                    LayoutCachedTop =375
                    LayoutCachedWidth =5550
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
                    Left =11280
                    Top =15
                    Width =1335
                    Name ="EmptyCell295"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    GridlineColor =10921638
                    LayoutCachedLeft =11280
                    LayoutCachedTop =15
                    LayoutCachedWidth =12615
                    LayoutCachedHeight =255
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =10005
                    Top =15
                    Width =1215
                    Name ="EmptyCell296"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    GridlineColor =10921638
                    LayoutCachedLeft =10005
                    LayoutCachedTop =15
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =255
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8655
                    Top =15
                    Width =1290
                    Name ="EmptyCell297"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    GridlineColor =10921638
                    LayoutCachedLeft =8655
                    LayoutCachedTop =15
                    LayoutCachedWidth =9945
                    LayoutCachedHeight =255
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =6750
                    Top =15
                    Width =1080
                    Name ="EmptyCell298"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    GridlineColor =10921638
                    LayoutCachedLeft =6750
                    LayoutCachedTop =15
                    LayoutCachedWidth =7830
                    LayoutCachedHeight =255
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =5610
                    Top =15
                    Width =1080
                    Name ="EmptyCell299"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    GridlineColor =10921638
                    LayoutCachedLeft =5610
                    LayoutCachedTop =15
                    LayoutCachedWidth =6690
                    LayoutCachedHeight =255
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4470
                    Top =15
                    Width =1080
                    Name ="EmptyCell300"
                    GroupTable =1
                    TopPadding =15
                    BottomPadding =120
                    GridlineColor =10921638
                    LayoutCachedLeft =4470
                    LayoutCachedTop =15
                    LayoutCachedWidth =5550
                    LayoutCachedHeight =255
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
                    Left =30
                    Top =30
                    Width =0
                    Height =270
                    FontSize =9
                    ForeColor =5026082
                    Name ="SGrpSchID"
                    ControlSource ="SchID"
                    GroupTable =1
                    RightPadding =38

                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =30
                    LayoutCachedHeight =300
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
                    Left =1020
                    Top =30
                    Width =1140
                    Height =270
                    FontSize =10
                    TabIndex =2
                    Name ="UnitNum"
                    GroupTable =1
                    LeftPadding =60
                    RightPadding =75

                    LayoutCachedLeft =1020
                    LayoutCachedTop =30
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =300
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
                    Left =2265
                    Top =30
                    Width =2145
                    Height =270
                    FontSize =10
                    TabIndex =3
                    Name ="UnitVIN"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1

                    LayoutCachedLeft =2265
                    LayoutCachedTop =30
                    LayoutCachedWidth =4410
                    LayoutCachedHeight =300
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
                    Left =5610
                    Top =30
                    Width =1080
                    Height =270
                    FontSize =10
                    TabIndex =5
                    BackColor =13237503
                    Name ="BODT"
                    ControlSource ="StartOrPurposedDate"
                    ValidationText ="BO Date is required"
                    GroupTable =1

                    LayoutCachedLeft =5610
                    LayoutCachedTop =30
                    LayoutCachedWidth =6690
                    LayoutCachedHeight =300
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
                    Left =6750
                    Top =30
                    Width =1080
                    Height =270
                    FontSize =10
                    TabIndex =6
                    BackColor =13237503
                    Name ="BuyOutCost"
                    ValidationRule =">=0"
                    ValidationText ="Buyout Price can not be a neg number"
                    DefaultValue ="0"
                    GroupTable =1

                    LayoutCachedLeft =6750
                    LayoutCachedTop =30
                    LayoutCachedWidth =7830
                    LayoutCachedHeight =300
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
                    Left =11280
                    Top =30
                    Width =1335
                    Height =270
                    FontSize =10
                    TabIndex =9
                    BackColor =16381933
                    Name ="NoticeDate"
                    ControlSource ="AssigneeNotified"
                    DefaultValue ="0"
                    GroupTable =1

                    LayoutCachedLeft =11280
                    LayoutCachedTop =30
                    LayoutCachedWidth =12615
                    LayoutCachedHeight =300
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
                    Left =7890
                    Top =30
                    Width =705
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    Name ="UnitStatus"
                    ControlSource ="TermMonths"
                    GroupTable =1

                    LayoutCachedLeft =7890
                    LayoutCachedTop =30
                    LayoutCachedWidth =8595
                    LayoutCachedHeight =300
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
                    Height =270
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

                    LayoutCachedLeft =105
                    LayoutCachedTop =30
                    LayoutCachedWidth =930
                    LayoutCachedHeight =300
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
                    Left =10005
                    Top =30
                    Width =1215
                    Height =270
                    FontSize =10
                    TabIndex =10
                    BackColor =16317438
                    Name ="IFDate"
                    GroupTable =1

                    LayoutCachedLeft =10005
                    LayoutCachedTop =30
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =300
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
                    Left =12675
                    Top =30
                    Width =1275
                    Height =270
                    FontSize =10
                    TabIndex =11
                    BackColor =16317438
                    Name ="Rent"
                    ControlSource ="Rent"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1

                    LayoutCachedLeft =12675
                    LayoutCachedTop =30
                    LayoutCachedWidth =13950
                    LayoutCachedHeight =300
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
                    Left =4470
                    Top =30
                    Width =1080
                    Height =270
                    FontSize =9
                    TabIndex =4
                    Name ="SellerInvoice"
                    ControlSource ="OrginalBLCD"
                    GroupTable =1

                    LayoutCachedLeft =4470
                    LayoutCachedTop =30
                    LayoutCachedWidth =5550
                    LayoutCachedHeight =300
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
                    Left =8655
                    Top =30
                    Width =1290
                    Height =270
                    FontSize =10
                    TabIndex =8
                    BackColor =16317438
                    Name ="UnitIFBank"
                    StatusBarText ="bank used"
                    GroupTable =1

                    LayoutCachedLeft =8655
                    LayoutCachedTop =30
                    LayoutCachedWidth =9945
                    LayoutCachedHeight =300
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
