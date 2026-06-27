Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =17784
    DatasheetFontHeight =11
    ItemSuffix =181
    Left =29610
    Top =12330
    Right =-18076
    Bottom =21960
    ShortcutMenuBar ="SortClip"
    OrderBy ="[BillingSubForm].[Billdt] DESC"
    RecSrcDt = Begin
        0xb60cae75f0f1e440
    End
    RecordSource ="SELECT Bills.*, MstrLease.MLNo, Schedule.SchID, Schedule.SchFSL, Schedule.SchNo,"
        " IIf([billtype]=\"PerDiem\",1,IIf([BillType]=\"rent\",2,3)) AS SRT, MstrLease.Cl"
        "ient, SchGrp.SGrpNO AS GroupNo FROM (MstrLease INNER JOIN (Schedule LEFT JOIN Bi"
        "lls ON Schedule.SchID = Bills.BillSchRef) ON MstrLease.MLID = Schedule.SchMLRef)"
        " LEFT JOIN SchGrp ON Bills.GroupId = SchGrp.SGrpID ORDER BY Bills.Billfrom DESC;"
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
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
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin ToggleButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =810
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =630
                    Width =17784
                    Height =180
                    Name ="Image32"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =630
                    LayoutCachedWidth =17784
                    LayoutCachedHeight =810
                    TabIndex =7
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =75
                    Top =390
                    Width =1500
                    Height =270
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Label17"
                    Caption ="Bill Type"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =75
                    LayoutCachedTop =390
                    LayoutCachedWidth =1575
                    LayoutCachedHeight =660
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =1575
                    Top =390
                    Width =1470
                    Height =270
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Label18"
                    Caption ="Bill No"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1575
                    LayoutCachedTop =390
                    LayoutCachedWidth =3045
                    LayoutCachedHeight =660
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =3045
                    Top =390
                    Width =1155
                    Height =270
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Label20"
                    Caption ="Bill Date"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3045
                    LayoutCachedTop =390
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =660
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =7305
                    Top =390
                    Width =1152
                    Height =270
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Label21"
                    Caption ="Bill From"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7305
                    LayoutCachedTop =390
                    LayoutCachedWidth =8457
                    LayoutCachedHeight =660
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =8460
                    Top =390
                    Width =1155
                    Height =270
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Label22"
                    Caption ="Bill To"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8460
                    LayoutCachedTop =390
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =660
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =9615
                    Top =390
                    Width =1155
                    Height =270
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Label23"
                    Caption ="Bill Due "
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9615
                    LayoutCachedTop =390
                    LayoutCachedWidth =10770
                    LayoutCachedHeight =660
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =11340
                    Top =390
                    Width =1185
                    Height =270
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Label24"
                    Caption ="Bill Sent"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11340
                    LayoutCachedTop =390
                    LayoutCachedWidth =12525
                    LayoutCachedHeight =660
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =6750
                    Top =390
                    Width =555
                    Height =270
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Label34"
                    Caption ="FSL"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6750
                    LayoutCachedTop =390
                    LayoutCachedWidth =7305
                    LayoutCachedHeight =660
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15405
                    Top =75
                    Width =1170
                    Height =390
                    TabIndex =5
                    Name ="AddBillBtn"
                    Caption ="Add Bill"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15405
                    LayoutCachedTop =75
                    LayoutCachedWidth =16575
                    LayoutCachedHeight =465
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackColor =15527148
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =105
                    Top =60
                    Width =10710
                    Height =315
                    ColumnOrder =0
                    FontSize =12
                    BorderColor =0
                    Name ="Header"
                    ControlSource ="=\"MLNo:<b> \" & UCase([Forms]![Billing]![sellease].[column](1) & \" </b>Sch:<b>"
                        " \" & [SchNo] & \" </b>Grp:<b> \" & [Forms]![Billing]![GroupList].[column](1))"
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedLeft =105
                    LayoutCachedTop =60
                    LayoutCachedWidth =10815
                    LayoutCachedHeight =375
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =10770
                    Top =390
                    Width =570
                    Height =270
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Label57"
                    Caption ="Receipt"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10770
                    LayoutCachedTop =390
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =660
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =4200
                    Top =390
                    Width =510
                    Height =270
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Label63"
                    Caption ="Rev"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4200
                    LayoutCachedTop =390
                    LayoutCachedWidth =4710
                    LayoutCachedHeight =660
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =4710
                    Top =390
                    Width =2040
                    Height =270
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Label64"
                    Caption ="Reason"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4710
                    LayoutCachedTop =390
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =660
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13425
                    Top =60
                    Width =600
                    Height =210
                    ColumnOrder =3
                    FontSize =8
                    TabIndex =2
                    Name ="SchNo"
                    ControlSource ="[schno]"

                    LayoutCachedLeft =13425
                    LayoutCachedTop =60
                    LayoutCachedWidth =14025
                    LayoutCachedHeight =270
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12825
                    Top =60
                    Width =600
                    Height =210
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =1
                    Name ="Lease"
                    ControlSource ="=[Forms]![Billing]![sellease]"

                    LayoutCachedLeft =12825
                    LayoutCachedTop =60
                    LayoutCachedWidth =13425
                    LayoutCachedHeight =270
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14025
                    Top =60
                    Width =600
                    Height =210
                    ColumnOrder =4
                    FontSize =8
                    TabIndex =3
                    Name ="ClientID"
                    ControlSource ="Client"

                    LayoutCachedLeft =14025
                    LayoutCachedTop =60
                    LayoutCachedWidth =14625
                    LayoutCachedHeight =270
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14625
                    Top =60
                    Width =600
                    Height =210
                    ColumnOrder =5
                    FontSize =8
                    TabIndex =4
                    Name ="MLNo"
                    ControlSource ="MLNo"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"

                    LayoutCachedLeft =14625
                    LayoutCachedTop =60
                    LayoutCachedWidth =15225
                    LayoutCachedHeight =270
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10815
                    Top =180
                    Width =900
                    Height =255
                    BackColor =15527148
                    Name ="Label88"
                    Caption ="Due On"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10815
                    LayoutCachedTop =180
                    LayoutCachedWidth =11715
                    LayoutCachedHeight =435
                    ColumnStart =9
                    ColumnEnd =9
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =555
                    Top =450
                    Width =225
                    Height =210
                    ColumnOrder =6
                    FontSize =8
                    TabIndex =6
                    Name ="txtCurrent"

                    LayoutCachedLeft =555
                    LayoutCachedTop =450
                    LayoutCachedWidth =780
                    LayoutCachedHeight =660
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =17055
                    Top =105
                    Width =645
                    Height =555
                    FontSize =10
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Label130"
                    Caption ="Strip\015\012Rent"
                    GroupTable =3
                    RightPadding =38
                    LayoutCachedLeft =17055
                    LayoutCachedTop =105
                    LayoutCachedWidth =17700
                    LayoutCachedHeight =660
                    LayoutGroup =2
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =1140
            Name ="Detail"
            AlternateBackColor =16316664
            Begin
                Begin EmptyCell
                    Left =6750
                    Top =30
                    Width =555
                    Height =300
                    Name ="EmptyCell132"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =6750
                    LayoutCachedTop =30
                    LayoutCachedWidth =7305
                    LayoutCachedHeight =330
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =17130
                    Top =195
                    Width =420
                    Height =330
                    BackColor =8210719
                    Name ="Box131"
                    LayoutCachedLeft =17130
                    LayoutCachedTop =195
                    LayoutCachedWidth =17550
                    LayoutCachedHeight =525
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =675
                    Top =245
                    Width =45
                    Height =645
                    TabIndex =5
                    ForeColor =16777215
                    Name ="Text99"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00420069006c006c00490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End

                    LayoutCachedLeft =675
                    LayoutCachedTop =245
                    LayoutCachedWidth =720
                    LayoutCachedHeight =890
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x420069006c006c00490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =675
                    Top =870
                    Width =15408
                    Height =44
                    TabIndex =23
                    ForeColor =16777215
                    Name ="Text97"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00420069006c006c00490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End

                    LayoutCachedLeft =675
                    LayoutCachedTop =870
                    LayoutCachedWidth =16083
                    LayoutCachedHeight =914
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x420069006c006c00490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1575
                    Top =30
                    Width =1470
                    Height =300
                    TabIndex =8
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    BorderColor =-2147483615
                    Name ="BillNo"
                    ControlSource ="BillNo"
                    StatusBarText ="from Quickbooks or bank if type is BankBill"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1575
                    LayoutCachedTop =30
                    LayoutCachedWidth =3045
                    LayoutCachedHeight =330
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3045
                    Top =30
                    Width =1155
                    Height =300
                    TabIndex =9
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    BorderColor =-2147483615
                    Name ="Billdt"
                    ControlSource ="Billdt"
                    Format ="mm/dd/yy"
                    StatusBarText ="date bill generated"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3045
                    LayoutCachedTop =30
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =330
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7305
                    Top =30
                    Width =1152
                    Height =300
                    TabIndex =12
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    BorderColor =-2147483615
                    Name ="Billfrom"
                    ControlSource ="Billfrom"
                    Format ="mm/dd/yy"
                    StatusBarText ="Leave bank to allow system to determine from date for 1st PD Bill"
                    ControlTipText ="Leave bank to allow system to determine from date for 1st PD Bill"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7305
                    LayoutCachedTop =30
                    LayoutCachedWidth =8457
                    LayoutCachedHeight =330
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8460
                    Top =30
                    Width =1155
                    Height =300
                    TabIndex =13
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    BorderColor =-2147483615
                    Name ="Billto"
                    ControlSource ="Billto"
                    Format ="mm/dd/yy"
                    StatusBarText ="end of billing period"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8460
                    LayoutCachedTop =30
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =330
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9615
                    Top =30
                    Width =1155
                    Height =300
                    TabIndex =14
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    BorderColor =-2147483615
                    Name ="BillDueDt"
                    ControlSource ="BillDueDt"
                    Format ="mm/dd/yy"
                    StatusBarText ="date pmt due (defaults to start of billing period)"
                    ControlTipText ="date pmt due (defaults to start of billing period)"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9615
                    LayoutCachedTop =30
                    LayoutCachedWidth =10770
                    LayoutCachedHeight =330
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11340
                    Top =30
                    Width =1185
                    Height =300
                    TabIndex =15
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    BorderColor =-2147483615
                    Name ="BillSent"
                    ControlSource ="BillSent"
                    StatusBarText ="date bill sent to Lessee"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11340
                    LayoutCachedTop =30
                    LayoutCachedWidth =12525
                    LayoutCachedHeight =330
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    IsHyperlink = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4170
                    Top =495
                    Width =7995
                    Height =270
                    FontSize =9
                    TabIndex =17
                    BorderColor =-2147483615
                    ForeColor =1279872587
                    Name ="BillLink"
                    ControlSource ="BillLink"
                    StatusBarText ="hyperlink to scan of billing package."
                    ShortcutMenuBar ="FormsCmd"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4170
                    LayoutCachedTop =495
                    LayoutCachedWidth =12165
                    LayoutCachedHeight =765
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =161
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15645
                    Top =570
                    Width =555
                    Height =225
                    FontSize =10
                    FontWeight =700
                    TabIndex =21
                    ForeColor =-2147483615
                    Name ="BillID"
                    ControlSource ="BillID"

                    LayoutCachedLeft =15645
                    LayoutCachedTop =570
                    LayoutCachedWidth =16200
                    LayoutCachedHeight =795
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =161
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14295
                    Top =570
                    Width =1320
                    Height =225
                    FontSize =9
                    TabIndex =20
                    BorderColor =0
                    ForeColor =5026082
                    Name ="BillSchRef"
                    ControlSource ="BillSchRef"
                    StatusBarText ="link to schedule table"

                    LayoutCachedLeft =14295
                    LayoutCachedTop =570
                    LayoutCachedWidth =15615
                    LayoutCachedHeight =795
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =6900
                    Top =72
                    Width =240
                    Height =300
                    BorderColor =-2147483615
                    Name ="SchFSL"
                    ControlSource ="SchFSL"
                    StatusBarText ="FSL Lease"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6900
                    LayoutCachedTop =72
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =372
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =14400
                    Top =195
                    Width =900
                    TabIndex =3
                    Name ="AddBAdjBtn"
                    Caption ="Add Adj."
                    LeftPadding =120
                    TopPadding =105
                    RightPadding =135
                    BottomPadding =165
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="BilAdjForm"
                            Argument ="0"
                            Argument =""
                            Argument ="=\"[BillAdjBillref]=\" & [BillID]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"AddBAdjBtn\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argumen"
                                "t Name=\"FormName\">BilAdjForm</A"
                        End
                        Begin
                            Comment ="_AXL:rgument><Argument Name=\"WhereCondition\">=\"[BillAdjBillref]=\" &amp; [Bil"
                                "lID]</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =14400
                    LayoutCachedTop =195
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =555
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =8872517
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =11148
                    Top =72
                    Width =345
                    Height =300
                    TabIndex =1
                    BorderColor =-2147483615
                    Name ="Check56"
                    ControlSource ="DueonRec"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11148
                    LayoutCachedTop =72
                    LayoutCachedWidth =11493
                    LayoutCachedHeight =372
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                End
                Begin TextBox
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4200
                    Top =30
                    Width =510
                    Height =300
                    TabIndex =10
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    BorderColor =-2147483615
                    Name ="RevNo"
                    ControlSource ="RevNo"
                    StatusBarText ="Revision Number"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4200
                    LayoutCachedTop =30
                    LayoutCachedWidth =4710
                    LayoutCachedHeight =330
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4710
                    Top =30
                    Width =2040
                    Height =300
                    TabIndex =11
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    BorderColor =-2147483615
                    Name ="RevReason"
                    ControlSource ="RevReason"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4710
                    LayoutCachedTop =30
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =330
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13200
                    Top =195
                    Width =1140
                    TabIndex =2
                    Name ="ReviewSchedule"
                    Caption ="Sch Form"
                    OnClick ="[Event Procedure]"
                    LeftPadding =120
                    TopPadding =105
                    RightPadding =135
                    BottomPadding =165

                    LayoutCachedLeft =13200
                    LayoutCachedTop =195
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =555
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =8872517
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =3195
                    Top =480
                    Width =885
                    Height =300
                    FontSize =9
                    FontWeight =700
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    Name ="Label25"
                    Caption ="Bill PDF:"
                    LayoutCachedLeft =3195
                    LayoutCachedTop =480
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =780
                    ForeThemeColorIndex =4
                End
                Begin CommandButton
                    OverlapFlags =127
                    Left =16080
                    Top =180
                    Width =855
                    TabIndex =4
                    Name ="Runbill"
                    Caption ="Run Bill"
                    OnClick ="[Event Procedure]"
                    LeftPadding =120
                    TopPadding =105
                    RightPadding =135
                    BottomPadding =165

                    LayoutCachedLeft =16080
                    LayoutCachedTop =180
                    LayoutCachedWidth =16935
                    LayoutCachedHeight =540
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1530
                    Top =510
                    Width =885
                    FontSize =9
                    TabIndex =19
                    BorderColor =0
                    ForeColor =12566463
                    Name ="GroupId"
                    ControlSource ="GroupId"

                    LayoutCachedLeft =1530
                    LayoutCachedTop =510
                    LayoutCachedWidth =2415
                    LayoutCachedHeight =750
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =120
                    Top =495
                    Width =1245
                    FontSize =9
                    FontWeight =700
                    TabIndex =16
                    BorderColor =-2147483615
                    ForeColor =8210719
                    Name ="GroupBill"
                    ControlSource ="=IIf(IsNull([GroupId]),\"Sch Bill\",\"Group \" & [GroupNo])"

                    LayoutCachedLeft =120
                    LayoutCachedTop =495
                    LayoutCachedWidth =1365
                    LayoutCachedHeight =735
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =16080
                    Top =765
                    Width =855
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =22
                    ForeColor =2366701
                    Name ="DeleteMe"
                    Caption ="DELETE"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =16080
                    LayoutCachedTop =765
                    LayoutCachedWidth =16935
                    LayoutCachedHeight =1005
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =1643706
                    HoverColor =1643706
                    PressedColor =0
                    QuickStyle =22
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12255
                    Top =495
                    Width =870
                    Height =210
                    FontSize =8
                    TabIndex =18
                    Name ="Makelink"
                    Caption ="Link Bill"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =12255
                    LayoutCachedTop =495
                    LayoutCachedWidth =13125
                    LayoutCachedHeight =705
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =4
                    BackShade =75.0
                    HoverColor =4754549
                    PressedColor =0
                    QuickStyle =33
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =75
                    Top =30
                    Width =1500
                    Height =300
                    TabIndex =7
                    BorderColor =-2147483615
                    Name ="BillType"
                    ControlSource ="BillType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_BillType.BillType, Lookup_BillType.BillTypeID FROM Lookup_BillType"
                        ";"
                    ColumnWidths ="1440;0"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Bill Type is required"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =75
                    LayoutCachedTop =30
                    LayoutCachedWidth =1575
                    LayoutCachedHeight =330
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GroupTable =1
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =17265
                    Top =270
                    Width =315
                    Height =300
                    TabIndex =6
                    BorderColor =-2147483615
                    Name ="UseSchRent"
                    ControlSource ="UseSchRent"
                    BottomPadding =38

                    LayoutCachedLeft =17265
                    LayoutCachedTop =270
                    LayoutCachedWidth =17580
                    LayoutCachedHeight =570
                End
                Begin EmptyCell
                    Left =10770
                    Top =30
                    Width =570
                    Height =300
                    Name ="EmptyCell138"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =0
                    LayoutCachedLeft =10770
                    LayoutCachedTop =30
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =330
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =119
                    Left =15360
                    Top =180
                    Width =660
                    TabIndex =24
                    Name ="InvoiceToExcel"
                    Caption ="Excel"
                    OnClick ="[Event Procedure]"
                    GroupTable =4
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =15360
                    LayoutCachedTop =180
                    LayoutCachedWidth =16020
                    LayoutCachedHeight =540
                    LayoutGroup =3
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =9
                    BackTint =40.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    PressedForeThemeColorIndex =2
                    GroupTable =4
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13560
                    Top =570
                    Width =720
                    Height =225
                    FontSize =9
                    TabIndex =25
                    BorderColor =0
                    ForeColor =5026082
                    Name ="BillTypeID"
                    ControlSource ="BillTypeID"
                    StatusBarText ="link to schedule table"

                    LayoutCachedLeft =13560
                    LayoutCachedTop =570
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =795
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
' See "BillingSubForm.cls"
