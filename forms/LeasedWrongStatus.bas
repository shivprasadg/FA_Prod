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
    Width =20190
    DatasheetFontHeight =11
    ItemSuffix =68
    Left =1230
    Top =7020
    Right =19620
    Bottom =10920
    Filter ="(RemarketUnitStatusAudit.TargetSale<>0 Or RemarketUnitStatusAudit.TargetSale Is "
        "Null)"
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x24a025e10f86e440
    End
    RecordSource ="RemarketWrongStatusUnits"
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
        Begin FormHeader
            Height =999
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =1
                    BackStyle =1
                    OverlapFlags =93
                    Left =30
                    Top =600
                    Width =20160
                    Height =390
                    BackColor =15527148
                    BorderColor =-2147483617
                    Name ="Box22"
                    HorizontalAnchor =2
                    LayoutCachedLeft =30
                    LayoutCachedTop =600
                    LayoutCachedWidth =20190
                    LayoutCachedHeight =990
                End
                Begin Label
                    OverlapFlags =223
                    Left =7404
                    Top =660
                    Width =1656
                    Height =315
                    ForeColor =-2147483615
                    Name ="Label6"
                    Caption ="VIN Number:"
                    GroupTable =1
                    GridlineColor =4754549
                    LayoutCachedLeft =7404
                    LayoutCachedTop =660
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =975
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    Left =5496
                    Top =660
                    Width =768
                    Height =315
                    ForeColor =-2147483615
                    Name ="Label7"
                    Caption ="♦Asset:"
                    GroupTable =1
                    RightPadding =38
                    GridlineColor =4754549
                    LayoutCachedLeft =5496
                    LayoutCachedTop =660
                    LayoutCachedWidth =6264
                    LayoutCachedHeight =975
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    Left =9984
                    Top =660
                    Width =540
                    Height =315
                    ForeColor =-2147483615
                    Name ="Label8"
                    Caption ="Year:"
                    GroupTable =1
                    GridlineColor =4754549
                    LayoutCachedLeft =9984
                    LayoutCachedTop =660
                    LayoutCachedWidth =10524
                    LayoutCachedHeight =975
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    Left =18804
                    Top =660
                    Width =1152
                    Height =315
                    ForeColor =-2147483615
                    Name ="Label9"
                    Caption ="Inspected:"
                    GroupTable =1
                    GridlineColor =4754549
                    LayoutCachedLeft =18804
                    LayoutCachedTop =660
                    LayoutCachedWidth =19956
                    LayoutCachedHeight =975
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    Left =10584
                    Top =660
                    Width =2760
                    Height =315
                    ForeColor =-2147483615
                    Name ="Label10"
                    Caption ="Description:"
                    GroupTable =1
                    GridlineColor =4754549
                    LayoutCachedLeft =10584
                    LayoutCachedTop =660
                    LayoutCachedWidth =13344
                    LayoutCachedHeight =975
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =1
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =9132
                    Top =660
                    Width =780
                    Height =315
                    BackColor =14211288
                    ForeColor =-2147483615
                    Name ="Label12"
                    Caption ="Status:"
                    GroupTable =1
                    RightPadding =38
                    GridlineColor =4754549
                    LayoutCachedLeft =9132
                    LayoutCachedTop =660
                    LayoutCachedWidth =9912
                    LayoutCachedHeight =975
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    Left =2676
                    Top =660
                    Width =1536
                    Height =315
                    ForeColor =-2147483615
                    Name ="Label19"
                    Caption ="SCH:"
                    GroupTable =1
                    GridlineColor =4754549
                    LayoutCachedLeft =2676
                    LayoutCachedTop =660
                    LayoutCachedWidth =4212
                    LayoutCachedHeight =975
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    Left =13416
                    Top =660
                    Width =1740
                    Height =315
                    ForeColor =-2147483615
                    Name ="Label36"
                    Caption ="♦Buyer"
                    GroupTable =1
                    RightPadding =38
                    GridlineColor =4754549
                    LayoutCachedLeft =13416
                    LayoutCachedTop =660
                    LayoutCachedWidth =15156
                    LayoutCachedHeight =975
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    Left =36
                    Top =660
                    Width =2568
                    Height =315
                    ForeColor =-2147483615
                    Name ="Label38"
                    Caption ="♦Company"
                    GroupTable =1
                    RightPadding =38
                    GridlineColor =4754549
                    LayoutCachedLeft =36
                    LayoutCachedTop =660
                    LayoutCachedWidth =2604
                    LayoutCachedHeight =975
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    Left =4272
                    Top =660
                    Width =1152
                    Height =315
                    ForeColor =-2147483615
                    Name ="Label42"
                    Caption ="Group"
                    GroupTable =1
                    GridlineColor =4754549
                    LayoutCachedLeft =4272
                    LayoutCachedTop =660
                    LayoutCachedWidth =5424
                    LayoutCachedHeight =975
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =15228
                    Top =660
                    Width =1308
                    Height =315
                    ForeColor =-2147483615
                    Name ="Label46"
                    Caption ="Buy Out "
                    GroupTable =1
                    GridlineColor =4754549
                    LayoutCachedLeft =15228
                    LayoutCachedTop =660
                    LayoutCachedWidth =16536
                    LayoutCachedHeight =975
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =16596
                    Top =660
                    Width =972
                    Height =315
                    ForeColor =-2147483615
                    Name ="Label50"
                    Caption ="Sold Price"
                    GroupTable =1
                    GridlineColor =4754549
                    LayoutCachedLeft =16596
                    LayoutCachedTop =660
                    LayoutCachedWidth =17568
                    LayoutCachedHeight =975
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =18330
                    Width =225
                    Height =210
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedLeft =18330
                    LayoutCachedWidth =18555
                    LayoutCachedHeight =210
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =6336
                    Top =660
                    Width =1008
                    Height =315
                    BackColor =4754549
                    ForeColor =-2147483615
                    Name ="Label54"
                    Caption ="UnitNo:"
                    GroupTable =1
                    GridlineColor =4754549
                    LayoutCachedLeft =6336
                    LayoutCachedTop =660
                    LayoutCachedWidth =7344
                    LayoutCachedHeight =975
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =17628
                    Top =660
                    Width =1116
                    Height =315
                    ForeColor =-2147483615
                    Name ="Label58"
                    Caption ="InvPaid"
                    GroupTable =1
                    GridlineColor =4754549
                    LayoutCachedLeft =17628
                    LayoutCachedTop =660
                    LayoutCachedWidth =18744
                    LayoutCachedHeight =975
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =20016
                    Top =660
                    Width =0
                    Height =315
                    ForeColor =16777215
                    Name ="Label60"
                    Caption =" "
                    GroupTable =1
                    GridlineColor =4754549
                    LayoutCachedLeft =20016
                    LayoutCachedTop =660
                    LayoutCachedWidth =20016
                    LayoutCachedHeight =975
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =20076
                    Top =660
                    Width =0
                    Height =315
                    ForeColor =16777215
                    Name ="Label62"
                    Caption =" "
                    GroupTable =1
                    GridlineColor =4754549
                    LayoutCachedLeft =20076
                    LayoutCachedTop =660
                    LayoutCachedWidth =20076
                    LayoutCachedHeight =975
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =223
                    Left =150
                    Top =90
                    Width =12135
                    Height =795
                    FontSize =28
                    FontWeight =700
                    Name ="Label63"
                    Caption ="Remarketing Units with the 'Wrong' status"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =150
                    LayoutCachedTop =90
                    LayoutCachedWidth =12285
                    LayoutCachedHeight =885
                End
                Begin Label
                    OverlapFlags =215
                    Left =135
                    Top =75
                    Width =10860
                    Height =795
                    FontSize =28
                    FontWeight =700
                    ForeColor =14211288
                    Name ="Label64"
                    Caption ="Remarketing Units with the 'Wrong' status"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =135
                    LayoutCachedTop =75
                    LayoutCachedWidth =10995
                    LayoutCachedHeight =870
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =12660
                    Top =135
                    Width =3210
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label65"
                    Caption ="'double-click' a field with a '♦' to see details"
                    LayoutCachedLeft =12660
                    LayoutCachedTop =135
                    LayoutCachedWidth =15870
                    LayoutCachedHeight =330
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =12660
                    Top =315
                    Width =3210
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label66"
                    Caption ="'right-click'  on a field to establish a filter"
                    LayoutCachedLeft =12660
                    LayoutCachedTop =315
                    LayoutCachedWidth =15870
                    LayoutCachedHeight =510
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =16530
                    Top =195
                    Width =1275
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    ForeColor =4754549
                    Name ="ExcelExport"
                    Caption ="Excel Export"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =16530
                    LayoutCachedTop =195
                    LayoutCachedWidth =17805
                    LayoutCachedHeight =480
                End
            End
        End
        Begin Section
            Height =350
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =7404
                    Top =36
                    Width =1656
                    Height =276
                    ColumnWidth =2805
                    FontSize =9
                    TabIndex =5
                    Name ="VIN"
                    ControlSource ="VIN"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =7404
                    LayoutCachedTop =36
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5496
                    Top =36
                    Width =768
                    Height =276
                    ColumnWidth =1125
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    BackColor =15527148
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="AssetID"
                    ControlSource ="AssetID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =5496
                    LayoutCachedTop =36
                    LayoutCachedWidth =6264
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9984
                    Top =36
                    Width =540
                    Height =276
                    FontSize =9
                    TabIndex =7
                    Name ="MYear"
                    ControlSource ="MYear"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =9984
                    LayoutCachedTop =36
                    LayoutCachedWidth =10524
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =18804
                    Top =36
                    Width =1152
                    Height =276
                    FontSize =9
                    TabIndex =14
                    Name ="inspected"
                    ControlSource ="Inspected"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =18804
                    LayoutCachedTop =36
                    LayoutCachedWidth =19956
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =10584
                    Top =36
                    Width =2760
                    Height =276
                    FontSize =9
                    TabIndex =8
                    Name ="Description"
                    ControlSource ="Desc"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =10584
                    LayoutCachedTop =36
                    LayoutCachedWidth =13344
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9132
                    Top =36
                    Width =780
                    Height =276
                    ColumnWidth =1755
                    FontSize =9
                    TabIndex =6
                    BorderColor =-2147483617
                    Name ="Status"
                    ControlSource ="Status"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =9132
                    LayoutCachedTop =36
                    LayoutCachedWidth =9912
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2676
                    Top =36
                    Width =1536
                    Height =276
                    ColumnWidth =2145
                    FontSize =9
                    TabIndex =1
                    Name ="Schedule"
                    ControlSource ="Schedule"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =2676
                    LayoutCachedTop =36
                    LayoutCachedWidth =4212
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13416
                    Top =36
                    Width =1740
                    Height =276
                    FontSize =9
                    FontWeight =700
                    TabIndex =9
                    BackColor =15527148
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Buyer"
                    ControlSource ="Buyer"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =13416
                    LayoutCachedTop =36
                    LayoutCachedWidth =15156
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =36
                    Width =2568
                    Height =276
                    FontSize =9
                    FontWeight =700
                    BackColor =15527148
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="CompanyName"
                    ControlSource ="CompanyName"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =2604
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4272
                    Top =36
                    Width =1152
                    Height =276
                    FontSize =9
                    TabIndex =2
                    Name ="UnitGroup"
                    ControlSource ="UnitGroup"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =4272
                    LayoutCachedTop =36
                    LayoutCachedWidth =5424
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15228
                    Top =36
                    Width =1308
                    Height =276
                    FontSize =9
                    TabIndex =10
                    Name ="BuyOutDate"
                    ControlSource ="BuyOutDate"
                    Format ="Short Date"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =15228
                    LayoutCachedTop =36
                    LayoutCachedWidth =16536
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =16596
                    Top =36
                    Width =972
                    Height =276
                    FontSize =9
                    TabIndex =11
                    Name ="SoldPrice"
                    ControlSource ="SoldPrice"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =16596
                    LayoutCachedTop =36
                    LayoutCachedWidth =17568
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =315
                    Width =20160
                    Height =29
                    TabIndex =16
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0041007300730065007400490064005d003d005b0074007800740043007500 ,
                        0x7200720065006e0074005d0000000000
                    End

                    LayoutCachedTop =315
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =344
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00160000005b00 ,
                        0x41007300730065007400490064005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6336
                    Top =36
                    Width =1008
                    Height =276
                    FontSize =9
                    TabIndex =4
                    Name ="Unitnum"
                    ControlSource ="Unitnum"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =6336
                    LayoutCachedTop =36
                    LayoutCachedWidth =7344
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =17628
                    Top =36
                    Width =1116
                    Height =276
                    FontSize =9
                    TabIndex =12
                    Name ="InvoicePaidDate"
                    ControlSource ="InvoicePaidDate"
                    Format ="Short Date"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =17628
                    LayoutCachedTop =36
                    LayoutCachedWidth =18744
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =20016
                    Top =36
                    Width =0
                    Height =276
                    FontSize =9
                    TabIndex =13
                    Name ="BuyerId"
                    ControlSource ="BuyerId"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =20016
                    LayoutCachedTop =36
                    LayoutCachedWidth =20016
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =20076
                    Top =36
                    Width =0
                    Height =276
                    FontSize =9
                    TabIndex =15
                    Name ="clientID"
                    ControlSource ="clientID"
                    GroupTable =1
                    BottomPadding =38

                    LayoutCachedLeft =20076
                    LayoutCachedTop =36
                    LayoutCachedWidth =20076
                    LayoutCachedHeight =312
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
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
' See "LeasedWrongStatus.cls"
