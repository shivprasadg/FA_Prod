Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =17955
    DatasheetFontHeight =11
    ItemSuffix =42
    Left =3510
    Top =6990
    Right =22065
    Bottom =16050
    ShortcutMenuBar ="SortClip"
    OrderBy ="[BuyerOrderUnits].[AssetID]"
    RecSrcDt = Begin
        0x9ef60b5ff7b9e540
    End
    RecordSource ="BuyerOrderUnits"
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
            BorderColor =12632256
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =600
            Name ="FormHeader"
            Begin
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    Left =36
                    Top =300
                    Width =876
                    Height =255
                    FontSize =10
                    FontWeight =700
                    BackColor =15527148
                    Name ="Label0"
                    Caption ="♦Asset:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =36
                    LayoutCachedTop =300
                    LayoutCachedWidth =912
                    LayoutCachedHeight =555
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    Left =984
                    Top =300
                    Width =1908
                    Height =255
                    FontSize =10
                    BackColor =15527148
                    ForeColor =-2147483615
                    Name ="Label4"
                    Caption ="VIN:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =984
                    LayoutCachedTop =300
                    LayoutCachedWidth =2892
                    LayoutCachedHeight =555
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =3612
                    Top =300
                    Width =1560
                    Height =255
                    FontSize =10
                    BackColor =15527148
                    ForeColor =-2147483615
                    Name ="Label5"
                    Caption ="Make:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =3612
                    LayoutCachedTop =300
                    LayoutCachedWidth =5172
                    LayoutCachedHeight =555
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =5244
                    Top =300
                    Width =1908
                    Height =255
                    FontSize =10
                    BackColor =15527148
                    ForeColor =-2147483615
                    Name ="Label6"
                    Caption ="Model:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =5244
                    LayoutCachedTop =300
                    LayoutCachedWidth =7152
                    LayoutCachedHeight =555
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =7224
                    Top =300
                    Width =1752
                    Height =255
                    FontSize =10
                    BackColor =15527148
                    ForeColor =-2147483615
                    Name ="Label7"
                    Caption ="Type:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =7224
                    LayoutCachedTop =300
                    LayoutCachedWidth =8976
                    LayoutCachedHeight =555
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =2964
                    Top =300
                    Width =576
                    Height =255
                    FontSize =10
                    BackColor =15527148
                    ForeColor =-2147483615
                    Name ="Label11"
                    Caption ="YEAR:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =2964
                    LayoutCachedTop =300
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =555
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =9048
                    Top =300
                    Width =1332
                    Height =255
                    FontSize =10
                    BackColor =15527148
                    ForeColor =-2147483615
                    Name ="Label13"
                    Caption ="Sale Price:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =9048
                    LayoutCachedTop =300
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =555
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    Left =12720
                    Top =300
                    Width =1065
                    Height =255
                    FontSize =10
                    FontWeight =700
                    BorderColor =11830108
                    ForeColor =-2147483615
                    Name ="Label15"
                    Caption ="Requested:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =12720
                    LayoutCachedTop =300
                    LayoutCachedWidth =13785
                    LayoutCachedHeight =555
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    Left =13860
                    Top =300
                    Width =1065
                    Height =255
                    FontSize =10
                    FontWeight =700
                    BorderColor =11830108
                    ForeColor =-2147483615
                    Name ="Label17"
                    Caption ="Sent:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =13860
                    LayoutCachedTop =300
                    LayoutCachedWidth =14925
                    LayoutCachedHeight =555
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =15000
                    Top =300
                    Width =1065
                    Height =255
                    FontSize =10
                    BackColor =15527148
                    ForeColor =-2147483615
                    Name ="Label18"
                    Caption ="Invoice Paid:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =15000
                    LayoutCachedTop =300
                    LayoutCachedWidth =16065
                    LayoutCachedHeight =555
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =16140
                    Top =300
                    Width =600
                    Height =255
                    FontSize =10
                    BackColor =15527148
                    ForeColor =-2147483615
                    Name ="Label21"
                    Caption ="Status"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =16140
                    LayoutCachedTop =300
                    LayoutCachedWidth =16740
                    LayoutCachedHeight =555
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    Left =11580
                    Top =300
                    Width =1065
                    Height =255
                    FontSize =10
                    FontWeight =700
                    BorderColor =26163
                    ForeColor =26163
                    Name ="Label28"
                    Caption ="Received"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =11580
                    LayoutCachedTop =300
                    LayoutCachedWidth =12645
                    LayoutCachedHeight =555
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    SpecialEffect =3
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    Left =16845
                    Top =300
                    Width =465
                    Height =255
                    FontSize =10
                    BackColor =15527148
                    ForeColor =-2147483615
                    Name ="Label32"
                    Caption =" BOS"
                    LayoutCachedLeft =16845
                    LayoutCachedTop =300
                    LayoutCachedWidth =17310
                    LayoutCachedHeight =555
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    Left =10452
                    Top =300
                    Width =1056
                    Height =255
                    FontSize =10
                    FontWeight =700
                    BorderColor =26163
                    ForeColor =26163
                    Name ="Label27"
                    Caption =" Amount:"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38
                    LayoutCachedLeft =10452
                    LayoutCachedTop =300
                    LayoutCachedWidth =11508
                    LayoutCachedHeight =555
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15150
                    Width =855
                    Height =165
                    ColumnOrder =0
                    TabIndex =1
                    Name ="BuyerRef"
                    ControlSource ="BuyerID"

                    LayoutCachedLeft =15150
                    LayoutCachedWidth =16005
                    LayoutCachedHeight =165
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =186
                    TextAlign =2
                    Left =10500
                    Top =120
                    Width =2040
                    Height =225
                    FontSize =8
                    FontWeight =700
                    ForeColor =26163
                    Name ="Label107"
                    Caption ="┌─── Deposit ────┐"
                    LayoutCachedLeft =10500
                    LayoutCachedTop =120
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =345
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =186
                    TextAlign =2
                    Left =12705
                    Top =120
                    Width =2175
                    Height =225
                    FontSize =8
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label34"
                    Caption ="┌─── Invoice ────┐"
                    LayoutCachedLeft =12705
                    LayoutCachedTop =120
                    LayoutCachedWidth =14880
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =735
                    Width =525
                    Height =285
                    ColumnOrder =1
                    FontWeight =700
                    Name ="UnitCount"
                    ControlSource ="=Count([AssetID])"

                    LayoutCachedLeft =735
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =16020
                    Width =855
                    Height =165
                    ColumnOrder =2
                    TabIndex =2
                    Name ="InvoiceRequestNum"
                    ControlSource ="InvoiceRequestNum"

                    LayoutCachedLeft =16020
                    LayoutCachedWidth =16875
                    LayoutCachedHeight =165
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =398
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =984
                    Top =48
                    Width =1908
                    Height =312
                    FontSize =10
                    TabIndex =2
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =984
                    LayoutCachedTop =48
                    LayoutCachedWidth =2892
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =36
                    Top =48
                    Width =876
                    Height =312
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =26163
                    ForeColor =26163
                    Name ="AssetID"
                    ControlSource ="AssetID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =36
                    LayoutCachedTop =48
                    LayoutCachedWidth =912
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3612
                    Top =48
                    Width =1560
                    Height =312
                    FontSize =10
                    TabIndex =4
                    Name ="UnitMake"
                    ControlSource ="UnitMake"
                    StatusBarText ="unit make"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =3612
                    LayoutCachedTop =48
                    LayoutCachedWidth =5172
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5244
                    Top =48
                    Width =1908
                    Height =312
                    FontSize =10
                    TabIndex =5
                    Name ="UnitModel"
                    ControlSource ="UnitModel"
                    StatusBarText ="unit model"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =5244
                    LayoutCachedTop =48
                    LayoutCachedWidth =7152
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7224
                    Top =48
                    Width =1752
                    Height =312
                    ColumnWidth =3045
                    FontSize =10
                    TabIndex =6
                    Name ="UnitType"
                    ControlSource ="Stype"
                    StatusBarText ="UNIT TYPE"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =7224
                    LayoutCachedTop =48
                    LayoutCachedWidth =8976
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9048
                    Top =48
                    Width =1332
                    Height =312
                    FontSize =10
                    TabIndex =7
                    Name ="SoldPrice"
                    ControlSource ="SoldPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =9048
                    LayoutCachedTop =48
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15000
                    Top =48
                    Width =1065
                    Height =312
                    FontSize =10
                    TabIndex =12
                    Name ="RmktSaleDt"
                    ControlSource ="InvoicePaidDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =15000
                    LayoutCachedTop =48
                    LayoutCachedWidth =16065
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =2
                    IMESentenceMode =3
                    Left =16140
                    Top =48
                    Width =600
                    Height =312
                    ColumnWidth =1965
                    FontSize =10
                    TabIndex =13
                    BorderColor =-2147483615
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =16140
                    LayoutCachedTop =48
                    LayoutCachedWidth =16740
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13860
                    Top =48
                    Width =1065
                    Height =312
                    FontSize =10
                    TabIndex =11
                    BorderColor =-2147483615
                    Name ="InvoiceSentDate"
                    ControlSource ="InvoiceSentDate"
                    Format ="Short Date"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =13860
                    LayoutCachedTop =48
                    LayoutCachedWidth =14925
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12720
                    Top =48
                    Width =1065
                    Height =312
                    FontSize =10
                    TabIndex =10
                    BorderColor =-2147483615
                    Name ="RequestDate"
                    ControlSource ="InvoiceRequestDate"
                    Format ="Short Date"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =12720
                    LayoutCachedTop =48
                    LayoutCachedWidth =13785
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2964
                    Top =48
                    Width =576
                    Height =312
                    FontSize =10
                    TabIndex =3
                    BorderColor =-2147483615
                    Name ="Text29"
                    ControlSource ="unityr"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2964
                    LayoutCachedTop =48
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11580
                    Top =48
                    Width =1065
                    Height =312
                    FontSize =10
                    TabIndex =9
                    BackColor =14151142
                    Name ="RmktDepdt"
                    ControlSource ="RmktDepdt"
                    Format ="Short Date"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =11580
                    LayoutCachedTop =48
                    LayoutCachedWidth =12645
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10452
                    Top =48
                    Width =1056
                    Height =312
                    FontSize =10
                    TabIndex =8
                    BackColor =14151142
                    Name ="RmktDeposit"
                    ControlSource ="RmktDeposit"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =10452
                    LayoutCachedTop =48
                    LayoutCachedWidth =11508
                    LayoutCachedHeight =360
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =16920
                    Top =120
                    Width =375
                    BorderColor =12632256
                    Name ="Check31"
                    ControlSource ="pbos"

                    LayoutCachedLeft =16920
                    LayoutCachedTop =120
                    LayoutCachedWidth =17295
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =17220
                    Top =60
                    Width =720
                    Height =270
                    FontSize =9
                    TabIndex =14
                    Name ="UNSaleUnit"
                    Caption ="Un Sell"
                    OnClick ="[Event Procedure]"
                    BottomPadding =165

                    LayoutCachedLeft =17220
                    LayoutCachedTop =60
                    LayoutCachedWidth =17940
                    LayoutCachedHeight =330
                    RowStart =6
                    RowEnd =6
                    ColumnStart =6
                    ColumnEnd =6
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =5
                    BackTint =100.0
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    PressedColor =16777215
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
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
' See "ReMarketInvoiceSub.cls"
