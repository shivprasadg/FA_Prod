Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =204
    BorderStyle =3
    PictureAlignment =4
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15741
    DatasheetFontHeight =11
    ItemSuffix =1286
    Left =4845
    Top =9675
    Right =22605
    Bottom =15975
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x50da427acd97e540
    End
    OnDirty ="[Event Procedure]"
    Caption ="Child Parts"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    NavigationCaption ="Child Parts"
    FilterOnLoad =255
    ShowPageMargins =0
    DatasheetAlternateBackColor =16777215
    DatasheetGridlinesColor12 =15062992
    FitToScreen =255
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
            BorderColor =12632256
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =5160
            Name ="Detail"
            AlternateBackColor =14270637
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =1080
                    Width =15720
                    Height =210
                    BorderColor =16777215
                    Name ="Image767"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =1080
                    LayoutCachedWidth =15720
                    LayoutCachedHeight =1290
                    TabIndex =26
                    BorderShade =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =255
                    Width =15720
                    Height =1110
                    Name ="Box910"
                    HorizontalAnchor =2
                    LayoutCachedWidth =15720
                    LayoutCachedHeight =1110
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9060
                    Top =540
                    Width =2160
                    Height =300
                    TabIndex =5
                    Name ="ChildFinalCost"
                    ControlSource ="childFinalCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    ValidationText ="Child asset cost mut be >=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =9060
                    LayoutCachedTop =540
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =840
                    DatasheetCaption ="Final Cost"
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =255
                            TextAlign =2
                            Left =8340
                            Top =540
                            Width =720
                            Height =300
                            FontSize =10
                            Name ="Label19"
                            Caption ="Cost:"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =8340
                            LayoutCachedTop =540
                            LayoutCachedWidth =9060
                            LayoutCachedHeight =840
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            BackThemeColorIndex =2
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6156
                    Top =2208
                    Width =2208
                    Height =252
                    FontSize =10
                    TabIndex =11
                    Name ="ChildInvoiceNum"
                    ControlSource ="ChildInvoiceNum"
                    GroupTable =2

                    LayoutCachedLeft =6156
                    LayoutCachedTop =2208
                    LayoutCachedWidth =8364
                    LayoutCachedHeight =2460
                    DatasheetCaption ="Invoice #"
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4416
                            Top =2208
                            Width =1668
                            Height =252
                            FontSize =8
                            Name ="Label21"
                            Caption ="Invoice #"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =4416
                            LayoutCachedTop =2208
                            LayoutCachedWidth =6084
                            LayoutCachedHeight =2460
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6156
                    Top =2532
                    Width =2208
                    Height =276
                    FontSize =10
                    TabIndex =13
                    Name ="ChildVendorPytDate"
                    ControlSource ="ChildVendorPytDate"
                    Format ="Short Date"
                    GroupTable =2

                    LayoutCachedLeft =6156
                    LayoutCachedTop =2532
                    LayoutCachedWidth =8364
                    LayoutCachedHeight =2808
                    DatasheetCaption ="Vendor Pyt Date"
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4416
                            Top =2532
                            Width =1668
                            Height =276
                            FontSize =8
                            Name ="Label23"
                            Caption ="Invoice Paid Date"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =4416
                            LayoutCachedTop =2532
                            LayoutCachedWidth =6084
                            LayoutCachedHeight =2808
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =13590
                    Top =2400
                    FontSize =10
                    TabIndex =16
                    Name ="ChildYear"
                    ControlSource ="ChildYear"

                    LayoutCachedLeft =13590
                    LayoutCachedTop =2400
                    LayoutCachedWidth =15030
                    LayoutCachedHeight =2640
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =13140
                            Top =2400
                            Width =390
                            Height =240
                            FontSize =8
                            Name ="Label33"
                            Caption ="Year"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =13140
                            LayoutCachedTop =2400
                            LayoutCachedWidth =13530
                            LayoutCachedHeight =2640
                            RowStart =6
                            RowEnd =6
                            ColumnStart =2
                            ColumnEnd =2
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6156
                    Top =3372
                    Width =2208
                    Height =252
                    FontSize =10
                    TabIndex =18
                    Name ="REFENG"
                    ControlSource ="REFENG"
                    StatusBarText ="REEFER ENG SERIAL NO"
                    GroupTable =18

                    LayoutCachedLeft =6156
                    LayoutCachedTop =3372
                    LayoutCachedWidth =8364
                    LayoutCachedHeight =3624
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =4416
                            Top =3372
                            Width =1668
                            Height =252
                            FontSize =8
                            Name ="Label60"
                            Caption ="Reefer Eng "
                            FontName ="Segoe UI"
                            GroupTable =18
                            LayoutCachedLeft =4416
                            LayoutCachedTop =3372
                            LayoutCachedWidth =6084
                            LayoutCachedHeight =3624
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6156
                    Top =3696
                    Width =2208
                    Height =252
                    FontSize =10
                    TabIndex =19
                    Name ="COMPRESSS"
                    ControlSource ="COMPRESSS"
                    StatusBarText ="COMPRESSOR SERIAL NO"
                    GroupTable =18

                    LayoutCachedLeft =6156
                    LayoutCachedTop =3696
                    LayoutCachedWidth =8364
                    LayoutCachedHeight =3948
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4416
                            Top =3696
                            Width =1668
                            Height =252
                            FontSize =8
                            Name ="Label61"
                            Caption ="Compressor"
                            FontName ="Segoe UI"
                            GroupTable =18
                            LayoutCachedLeft =4416
                            LayoutCachedTop =3696
                            LayoutCachedWidth =6084
                            LayoutCachedHeight =3948
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6156
                    Top =4344
                    Width =2208
                    Height =288
                    FontSize =10
                    TabIndex =21
                    Name ="MICROPROC"
                    ControlSource ="MICROPROC"
                    StatusBarText ="MICRO PROC SERIAL NO"
                    GroupTable =18

                    LayoutCachedLeft =6156
                    LayoutCachedTop =4344
                    LayoutCachedWidth =8364
                    LayoutCachedHeight =4632
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4416
                            Top =4344
                            Width =1668
                            Height =288
                            FontSize =8
                            Name ="Label62"
                            Caption ="Micro Processor"
                            FontName ="Segoe UI"
                            GroupTable =18
                            LayoutCachedLeft =4416
                            LayoutCachedTop =4344
                            LayoutCachedWidth =6084
                            LayoutCachedHeight =4632
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6156
                    Top =4020
                    Width =2208
                    Height =252
                    FontSize =10
                    TabIndex =20
                    Name ="ELECDISP"
                    ControlSource ="ELECDISP"
                    StatusBarText ="ELECTRONIC PROC SERIAL NO"
                    GroupTable =18

                    LayoutCachedLeft =6156
                    LayoutCachedTop =4020
                    LayoutCachedWidth =8364
                    LayoutCachedHeight =4272
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4416
                            Top =4020
                            Width =1668
                            Height =252
                            FontSize =8
                            Name ="Label63"
                            Caption ="Electronic Display"
                            FontName ="Segoe UI"
                            GroupTable =18
                            LayoutCachedLeft =4416
                            LayoutCachedTop =4020
                            LayoutCachedWidth =6084
                            LayoutCachedHeight =4272
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =18
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =14970
                    Top =1980
                    Width =270
                    Height =255
                    TabIndex =22
                    Name ="ChildEXCofa"
                    ControlSource ="ChildEXCofa"
                    StatusBarText ="exclude Child from CofA"

                    LayoutCachedLeft =14970
                    LayoutCachedTop =1980
                    LayoutCachedWidth =15240
                    LayoutCachedHeight =2235
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =13320
                            Top =1980
                            Width =1605
                            Height =255
                            FontSize =8
                            FontWeight =700
                            Name ="Label74"
                            Caption ="Exclude From COA "
                            FontName ="Segoe UI"
                            LayoutCachedLeft =13320
                            LayoutCachedTop =1980
                            LayoutCachedWidth =14925
                            LayoutCachedHeight =2235
                            RowStart =1
                            RowEnd =1
                            ColumnStart =4
                            ColumnEnd =5
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =14985
                    Top =1620
                    Width =255
                    Height =255
                    TabIndex =23
                    Name ="ExclFrGrpOEC"
                    ControlSource ="ExclFrGrpOEC"
                    StatusBarText ="exclude fro group OEC - not syndicated"
                    DefaultValue ="0"
                    ControlTipText ="Aounit is excluded from Group Cost/Fee%"

                    LayoutCachedLeft =14985
                    LayoutCachedTop =1620
                    LayoutCachedWidth =15240
                    LayoutCachedHeight =1875
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =13125
                            Top =1620
                            Width =1800
                            Height =255
                            FontSize =8
                            FontWeight =700
                            Name ="Label78"
                            Caption ="DO NOT Syndicate "
                            FontName ="Segoe UI"
                            ControlTipText ="Aounit is excluded from Group Cost/Fee%"
                            LayoutCachedLeft =13125
                            LayoutCachedTop =1620
                            LayoutCachedWidth =14925
                            LayoutCachedHeight =1875
                            RowStart =1
                            RowEnd =1
                            ColumnStart =7
                            ColumnEnd =7
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3165
                    Top =540
                    Width =5115
                    Height =300
                    TabIndex =4
                    Name ="VendorREF"
                    ControlSource ="ChildVendorREF"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName FROM Vendors WHERE (((Vendors.Vendor"
                        "Obsolete)=No)) ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;2160"
                    StatusBarText ="lookup from Vendor table"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Vendor is required"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =3165
                    LayoutCachedTop =540
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =840
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =2
                            Left =2400
                            Top =540
                            Width =780
                            Height =300
                            FontSize =10
                            Name ="Label14"
                            Caption ="Vendor:"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =2400
                            LayoutCachedTop =540
                            LayoutCachedWidth =3180
                            LayoutCachedHeight =840
                            RowStart =1
                            RowEnd =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2160
                    Left =3180
                    Top =120
                    Width =2160
                    Height =300
                    BoundColumn =1
                    Name ="ChildMake"
                    ControlSource ="ChildMake"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_ChildPartMake.ChildPartMakeID, LookUp_ChildPartMake.ChildPartMake "
                        "FROM LookUp_ChildPartMake ORDER BY LookUp_ChildPartMake.ChildPartMake;"
                    ColumnWidths ="0;2160"
                    StatusBarText ="unit make"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Unit make is required"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3180
                    LayoutCachedTop =120
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =420
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =2
                            Left =2400
                            Top =120
                            Width =780
                            Height =300
                            FontSize =10
                            Name ="Label2"
                            Caption ="Make:"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =2400
                            LayoutCachedTop =120
                            LayoutCachedWidth =3180
                            LayoutCachedHeight =420
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =9060
                    Top =120
                    Width =2160
                    Height =300
                    TabIndex =3
                    Name ="ChildType"
                    ControlSource ="ChildType"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="2880;0"
                    StatusBarText ="unit model"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =9060
                    LayoutCachedTop =120
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =420
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =8340
                            Top =120
                            Width =705
                            Height =315
                            FontSize =10
                            Name ="Label86"
                            Caption ="Type:"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =8340
                            LayoutCachedTop =120
                            LayoutCachedWidth =9045
                            LayoutCachedHeight =435
                            ColumnStart =1
                            ColumnEnd =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =2880
                    Left =6000
                    Top =120
                    Width =2280
                    Height =300
                    TabIndex =1
                    Name ="ChildClass"
                    ControlSource ="ChildClass"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT     Lookup_PartClasses.cPartClass,     Lookup_PartClasses.cPartClassID FR"
                        "OM     Lookup_PartClasses ORDER BY     Lookup_PartClasses.cPartClass;"
                    ColumnWidths ="2880;0"
                    StatusBarText ="UNIT TYPE"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Unit Type is required"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =6000
                    LayoutCachedTop =120
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =420
                    ColumnStart =5
                    ColumnEnd =8
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =2
                            Left =5340
                            Top =120
                            Width =660
                            Height =315
                            FontSize =10
                            Name ="Label4"
                            Caption ="Class:"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =5340
                            LayoutCachedTop =120
                            LayoutCachedWidth =6000
                            LayoutCachedHeight =435
                            ColumnStart =4
                            ColumnEnd =4
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13575
                    Top =4680
                    Width =2025
                    Height =285
                    ColumnWidth =675
                    FontSize =8
                    TabIndex =25
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =5026082
                    Name ="UnitREF"
                    ControlSource ="UnitREF"
                    StatusBarText ="link to Unit"
                    DefaultValue ="[Forms]![UnitsForm]![unitid]"

                    LayoutCachedLeft =13575
                    LayoutCachedTop =4680
                    LayoutCachedWidth =15600
                    LayoutCachedHeight =4965
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =3
                            Left =11520
                            Top =4680
                            Width =1995
                            Height =285
                            FontSize =8
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =5026082
                            Name ="Label3"
                            Caption ="UnitId"
                            LayoutCachedLeft =11520
                            LayoutCachedTop =4680
                            LayoutCachedWidth =13515
                            LayoutCachedHeight =4965
                            RowStart =3
                            RowEnd =3
                            ColumnStart =3
                            ColumnEnd =3
                            ThemeFontIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6156
                    Top =1884
                    Width =2208
                    Height =252
                    FontSize =10
                    TabIndex =9
                    Name ="ChildInvdt"
                    ControlSource ="ChildInvdt"
                    GroupTable =2

                    LayoutCachedLeft =6156
                    LayoutCachedTop =1884
                    LayoutCachedWidth =8364
                    LayoutCachedHeight =2136
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4416
                            Top =1884
                            Width =1668
                            Height =252
                            FontSize =8
                            Name ="Label93"
                            Caption ="Received"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =4416
                            LayoutCachedTop =1884
                            LayoutCachedWidth =6084
                            LayoutCachedHeight =2136
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =120
                    Top =1725
                    Width =3470
                    Height =2580
                    FontSize =8
                    TabIndex =17
                    BorderColor =12632256
                    Name ="ChildList"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;1440;1152;864;0"
                    OnClick ="[Event Procedure]"
                    GridlineColor =0

                    LayoutCachedLeft =120
                    LayoutCachedTop =1725
                    LayoutCachedWidth =3590
                    LayoutCachedHeight =4305
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =247
                    Left =210
                    Top =60
                    Width =1800
                    Height =480
                    FontSize =18
                    Name ="Label2153"
                    Caption ="Child Parts"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =210
                    LayoutCachedTop =60
                    LayoutCachedWidth =2010
                    LayoutCachedHeight =540
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2712
                    Top =4335
                    Width =879
                    FontSize =8
                    FontWeight =700
                    TabIndex =24
                    ForeColor =10040879
                    Name ="ChildPartSum"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =2712
                    LayoutCachedTop =4335
                    LayoutCachedWidth =3591
                    LayoutCachedHeight =4575
                    DatasheetCaption ="Final Cost"
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =1560
                            Top =4335
                            Width =1155
                            Height =240
                            FontSize =8
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label573"
                            Caption ="Total Child Cost"
                            LayoutCachedLeft =1560
                            LayoutCachedTop =4335
                            LayoutCachedWidth =2715
                            LayoutCachedHeight =4575
                            ColumnEnd =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =2
                            BorderThemeColorIndex =2
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =4416
                    Top =1560
                    Width =3948
                    Height =255
                    FontSize =8
                    Name ="Label677"
                    Caption =" Invoice"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LayoutCachedLeft =4416
                    LayoutCachedTop =1560
                    LayoutCachedWidth =8364
                    LayoutCachedHeight =1815
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =8424
                    Top =2208
                    Width =96
                    Height =252
                    Name ="EmptyCell724"
                    GroupTable =2
                    LayoutCachedLeft =8424
                    LayoutCachedTop =2208
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =2460
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =8424
                    Top =2532
                    Width =96
                    Height =276
                    Name ="EmptyCell725"
                    GroupTable =2
                    LayoutCachedLeft =8424
                    LayoutCachedTop =2532
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =2808
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =4416
                    Top =3060
                    Width =3948
                    Height =240
                    FontSize =8
                    Name ="Label757"
                    Caption ="Child Part Serials"
                    FontName ="Segoe UI"
                    GroupTable =18
                    LayoutCachedLeft =4416
                    LayoutCachedTop =3060
                    LayoutCachedWidth =8364
                    LayoutCachedHeight =3300
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =4
                    BackShade =75.0
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =8424
                    Top =3372
                    Width =72
                    Height =252
                    Name ="EmptyCell772"
                    GroupTable =18
                    LayoutCachedLeft =8424
                    LayoutCachedTop =3372
                    LayoutCachedWidth =8496
                    LayoutCachedHeight =3624
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =8424
                    Top =3696
                    Width =72
                    Height =252
                    Name ="EmptyCell773"
                    GroupTable =18
                    LayoutCachedLeft =8424
                    LayoutCachedTop =3696
                    LayoutCachedWidth =8496
                    LayoutCachedHeight =3948
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =8424
                    Top =4020
                    Width =72
                    Height =252
                    Name ="EmptyCell774"
                    GroupTable =18
                    LayoutCachedLeft =8424
                    LayoutCachedTop =4020
                    LayoutCachedWidth =8496
                    LayoutCachedHeight =4272
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =18
                End
                Begin EmptyCell
                    Left =8424
                    Top =4344
                    Width =72
                    Height =288
                    Name ="EmptyCell775"
                    GroupTable =18
                    LayoutCachedLeft =8424
                    LayoutCachedTop =4344
                    LayoutCachedWidth =8496
                    LayoutCachedHeight =4632
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =18
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =810
                    Top =600
                    Width =1200
                    Height =300
                    FontSize =10
                    TabIndex =7
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="ChildPartID"
                    ControlSource ="ChildPartID"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =810
                    LayoutCachedTop =600
                    LayoutCachedWidth =2010
                    LayoutCachedHeight =900
                    RowEnd =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =210
                            Top =600
                            Width =570
                            Height =300
                            FontSize =10
                            FontWeight =700
                            Name ="Label751"
                            Caption ="PartID"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =210
                            LayoutCachedTop =600
                            LayoutCachedWidth =780
                            LayoutCachedHeight =900
                            RowEnd =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11985
                    Top =540
                    Width =3375
                    Height =300
                    TabIndex =6
                    Name ="ChildVIN"
                    ControlSource ="ChildVIN"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11985
                    LayoutCachedTop =540
                    LayoutCachedWidth =15360
                    LayoutCachedHeight =840
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =2
                            Left =11340
                            Top =540
                            Width =645
                            Height =300
                            FontSize =10
                            Name ="Label7"
                            Caption ="Serial: "
                            FontName ="Segoe UI"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =11340
                            LayoutCachedTop =540
                            LayoutCachedWidth =11985
                            LayoutCachedHeight =840
                            RowEnd =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =120
                    Top =1365
                    Width =3470
                    FontSize =8
                    TabIndex =8
                    Name ="ListBoxHeader"
                    ControlSource ="=\"Child Parts Attached to Asset: \" & [Forms]![UnitsForm]![UnitID]"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =120
                    LayoutCachedTop =1365
                    LayoutCachedWidth =3590
                    LayoutCachedHeight =1605
                    ColumnEnd =1
                    BackThemeColorIndex =4
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin EmptyCell
                    Left =8424
                    Top =1884
                    Width =96
                    Height =252
                    Name ="EmptyCell1033"
                    GroupTable =2
                    LayoutCachedLeft =8424
                    LayoutCachedTop =1884
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =2136
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =8580
                    Top =1560
                    Width =3648
                    Height =255
                    FontSize =8
                    Name ="Label1067"
                    Caption ="Intrim Funding"
                    FontName ="Segoe UI"
                    GroupTable =2
                    LayoutCachedLeft =8580
                    LayoutCachedTop =1560
                    LayoutCachedWidth =12228
                    LayoutCachedHeight =1815
                    ColumnStart =3
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =9
                    BackShade =75.0
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =8424
                    Top =1560
                    Width =96
                    Height =255
                    Name ="EmptyCell1137"
                    GroupTable =2
                    LayoutCachedLeft =8424
                    LayoutCachedTop =1560
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =1815
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10092
                    Top =1884
                    Width =2136
                    Height =252
                    FontSize =10
                    TabIndex =10
                    Name ="Text1152"
                    ControlSource ="cIFDate"
                    Format ="Short Date"
                    GroupTable =2

                    LayoutCachedLeft =10092
                    LayoutCachedTop =1884
                    LayoutCachedWidth =12228
                    LayoutCachedHeight =2136
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8580
                            Top =1884
                            Width =1440
                            Height =252
                            FontSize =8
                            Name ="Label1153"
                            Caption ="IF Date"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =8580
                            LayoutCachedTop =1884
                            LayoutCachedWidth =10020
                            LayoutCachedHeight =2136
                            RowStart =1
                            RowEnd =1
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10092
                    Top =2532
                    Width =2136
                    Height =276
                    FontSize =10
                    TabIndex =14
                    Name ="Text1162"
                    ControlSource ="cIFAmount"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =2

                    LayoutCachedLeft =10092
                    LayoutCachedTop =2532
                    LayoutCachedWidth =12228
                    LayoutCachedHeight =2808
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8580
                            Top =2532
                            Width =1440
                            Height =276
                            FontSize =8
                            Name ="Label1163"
                            Caption ="IF Amount"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =8580
                            LayoutCachedTop =2532
                            LayoutCachedWidth =10020
                            LayoutCachedHeight =2808
                            RowStart =3
                            RowEnd =3
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin EmptyCell
                    Left =8424
                    Top =3060
                    Width =72
                    Name ="EmptyCell1199"
                    GroupTable =18
                    LayoutCachedLeft =8424
                    LayoutCachedTop =3060
                    LayoutCachedWidth =8496
                    LayoutCachedHeight =3300
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GroupTable =18
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =5760
                    Left =10092
                    Top =2880
                    Width =2136
                    FontSize =10
                    TabIndex =15
                    Name ="DrawDownID"
                    ControlSource ="DrawDownID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DrawDowns.DrawDownID, DrawDowns.DrawNumber AS Draw, DrawDowns.DrawName, B"
                        "anks.BankShortName AS IFBank FROM DrawDowns INNER JOIN Banks ON DrawDowns.DrawBa"
                        "nkId = Banks.BankID WHERE (((DrawDowns.GroupID)=[Forms]![UnitsForm]![GroupRef]))"
                        ";"
                    ColumnWidths ="0;720;2880;720"
                    GroupTable =2

                    LayoutCachedLeft =10092
                    LayoutCachedTop =2880
                    LayoutCachedWidth =12228
                    LayoutCachedHeight =3120
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8580
                            Top =2880
                            Width =1440
                            Height =240
                            FontSize =8
                            Name ="Label1173"
                            Caption ="Draw Down"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =8580
                            LayoutCachedTop =2880
                            LayoutCachedWidth =10020
                            LayoutCachedHeight =3120
                            RowStart =4
                            RowEnd =4
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6840
                    Left =10092
                    Top =2208
                    Width =2136
                    Height =252
                    FontSize =10
                    TabIndex =12
                    Name ="cIfBankID"
                    ControlSource ="cIFBankID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankID, Banks.BankShortName, Banks.BankName FROM Banks WHERE (((Ban"
                        "ks.BankType)=\"1\" Or (Banks.BankType)=\"3\")) ORDER BY Banks.BankShortName;"
                    ColumnWidths ="0;1800;5040"
                    StatusBarText ="bank used"
                    GroupTable =2
                    AllowValueListEdits =0

                    LayoutCachedLeft =10092
                    LayoutCachedTop =2208
                    LayoutCachedWidth =12228
                    LayoutCachedHeight =2460
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8580
                            Top =2208
                            Width =1440
                            Height =252
                            FontSize =8
                            Name ="Label357"
                            Caption ="Loan Bank"
                            FontName ="Segoe UI"
                            GroupTable =2
                            LayoutCachedLeft =8580
                            LayoutCachedTop =2208
                            LayoutCachedWidth =10020
                            LayoutCachedHeight =2460
                            RowStart =2
                            RowEnd =2
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =2
                        End
                    End
                End
                Begin EmptyCell
                    Left =4416
                    Top =2880
                    Width =1668
                    Name ="EmptyCell1233"
                    GroupTable =2
                    LayoutCachedLeft =4416
                    LayoutCachedTop =2880
                    LayoutCachedWidth =6084
                    LayoutCachedHeight =3120
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6156
                    Top =2880
                    Width =2208
                    Name ="EmptyCell1234"
                    GroupTable =2
                    LayoutCachedLeft =6156
                    LayoutCachedTop =2880
                    LayoutCachedWidth =8364
                    LayoutCachedHeight =3120
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =8424
                    Top =2880
                    Width =96
                    Name ="EmptyCell1235"
                    GroupTable =2
                    LayoutCachedLeft =8424
                    LayoutCachedTop =2880
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =3120
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =4416
                    Top =3180
                    Width =1668
                    Name ="EmptyCell1238"
                    GroupTable =2
                    LayoutCachedLeft =4416
                    LayoutCachedTop =3180
                    LayoutCachedWidth =6084
                    LayoutCachedHeight =3420
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6156
                    Top =3180
                    Width =2208
                    Name ="EmptyCell1239"
                    GroupTable =2
                    LayoutCachedLeft =6156
                    LayoutCachedTop =3180
                    LayoutCachedWidth =8364
                    LayoutCachedHeight =3420
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =8424
                    Top =3180
                    Width =96
                    Name ="EmptyCell1240"
                    GroupTable =2
                    LayoutCachedLeft =8424
                    LayoutCachedTop =3180
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =3420
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =10092
                    Top =3180
                    Width =2136
                    Name ="EmptyCell1251"
                    GroupTable =2
                    LayoutCachedLeft =10092
                    LayoutCachedTop =3180
                    LayoutCachedWidth =12228
                    LayoutCachedHeight =3420
                    RowStart =5
                    RowEnd =5
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =8580
                    Top =3180
                    Name ="EmptyCell1252"
                    GroupTable =2
                    LayoutCachedLeft =8580
                    LayoutCachedTop =3180
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =3420
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11985
                    Top =120
                    Width =3375
                    Height =300
                    TabIndex =2
                    Name ="cDesc"
                    ControlSource ="cDesc"

                    LayoutCachedLeft =11985
                    LayoutCachedTop =120
                    LayoutCachedWidth =15360
                    LayoutCachedHeight =420
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =11347
                            Top =120
                            Width =585
                            Height =315
                            FontSize =10
                            Name ="Label1265"
                            Caption ="Desc:"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =11347
                            LayoutCachedTop =120
                            LayoutCachedWidth =11932
                            LayoutCachedHeight =435
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ChildParts_sub.cls"
