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
    Width =21600
    DatasheetFontHeight =11
    ItemSuffix =307
    Left =5250
    Top =3450
    Right =20085
    Bottom =14085
    ShortcutMenuBar ="SortClip"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0x645ddc5e3577e440
    End
    RecordSource ="Bulk_Update_SelectedUnits"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
            Height =1350
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Top =1152
                    Width =21600
                    Height =192
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =1152
                    LayoutCachedWidth =21600
                    LayoutCachedHeight =1344
                    TabIndex =9
                    LayoutGroup =1
                    GroupTable =7
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =225
                    Width =2730
                    Height =510
                    FontSize =20
                    FontWeight =700
                    Name ="Label115"
                    Caption ="Interim Funding"
                    LayoutCachedLeft =120
                    LayoutCachedTop =225
                    LayoutCachedWidth =2850
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =4080
                    Top =375
                    Width =1935
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="IFBankSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4080
                    LayoutCachedTop =375
                    LayoutCachedWidth =6015
                    LayoutCachedHeight =615
                    ColumnStart =3
                    ColumnEnd =3
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
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =10965
                    Top =90
                    Width =765
                    Height =210
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedLeft =10965
                    LayoutCachedTop =90
                    LayoutCachedWidth =11730
                    LayoutCachedHeight =300
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =885
                    Top =1035
                    Width =1305
                    Height =210
                    FontSize =9
                    Name ="Label182"
                    Caption ="Unit #"
                    RightPadding =38
                    LayoutCachedLeft =885
                    LayoutCachedTop =1035
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =1245
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =2265
                    Top =1035
                    Width =1740
                    Height =210
                    FontSize =9
                    Name ="Label181"
                    Caption ="VIN"
                    RightPadding =38
                    LayoutCachedLeft =2265
                    LayoutCachedTop =1035
                    LayoutCachedWidth =4005
                    LayoutCachedHeight =1245
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =30
                    Top =1035
                    Width =780
                    Height =210
                    FontSize =9
                    Name ="Label180"
                    Caption ="♦AssetID"
                    RightPadding =38
                    LayoutCachedLeft =30
                    LayoutCachedTop =1035
                    LayoutCachedWidth =810
                    LayoutCachedHeight =1245
                    RowStart =3
                    RowEnd =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =8805
                    Top =1035
                    Width =1335
                    Height =225
                    FontSize =9
                    Name ="LabelIRP"
                    Caption ="Pay Off Date"
                    RightPadding =38
                    LayoutCachedLeft =8805
                    LayoutCachedTop =1035
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =1260
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =6060
                    Top =1035
                    Width =1260
                    Height =210
                    FontSize =9
                    Name ="Label175"
                    Caption ="Amount"
                    RightPadding =38
                    LayoutCachedLeft =6060
                    LayoutCachedTop =1035
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =1245
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =7380
                    Top =1035
                    Width =1365
                    Height =225
                    FontSize =9
                    Name ="Label176"
                    Caption ="Fund Date"
                    RightPadding =38
                    LayoutCachedLeft =7380
                    LayoutCachedTop =1035
                    LayoutCachedWidth =8745
                    LayoutCachedHeight =1260
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =7380
                    Top =375
                    Width =1365
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="IFDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7380
                    LayoutCachedTop =375
                    LayoutCachedWidth =8745
                    LayoutCachedHeight =615
                    ColumnStart =6
                    ColumnEnd =6
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
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7380
                    Top =720
                    Width =1365
                    Height =255
                    FontSize =9
                    TabIndex =8
                    Name ="IFDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =7380
                    LayoutCachedTop =720
                    LayoutCachedWidth =8745
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10320
                    Top =720
                    Width =1260
                    Height =255
                    FontSize =9
                    TabIndex =7
                    Name ="IFAmount"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =10320
                    LayoutCachedTop =720
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =6060
                    Top =375
                    Width =1260
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="IFAmountSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6060
                    LayoutCachedTop =375
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =615
                    ColumnStart =5
                    ColumnEnd =5
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
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =4080
                    Top =1035
                    Width =1935
                    Height =210
                    FontSize =9
                    Name ="Label220"
                    Caption ="Bank"
                    RightPadding =38
                    LayoutCachedLeft =4080
                    LayoutCachedTop =1035
                    LayoutCachedWidth =6015
                    LayoutCachedHeight =1245
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                End
                Begin ComboBox
                    Enabled = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =6840
                    Left =4080
                    Top =720
                    Width =1935
                    Height =255
                    FontSize =9
                    TabIndex =5
                    Name ="IFBankCBO"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Clients.ClientShNm, Clients.clientCompanyName FROM Clients WHERE (((Clien"
                        "ts.clientType)=\"bank\")) ORDER BY Clients.ClientShNm;"
                    ColumnWidths ="1800;5040"
                    StatusBarText ="bank used"
                    RightPadding =38
                    AllowValueListEdits =0

                    LayoutCachedLeft =4080
                    LayoutCachedTop =720
                    LayoutCachedWidth =6015
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =8805
                    Top =375
                    Width =1335
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="PayOffDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8805
                    LayoutCachedTop =375
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =615
                    ColumnStart =4
                    ColumnEnd =4
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
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8805
                    Top =720
                    Width =1335
                    Height =255
                    FontSize =9
                    TabIndex =6
                    Name ="PayOffDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =8805
                    LayoutCachedTop =720
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =975
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =6060
                    Top =720
                    Width =1260
                    Height =210
                    FontSize =9
                    Name ="Label306"
                    Caption ="Interim All"
                    RightPadding =38
                    LayoutCachedLeft =6060
                    LayoutCachedTop =720
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =930
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =2
                    BackShade =75.0
                    ForeThemeColorIndex =1
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
                    Left =8805
                    Top =45
                    Width =1335
                    Height =270
                    FontSize =9
                    TabIndex =5
                    Name ="UnitPayOff"
                    ControlSource ="UnitPayOff"
                    Format ="Short Date"
                    StatusBarText ="yes unit is IRP'ed"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =8805
                    LayoutCachedTop =45
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6060
                    Top =45
                    Width =1260
                    Height =270
                    FontSize =9
                    TabIndex =6
                    Name ="UnitIFAmt"
                    ControlSource ="UnitIFAmt"
                    StatusBarText ="IRP Acct Number"
                    AfterUpdate ="[Event Procedure]"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =6060
                    LayoutCachedTop =45
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7380
                    Top =45
                    Width =1365
                    Height =270
                    FontSize =9
                    TabIndex =7
                    Name ="UnitIFDate"
                    ControlSource ="UnitIFDate"
                    Format ="Short Date"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =7380
                    LayoutCachedTop =45
                    LayoutCachedWidth =8745
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =30
                    Top =45
                    Width =780
                    Height =270
                    FontSize =9
                    TabIndex =1
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =30
                    LayoutCachedTop =45
                    LayoutCachedWidth =810
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =2265
                    Top =45
                    Width =1740
                    Height =270
                    FontSize =9
                    TabIndex =3
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    StatusBarText ="VIN number must be unique in table"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2265
                    LayoutCachedTop =45
                    LayoutCachedWidth =4005
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =885
                    Top =45
                    Width =1305
                    Height =270
                    FontSize =9
                    TabIndex =2
                    Name ="UnitUnitNum"
                    ControlSource ="UnitUnitNum"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =885
                    LayoutCachedTop =45
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =314
                    Width =19584
                    Height =72
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
                    LayoutCachedWidth =19584
                    LayoutCachedHeight =386
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =6840
                    Left =4080
                    Top =45
                    Width =1935
                    Height =270
                    FontSize =9
                    TabIndex =4
                    Name ="UnitIFBank"
                    ControlSource ="UnitIFBank"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Clients.ClientShNm, Clients.clientCompanyName FROM Clients WHERE (((Clien"
                        "ts.clientType)=\"bank\")) ORDER BY Clients.ClientShNm;"
                    ColumnWidths ="1800;5040"
                    StatusBarText ="bank used"
                    AfterUpdate ="[Event Procedure]"
                    RightPadding =38
                    BottomPadding =38
                    AllowValueListEdits =0

                    LayoutCachedLeft =4080
                    LayoutCachedTop =45
                    LayoutCachedWidth =6015
                    LayoutCachedHeight =315
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11220
                    Width =1680
                    Height =270
                    FontSize =9
                    TabIndex =8
                    Name ="VendorCost"
                    ControlSource ="VendorCost"
                    StatusBarText ="IRP Acct Number"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =11220
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =270
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
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
' See "BulkChange_Interim.cls"
