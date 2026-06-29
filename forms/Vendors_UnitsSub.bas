Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =19651
    DatasheetFontHeight =11
    ItemSuffix =138
    Left =4485
    Top =12540
    Right =21690
    Bottom =17835
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x5651f8e3b2d9e540
    End
    RecordSource ="SELECT vw_SixKeys.* FROM vw_SixKeys INNER JOIN Units ON vw_SixKeys.UnitID = Unit"
        "s.UnitID WHERE (((vw_SixKeys.mlorig)=\"FA\") AND ((vw_SixKeys.VendorId)=[Forms]!"
        "[VendorsForm]![VendorID])) ORDER BY vw_SixKeys.ClientShortName, vw_SixKeys.Asset"
        "ID DESC;"
    Caption ="UnitsIn_sub"
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
            BorderColor =-2147483609
            FontName ="Calibri"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            BackColor =-2147483613
            BorderColor =-2147483609
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BackColor =-2147483613
            BorderColor =-2147483609
        End
        Begin CommandButton
            TextFontCharSet =238
            FontSize =11
            FontWeight =400
            ForeColor =-2147483616
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =-2147483609
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin BoundObjectFrame
            SizeMode =3
            BorderLineStyle =0
            BackStyle =0
            LabelX =-1800
            BorderColor =-2147483609
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =238
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =-2147483609
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin CustomControl
            OldBorderStyle =1
            BorderColor =-2147483609
        End
        Begin ToggleButton
            TextFontCharSet =238
            FontSize =11
            FontWeight =400
            ForeColor =-2147483616
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =238
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Attachment
            BackStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =4800
            Height =3840
            BorderColor =-2147483609
            LabelX =-1800
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =353
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    OldBorderStyle =0
                    PictureType =2
                    Top =143
                    Width =19080
                    Height =210
                    BorderColor =15527148
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"

                    LayoutCachedTop =143
                    LayoutCachedWidth =19080
                    LayoutCachedHeight =353
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =16752
                    Width =0
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="MLNo_Label"
                    Caption ="MLNo"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16752
                    LayoutCachedWidth =16752
                    LayoutCachedHeight =270
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =13188
                    Width =1056
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitFinalCost_Label"
                    Caption ="UnitCost"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13188
                    LayoutCachedWidth =14244
                    LayoutCachedHeight =270
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =16764
                    Width =0
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="SchNo_Label"
                    Caption ="SchNo"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16764
                    LayoutCachedWidth =16764
                    LayoutCachedHeight =270
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =1752
                    Width =708
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitID_Label"
                    Caption ="AssetID"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1752
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =270
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =4464
                    Width =1020
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitUnitNum_Label"
                    Caption ="Unit #"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4464
                    LayoutCachedWidth =5484
                    LayoutCachedHeight =270
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =12216
                    Width =972
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="unitinserv_Label"
                    Caption ="In Service"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12216
                    LayoutCachedWidth =13188
                    LayoutCachedHeight =270
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =14244
                    Width =1056
                    Height =270
                    FontSize =9
                    Name ="UnitPlateNum_Label"
                    Caption ="Vendor Cost"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14244
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =270
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =3
                    BackShade =50.0
                    ForeThemeColorIndex =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =9636
                    Width =636
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitStatus_Label"
                    Caption ="Status"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9636
                    LayoutCachedWidth =10272
                    LayoutCachedHeight =270
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =10272
                    Width =972
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitEstDelDate_Label"
                    Caption ="Est. Delv."
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10272
                    LayoutCachedWidth =11244
                    LayoutCachedHeight =270
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =11244
                    Width =972
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitActualDelDate_Label"
                    Caption ="Actual Delv."
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11244
                    LayoutCachedWidth =12216
                    LayoutCachedHeight =270
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =2496
                    Width =1968
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitVIN_Label"
                    Caption ="♦VIN"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2496
                    LayoutCachedWidth =4464
                    LayoutCachedHeight =270
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =5484
                    Width =4140
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitYr_Label"
                    Caption ="Unit Description"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5484
                    LayoutCachedWidth =9624
                    LayoutCachedHeight =270
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =16776
                    Width =0
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label67"
                    Caption ="Text66"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16776
                    LayoutCachedWidth =16776
                    LayoutCachedHeight =270
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =36
                    Width =1656
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label106"
                    Caption ="Client"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =36
                    LayoutCachedWidth =1692
                    LayoutCachedHeight =270
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin EmptyCell
                    Left =15300
                    Height =270
                    Name ="EmptyCell133"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15300
                    LayoutCachedWidth =16740
                    LayoutCachedHeight =270
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =3
                End
            End
        End
        Begin Section
            Height =225
            Name ="Detail"
            AlternateBackColor =16381933
            Begin
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1752
                    Width =708
                    Height =225
                    ColumnWidth =1440
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =1
                    BorderColor =-2147483615
                    Name ="UnitID"
                    ControlSource ="AssetID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1752
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =16752
                    Width =0
                    Height =225
                    ColumnWidth =1860
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =12
                    Name ="MLNo"
                    ControlSource ="MLNo"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16752
                    LayoutCachedWidth =16752
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =16764
                    Width =0
                    Height =225
                    ColumnWidth =3000
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =13
                    Name ="SGrpID"
                    ControlSource ="SGrpID"
                    StatusBarText ="schedule number starts with 1"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16764
                    LayoutCachedWidth =16764
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9636
                    Width =636
                    Height =225
                    ColumnWidth =840
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =5
                    BorderColor =-2147483615
                    Name ="UnitStatus"
                    ControlSource ="Status"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9636
                    LayoutCachedWidth =10272
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10272
                    Width =972
                    Height =225
                    ColumnWidth =2700
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =6
                    Name ="UnitEstDelDate"
                    ControlSource ="EsitmatedDeliveryDate"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10272
                    LayoutCachedWidth =11244
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2496
                    Width =1968
                    Height =225
                    ColumnWidth =2100
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =2
                    Name ="UnitVIN"
                    ControlSource ="Vin"
                    StatusBarText ="VIN number must be unique in table"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2496
                    LayoutCachedWidth =4464
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4464
                    Width =1020
                    Height =225
                    ColumnWidth =1500
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =3
                    Name ="UnitUnitNum"
                    ControlSource ="Unitnum"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4464
                    LayoutCachedWidth =5484
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5484
                    Width =4140
                    Height =225
                    ColumnWidth =990
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =4
                    Name ="UnitYr"
                    ControlSource ="UnitDescShortVer"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5484
                    LayoutCachedWidth =9624
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13188
                    Width =1056
                    Height =225
                    ColumnWidth =3000
                    ColumnOrder =13
                    FontSize =9
                    TabIndex =9
                    Name ="UnitCost"
                    ControlSource ="UnitCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13188
                    LayoutCachedWidth =14244
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11244
                    Width =972
                    Height =225
                    ColumnWidth =2700
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =7
                    Name ="UnitActualDelDate"
                    ControlSource ="ActualDeliveryDate"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11244
                    LayoutCachedWidth =12216
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12216
                    Width =972
                    Height =225
                    ColumnWidth =1275
                    FontSize =9
                    TabIndex =8
                    Name ="unitinserv"
                    ControlSource ="InServiceDate"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12216
                    LayoutCachedWidth =13188
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =16776
                    Width =0
                    Height =225
                    FontSize =9
                    TabIndex =14
                    Name ="SchID"
                    ControlSource ="SchID"
                    StatusBarText ="schedule number starts with 1"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16776
                    LayoutCachedWidth =16776
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =36
                    Width =1656
                    Height =225
                    FontSize =9
                    Name ="ClientShortName"
                    ControlSource ="ClientShortName"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =36
                    LayoutCachedWidth =1692
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15300
                    Height =225
                    FontSize =9
                    TabIndex =11
                    Name ="VendorId"
                    ControlSource ="VendorId"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15300
                    LayoutCachedWidth =16740
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14244
                    Width =1056
                    Height =225
                    FontSize =9
                    TabIndex =10
                    Name ="VendorCost"
                    ControlSource ="VendorCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14244
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =225
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =3
                End
            End
        End
        Begin FormFooter
            Height =15
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4155
                    Top =15
                    Height =0
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =4155
                    LayoutCachedTop =15
                    LayoutCachedWidth =5595
                    LayoutCachedHeight =15
                End
            End
        End
    End
End
CodeBehindForm
' See "Vendors_UnitsSub.cls"
