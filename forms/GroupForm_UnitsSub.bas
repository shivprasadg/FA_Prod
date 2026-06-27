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
    AllowUpdating =2
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =19440
    DatasheetFontHeight =11
    ItemSuffix =120
    Left =5145
    Top =11310
    Right =24210
    Bottom =18810
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x310ce2a7732ee640
    End
    RecordSource ="SELECT vw_SixKeys.*, Units.UnitCOASent, Nz([UnitCOABack],[UnitCofACopy]) AS COA,"
        " Units.UnitUnitNum FROM vw_SixKeys INNER JOIN Units ON vw_SixKeys.UnitID = Units"
        ".UnitID ORDER BY vw_SixKeys.AssetID;"
    Caption ="UnitsIn_sub"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
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
                    Width =19440
                    Height =210
                    BorderColor =15527148
                    Name ="Image71"
                    Picture ="LtGrey2_Shadow"

                    LayoutCachedTop =143
                    LayoutCachedWidth =19440
                    LayoutCachedHeight =353
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =19125
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
                    LayoutCachedLeft =19125
                    LayoutCachedWidth =19125
                    LayoutCachedHeight =270
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =12780
                    Width =1290
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitFinalCost_Label"
                    Caption ="OEC"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12780
                    LayoutCachedWidth =14070
                    LayoutCachedHeight =270
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =19125
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
                    LayoutCachedLeft =19125
                    LayoutCachedWidth =19125
                    LayoutCachedHeight =270
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =30
                    Width =705
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitID_Label"
                    Caption ="AssetID"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =30
                    LayoutCachedWidth =735
                    LayoutCachedHeight =270
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =15045
                    Width =2085
                    Height =270
                    FontSize =9
                    Name ="ULoc_Label"
                    Caption ="Location"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15045
                    LayoutCachedWidth =17130
                    LayoutCachedHeight =270
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =2730
                    Width =994
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
                    LayoutCachedLeft =2730
                    LayoutCachedWidth =3724
                    LayoutCachedHeight =270
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =11805
                    Width =975
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
                    LayoutCachedLeft =11805
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =270
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =14070
                    Width =975
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitPlateNum_Label"
                    Caption ="Plate"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14070
                    LayoutCachedWidth =15045
                    LayoutCachedHeight =270
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =7275
                    Width =630
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
                    LayoutCachedLeft =7275
                    LayoutCachedWidth =7905
                    LayoutCachedHeight =270
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =7905
                    Width =975
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
                    LayoutCachedLeft =7905
                    LayoutCachedWidth =8880
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
                    Left =8880
                    Width =975
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
                    LayoutCachedLeft =8880
                    LayoutCachedWidth =9855
                    LayoutCachedHeight =270
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =765
                    Width =1965
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
                    LayoutCachedLeft =765
                    LayoutCachedWidth =2730
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
                    Left =3720
                    Width =3555
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
                    LayoutCachedLeft =3720
                    LayoutCachedWidth =7275
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
                    Left =9855
                    Width =975
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="UnitCOASent_Label"
                    Caption ="CoA Sent"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9855
                    LayoutCachedWidth =10830
                    LayoutCachedHeight =270
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    Left =19125
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
                    LayoutCachedLeft =19125
                    LayoutCachedWidth =19125
                    LayoutCachedHeight =270
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10830
                    Width =975
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label86"
                    Caption ="CoA Back"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10830
                    LayoutCachedWidth =11805
                    LayoutCachedHeight =270
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =17130
                    Width =1155
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label98"
                    Caption ="Curr Rent"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17130
                    LayoutCachedWidth =18285
                    LayoutCachedHeight =270
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =18315
                    Width =765
                    Height =270
                    FontSize =9
                    ForeColor =-2147483615
                    Name ="Label106"
                    Caption ="Type"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =18315
                    LayoutCachedWidth =19080
                    LayoutCachedHeight =270
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
            End
        End
        Begin Section
            Height =255
            Name ="Detail"
            AlternateBackColor =16381933
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10830
                    Width =975
                    Height =255
                    FontSize =9
                    TabIndex =8
                    Name ="Text80"
                    ControlSource ="COA"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10830
                    LayoutCachedWidth =11805
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =30
                    Width =705
                    Height =255
                    ColumnWidth =1440
                    ColumnOrder =0
                    FontSize =9
                    BorderColor =-2147483615
                    Name ="UnitID"
                    ControlSource ="AssetID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =3
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =30
                    LayoutCachedWidth =735
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =19125
                    Width =0
                    Height =255
                    ColumnWidth =1860
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =15
                    Name ="MLNo"
                    ControlSource ="MLNo"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19125
                    LayoutCachedWidth =19125
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =19125
                    Width =0
                    Height =255
                    ColumnWidth =3000
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =16
                    Name ="SGrpID"
                    ControlSource ="SGrpID"
                    StatusBarText ="schedule number starts with 1"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19125
                    LayoutCachedWidth =19125
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7275
                    Width =630
                    Height =255
                    ColumnWidth =840
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =4
                    BorderColor =-2147483615
                    Name ="UnitStatus"
                    ControlSource ="Status"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7275
                    LayoutCachedWidth =7905
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7905
                    Width =975
                    Height =255
                    ColumnWidth =2700
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =5
                    Name ="UnitEstDelDate"
                    ControlSource ="EsitmatedDeliveryDate"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7905
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =765
                    Width =1965
                    Height =255
                    ColumnWidth =2100
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =1
                    Name ="UnitVIN"
                    ControlSource ="Vin"
                    StatusBarText ="VIN number must be unique in table"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =765
                    LayoutCachedWidth =2730
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3720
                    Width =3555
                    Height =255
                    ColumnWidth =990
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =3
                    Name ="UnitYr"
                    ControlSource ="UnitDescShortVer"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3720
                    LayoutCachedWidth =7275
                    LayoutCachedHeight =255
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
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9855
                    Width =975
                    Height =255
                    ColumnWidth =1290
                    ColumnOrder =16
                    FontSize =9
                    TabIndex =7
                    Name ="UnitCOASent"
                    ControlSource ="UnitCOASent"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9855
                    LayoutCachedWidth =10830
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12780
                    Width =1290
                    Height =255
                    ColumnWidth =3000
                    ColumnOrder =13
                    FontSize =9
                    TabIndex =10
                    Name ="UnitFinalCost"
                    ControlSource ="OECUnit"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12780
                    LayoutCachedWidth =14070
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8880
                    Width =975
                    Height =255
                    ColumnWidth =2700
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =6
                    Name ="UnitActualDelDate"
                    ControlSource ="ActualDeliveryDate"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8880
                    LayoutCachedWidth =9855
                    LayoutCachedHeight =255
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
                    Left =15045
                    Width =2085
                    Height =255
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =12
                    Name ="ULoc"
                    ControlSource ="Locationsn"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15045
                    LayoutCachedWidth =17130
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11805
                    Width =975
                    Height =255
                    ColumnWidth =1275
                    FontSize =9
                    TabIndex =9
                    Name ="unitinserv"
                    ControlSource ="InServiceDate"
                    Format ="mm/dd/yy"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11805
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14070
                    Width =975
                    Height =255
                    ColumnWidth =1740
                    FontSize =9
                    TabIndex =11
                    Name ="UnitPlateNum"
                    ControlSource ="UnitPlateNum"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14070
                    LayoutCachedWidth =15045
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =19125
                    Width =0
                    Height =255
                    FontSize =9
                    TabIndex =17
                    Name ="SchID"
                    ControlSource ="SchID"
                    StatusBarText ="schedule number starts with 1"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19125
                    LayoutCachedWidth =19125
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =1
                    GroupTable =3
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17130
                    Width =1155
                    Height =255
                    FontSize =9
                    TabIndex =13
                    Name ="UnitRent"
                    ControlSource ="CurrRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =17130
                    LayoutCachedWidth =18285
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =18315
                    Width =765
                    Height =255
                    FontSize =9
                    TabIndex =14
                    Name ="Text105"
                    ControlSource ="CurrRentType"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =3
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =18315
                    LayoutCachedWidth =19080
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2730
                    Width =994
                    Height =255
                    FontSize =9
                    TabIndex =2
                    BorderColor =-2147483615
                    Name ="Text118"
                    ControlSource ="Unitnum"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2730
                    LayoutCachedWidth =3724
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
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
' See "GroupForm_UnitsSub.cls"
