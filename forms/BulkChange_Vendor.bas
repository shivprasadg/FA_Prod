Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =22639
    DatasheetFontHeight =11
    ItemSuffix =765
    Left =5175
    Top =4650
    Right =20085
    Bottom =10440
    ShortcutMenuBar ="SortClip"
    OrderBy ="[Bulk_Update_SelectedUnits].[LState] DESC"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0xac79b62e1528e640
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Bulk_Update_SelectedUnits"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
            GridlineColor =12632256
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1155
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Top =960
                    Width =22579
                    Height =195
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =960
                    LayoutCachedWidth =22579
                    LayoutCachedHeight =1155
                    TabIndex =16
                    LayoutGroup =3
                    GroupTable =4
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =17160
                    Top =810
                    Width =1500
                    Height =270
                    FontSize =9
                    Name ="Label755"
                    Caption ="Unit IF Amount"
                    GroupTable =6
                    LayoutCachedLeft =17160
                    LayoutCachedTop =810
                    LayoutCachedWidth =18660
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =40.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =3
                    Left =60
                    Top =810
                    Width =864
                    Height =270
                    FontSize =9
                    Name ="Label73"
                    Caption ="♦AssetID"
                    GroupTable =6
                    LayoutCachedLeft =60
                    LayoutCachedTop =810
                    LayoutCachedWidth =924
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =2235
                    Top =810
                    Width =2115
                    Height =270
                    FontSize =9
                    Name ="Label74"
                    Caption ="VIN"
                    GroupTable =6
                    LayoutCachedLeft =2235
                    LayoutCachedTop =810
                    LayoutCachedWidth =4350
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =990
                    Top =810
                    Width =1185
                    Height =270
                    FontSize =9
                    Name ="Label75"
                    Caption ="Unit #"
                    GroupTable =6
                    LayoutCachedLeft =990
                    LayoutCachedTop =810
                    LayoutCachedWidth =2175
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =5040
                    Top =810
                    Width =3165
                    Height =270
                    FontSize =9
                    Name ="Label83"
                    Caption ="Vendor"
                    GroupTable =6
                    LayoutCachedLeft =5040
                    LayoutCachedTop =810
                    LayoutCachedWidth =8205
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5040
                    Top =495
                    Width =3165
                    Height =255
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =7
                    Name ="VendorCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName, Vendors.VendorSNm FROM Vendors WHER"
                        "E (((Vendors.VendorObsolete)=No)) ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;1800;720"
                    StatusBarText ="lookup from Vendor table"
                    Tag ="Medit"
                    GroupTable =6
                    AllowValueListEdits =0

                    LayoutCachedLeft =5040
                    LayoutCachedTop =495
                    LayoutCachedWidth =8205
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5040
                    Top =180
                    Width =3165
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="VendorSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =6
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5040
                    LayoutCachedTop =180
                    LayoutCachedWidth =8205
                    LayoutCachedHeight =435
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =6
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =18660
                    Top =60
                    Width =765
                    Height =210
                    ColumnOrder =6
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedLeft =18660
                    LayoutCachedTop =60
                    LayoutCachedWidth =19425
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17280
                    Top =300
                    Width =780
                    Height =450
                    ColumnOrder =7
                    FontSize =20
                    FontWeight =700
                    TabIndex =13
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =17280
                    LayoutCachedTop =300
                    LayoutCachedWidth =18060
                    LayoutCachedHeight =750
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =4410
                    Top =810
                    Width =570
                    Height =270
                    FontSize =9
                    Name ="Label189"
                    Caption ="Status"
                    GroupTable =6
                    LayoutCachedLeft =4410
                    LayoutCachedTop =810
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =60
                    Width =4515
                    Height =540
                    FontSize =20
                    FontWeight =700
                    Name ="Label115"
                    Caption ="Vendor | Invoices | Draws"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =4575
                    LayoutCachedHeight =540
                    ForeThemeColorIndex =0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9915
                    Top =180
                    Width =1296
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="InvoiceSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =6
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9915
                    LayoutCachedTop =180
                    LayoutCachedWidth =11211
                    LayoutCachedHeight =435
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =6
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =8265
                    Top =810
                    Width =1584
                    Height =270
                    FontSize =9
                    Name ="Label124"
                    Caption ="Invoice Number"
                    GroupTable =6
                    LayoutCachedLeft =8265
                    LayoutCachedTop =810
                    LayoutCachedWidth =9849
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =9915
                    Top =810
                    Width =1296
                    Height =270
                    FontSize =9
                    Name ="Label125"
                    Caption ="Invoice Date"
                    GroupTable =6
                    LayoutCachedLeft =9915
                    LayoutCachedTop =810
                    LayoutCachedWidth =11211
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =11265
                    Top =810
                    Width =1290
                    Height =270
                    FontSize =9
                    Name ="Label126"
                    Caption ="Payment Date"
                    GroupTable =6
                    LayoutCachedLeft =11265
                    LayoutCachedTop =810
                    LayoutCachedWidth =12555
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9915
                    Top =495
                    Width =1296
                    Height =255
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =9
                    Name ="InvoiceDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    GroupTable =6

                    LayoutCachedLeft =9915
                    LayoutCachedTop =495
                    LayoutCachedWidth =11211
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11265
                    Top =180
                    Width =1290
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="VendorPayDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =6
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11265
                    LayoutCachedTop =180
                    LayoutCachedWidth =12555
                    LayoutCachedHeight =435
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =6
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11265
                    Top =495
                    Width =1290
                    Height =255
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =10
                    Name ="VendorPayDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    GroupTable =6

                    LayoutCachedLeft =11265
                    LayoutCachedTop =495
                    LayoutCachedWidth =12555
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8265
                    Top =180
                    Width =1584
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="InvoiceNumSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =6
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8265
                    LayoutCachedTop =180
                    LayoutCachedWidth =9849
                    LayoutCachedHeight =435
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =6
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8265
                    Top =495
                    Width =1584
                    Height =255
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =8
                    Name ="InvoiceNum"
                    Tag ="Medit"
                    GroupTable =6

                    LayoutCachedLeft =8265
                    LayoutCachedTop =495
                    LayoutCachedWidth =9849
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =14175
                    Top =810
                    Width =930
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label203"
                    Caption ="DRAW"
                    GroupTable =6
                    LayoutCachedLeft =14175
                    LayoutCachedTop =810
                    LayoutCachedWidth =15105
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14175
                    Top =180
                    Width =930
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="DrawDownSubmit"
                    Caption ="▼ Submit"
                    OnClick ="[Event Procedure]"
                    GroupTable =6
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14175
                    LayoutCachedTop =180
                    LayoutCachedWidth =15105
                    LayoutCachedHeight =435
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =6
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =4320
                    Left =14175
                    Top =495
                    Width =930
                    Height =255
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =12
                    BackColor =15590879
                    Name ="DrawDownCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DrawDowns.DrawDownId, DrawDowns.DrawNumber AS Draw, DrawDowns.DrawName, B"
                        "anks.BankShortName AS Bank, Banks.BankID FROM DrawDowns INNER JOIN Banks ON Draw"
                        "Downs.DrawBankId = Banks.bankID WHERE (((DrawDowns.GroupID)=[Forms]![BulkChange_"
                        "Form]![GrpID]));"
                    ColumnWidths ="0;432;1728;864;0"
                    Tag ="Medit"
                    GroupTable =6

                    LayoutCachedLeft =14175
                    LayoutCachedTop =495
                    LayoutCachedWidth =15105
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =60
                    Top =495
                    Width =864
                    Height =255
                    Name ="EmptyCell562"
                    GroupTable =6
                    LayoutCachedLeft =60
                    LayoutCachedTop =495
                    LayoutCachedWidth =924
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =990
                    Top =495
                    Width =1185
                    Height =255
                    Name ="EmptyCell563"
                    GroupTable =6
                    LayoutCachedLeft =990
                    LayoutCachedTop =495
                    LayoutCachedWidth =2175
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =2235
                    Top =495
                    Width =2115
                    Height =255
                    Name ="EmptyCell564"
                    GroupTable =6
                    LayoutCachedLeft =2235
                    LayoutCachedTop =495
                    LayoutCachedWidth =4350
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =60
                    Top =180
                    Width =864
                    Height =255
                    Name ="EmptyCell568"
                    GroupTable =6
                    LayoutCachedLeft =60
                    LayoutCachedTop =180
                    LayoutCachedWidth =924
                    LayoutCachedHeight =435
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =990
                    Top =180
                    Width =1185
                    Height =255
                    Name ="EmptyCell569"
                    GroupTable =6
                    LayoutCachedLeft =990
                    LayoutCachedTop =180
                    LayoutCachedWidth =2175
                    LayoutCachedHeight =435
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =2235
                    Top =180
                    Width =2115
                    Height =255
                    Name ="EmptyCell570"
                    GroupTable =6
                    LayoutCachedLeft =2235
                    LayoutCachedTop =180
                    LayoutCachedWidth =4350
                    LayoutCachedHeight =435
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =4410
                    Top =180
                    Width =570
                    Height =255
                    Name ="EmptyCell657"
                    GroupTable =6
                    LayoutCachedLeft =4410
                    LayoutCachedTop =180
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =435
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =4410
                    Top =495
                    Width =570
                    Height =255
                    Name ="EmptyCell658"
                    GroupTable =6
                    LayoutCachedLeft =4410
                    LayoutCachedTop =495
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =12615
                    Top =810
                    Width =1500
                    Height =270
                    FontSize =9
                    Name ="Label703"
                    Caption ="Vendor Cost"
                    GroupTable =6
                    LayoutCachedLeft =12615
                    LayoutCachedTop =810
                    LayoutCachedWidth =14115
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12615
                    Top =495
                    Width =1500
                    Height =255
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =11
                    Name ="VendorCostCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    ControlTipText ="Sets the Actual Vendor Cost"
                    GroupTable =6

                    LayoutCachedLeft =12615
                    LayoutCachedTop =495
                    LayoutCachedWidth =14115
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =40.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    CurrencySymbol ="$"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12615
                    Top =180
                    Width =1500
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="VendorCostSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Sets the Actual Vendor Cost for all of the Units Selected"
                    GroupTable =6
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12615
                    LayoutCachedTop =180
                    LayoutCachedWidth =14115
                    LayoutCachedHeight =435
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverShade =50.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =6
                    QuickStyle =31
                    QuickStyleMask =-117
                End
                Begin EmptyCell
                    Left =15165
                    Top =180
                    Width =1500
                    Height =255
                    Name ="EmptyCell722"
                    GroupTable =6
                    LayoutCachedLeft =15165
                    LayoutCachedTop =180
                    LayoutCachedWidth =16665
                    LayoutCachedHeight =435
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =15165
                    Top =495
                    Width =1500
                    Height =255
                    Name ="EmptyCell723"
                    GroupTable =6
                    LayoutCachedLeft =15165
                    LayoutCachedTop =495
                    LayoutCachedWidth =16665
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =15165
                    Top =810
                    Width =1500
                    Height =270
                    FontSize =9
                    Name ="Label726"
                    Caption ="Unit Cost"
                    GroupTable =6
                    LayoutCachedLeft =15165
                    LayoutCachedTop =810
                    LayoutCachedWidth =16665
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =18180
                    Top =420
                    Width =3300
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    Name ="SyncInvoiceToParts"
                    Caption ="SYNC Invoice/Dates to ALL Child Parts"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Log This Title IN"
                    GroupTable =7
                    RightPadding =45

                    LayoutCachedLeft =18180
                    LayoutCachedTop =420
                    LayoutCachedWidth =21480
                    LayoutCachedHeight =690
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =9
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =5
                    HoverShade =75.0
                    PressedThemeColorIndex =6
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =7
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =15120
                    Top =420
                    Width =2040
                    Height =270
                    FontSize =9
                    TabIndex =14
                    Name ="AddDrawDownBtn"
                    Caption ="Add A NEW Draw"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =15120
                    LayoutCachedTop =420
                    LayoutCachedWidth =17160
                    LayoutCachedHeight =690
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =4
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =16725
                    Top =180
                    Width =375
                    Height =255
                    Name ="EmptyCell743"
                    GroupTable =6
                    LayoutCachedLeft =16725
                    LayoutCachedTop =180
                    LayoutCachedWidth =17100
                    LayoutCachedHeight =435
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =16725
                    Top =495
                    Width =375
                    Height =255
                    Name ="EmptyCell744"
                    GroupTable =6
                    LayoutCachedLeft =16725
                    LayoutCachedTop =495
                    LayoutCachedWidth =17100
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =17160
                    Top =180
                    Width =1500
                    Height =255
                    Name ="EmptyCell751"
                    GroupTable =6
                    LayoutCachedLeft =17160
                    LayoutCachedTop =180
                    LayoutCachedWidth =18660
                    LayoutCachedHeight =435
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =17160
                    Top =495
                    Width =1500
                    Height =255
                    Name ="EmptyCell752"
                    GroupTable =6
                    LayoutCachedLeft =17160
                    LayoutCachedTop =495
                    LayoutCachedWidth =18660
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
                Begin EmptyCell
                    Left =16725
                    Top =810
                    Width =375
                    Height =270
                    Name ="EmptyCell764"
                    GroupTable =6
                    LayoutCachedLeft =16725
                    LayoutCachedTop =810
                    LayoutCachedWidth =17100
                    LayoutCachedHeight =1080
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
                End
            End
        End
        Begin Section
            Height =374
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =60
                    Top =30
                    Width =864
                    Height =255
                    FontSize =9
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =6

                    LayoutCachedLeft =60
                    LayoutCachedTop =30
                    LayoutCachedWidth =924
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2235
                    Top =30
                    Width =2115
                    Height =255
                    FontSize =9
                    TabIndex =2
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =6

                    LayoutCachedLeft =2235
                    LayoutCachedTop =30
                    LayoutCachedWidth =4350
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =990
                    Top =30
                    Width =1185
                    Height =255
                    FontSize =9
                    TabIndex =1
                    Name ="UnitUnitNum"
                    ControlSource ="UnitUnitNum"
                    GroupTable =6

                    LayoutCachedLeft =990
                    LayoutCachedTop =30
                    LayoutCachedWidth =2175
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5040
                    Top =30
                    Width =3165
                    Height =255
                    ColumnWidth =4290
                    FontSize =9
                    TabIndex =4
                    Name ="VendorREF"
                    ControlSource ="VendorREF"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName, Vendors.VendorSNm FROM Vendors WHER"
                        "E (((Vendors.VendorObsolete)=No)) ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;1800;720"
                    StatusBarText ="lookup from Vendor table"
                    GroupTable =6
                    AllowValueListEdits =0

                    LayoutCachedLeft =5040
                    LayoutCachedTop =30
                    LayoutCachedWidth =8205
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =302
                    Width =22579
                    Height =72
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
                    HorizontalAnchor =2

                    LayoutCachedTop =302
                    LayoutCachedWidth =22579
                    LayoutCachedHeight =374
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =4320
                    Left =14175
                    Top =30
                    Width =930
                    Height =255
                    FontSize =9
                    TabIndex =9
                    BackColor =15590879
                    Name ="DrawDownId"
                    ControlSource ="DrawDownId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DrawDowns.DrawDownId, DrawDowns.DrawNumber AS Draw, DrawDowns.DrawName, B"
                        "anks.BankShortName AS Bank, Banks.BankID FROM DrawDowns INNER JOIN Banks ON Draw"
                        "Downs.DrawBankId = Banks.bankID WHERE (((DrawDowns.GroupID)=[Forms]![BulkChange_"
                        "Form]![GrpID]));"
                    ColumnWidths ="0;432;1728;864;0"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    OnGotFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    Format ="Standard"
                    GroupTable =6

                    LayoutCachedLeft =14175
                    LayoutCachedTop =30
                    LayoutCachedWidth =15105
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8265
                    Top =30
                    Width =1584
                    Height =255
                    FontSize =9
                    TabIndex =5
                    Name ="UnitInvoiceNum"
                    ControlSource ="UnitInvoiceNum"
                    GroupTable =6

                    LayoutCachedLeft =8265
                    LayoutCachedTop =30
                    LayoutCachedWidth =9849
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9915
                    Top =30
                    Width =1296
                    Height =255
                    FontSize =9
                    TabIndex =6
                    Name ="UnitInvoiceDate"
                    ControlSource ="UnitInvoiceDate"
                    Format ="Short Date"
                    GroupTable =6

                    LayoutCachedLeft =9915
                    LayoutCachedTop =30
                    LayoutCachedWidth =11211
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11265
                    Top =30
                    Width =1290
                    Height =255
                    FontSize =9
                    TabIndex =7
                    Name ="UnitVendorPytDate"
                    ControlSource ="UnitVendorPytDate"
                    Format ="Short Date"
                    GroupTable =6

                    LayoutCachedLeft =11265
                    LayoutCachedTop =30
                    LayoutCachedWidth =12555
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4410
                    Top =30
                    Width =570
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    GroupTable =6

                    LayoutCachedLeft =4410
                    LayoutCachedTop =30
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12615
                    Top =30
                    Width =1500
                    Height =255
                    FontSize =9
                    TabIndex =8
                    Name ="VendorCost"
                    ControlSource ="VendorCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Sets the Actual Vendor Cost"
                    GroupTable =6

                    LayoutCachedLeft =12615
                    LayoutCachedTop =30
                    LayoutCachedWidth =14115
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =40.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15165
                    Top =30
                    Width =1500
                    Height =255
                    FontSize =9
                    TabIndex =10
                    Name ="UnitFinalCost"
                    ControlSource ="UnitFinalCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =6

                    LayoutCachedLeft =15165
                    LayoutCachedTop =30
                    LayoutCachedWidth =16665
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17160
                    Top =30
                    Width =1500
                    Height =255
                    FontSize =9
                    TabIndex =11
                    Name ="UnitIFAmt"
                    ControlSource ="UnitIFAmt"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="IRP Acct Number"
                    GroupTable =6

                    LayoutCachedLeft =17160
                    LayoutCachedTop =30
                    LayoutCachedWidth =18660
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =6
                    BackTint =40.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    CurrencySymbol ="$"
                End
                Begin EmptyCell
                    Left =16725
                    Top =30
                    Width =375
                    Height =255
                    Name ="EmptyCell763"
                    GroupTable =6
                    LayoutCachedLeft =16725
                    LayoutCachedTop =30
                    LayoutCachedWidth =17100
                    LayoutCachedHeight =285
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GroupTable =6
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
' See "BulkChange_Vendor.cls"
