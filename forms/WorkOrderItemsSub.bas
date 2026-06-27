Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =1
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12585
    DatasheetFontHeight =11
    ItemSuffix =536
    Left =7515
    Top =3555
    Right =19905
    Bottom =7950
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xcc1671b93331e640
    End
    RecordSource ="SELECT * FROM WorkOrderItems"
    OnCurrent ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =480
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =240
                    Width =7800
                    Height =210
                    BorderColor =16777215
                    Name ="Image142"
                    Picture ="1_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =240
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =450
                    TabIndex =4
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =120
                    Width =2160
                    Height =300
                    FontSize =12
                    FontWeight =700
                    Name ="Label1053"
                    Caption ="Work Order Items"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =119
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3180
                    Width =360
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="WorkOrderID"
                    ControlSource ="WorkOrderID"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =3180
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =240
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =6
                    ForeTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2340
                    Width =780
                    Height =270
                    FontSize =8
                    FontWeight =700
                    Name ="wiLineItemID"
                    ControlSource ="wiLineItemID"
                    StatusBarText ="unit link"
                    LeftPadding =90
                    RightPadding =105

                    LayoutCachedLeft =2340
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =270
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =6
                    ForeTint =60.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =8100
                    Top =90
                    Height =240
                    FontSize =10
                    TabIndex =3
                    ForeColor =0
                    Name ="AddNew"
                    Caption ="Add An Item"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =8100
                    LayoutCachedTop =90
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =330
                    ColumnStart =3
                    ColumnEnd =3
                    ForeTint =100.0
                    UseTheme =1
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =9
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =119
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6630
                    Width =1170
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    ForeColor =10040879
                    Name ="ExpenseSum"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =6630
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =240
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
                            OverlapFlags =127
                            TextAlign =1
                            Left =5610
                            Width =1020
                            Height =240
                            FontSize =8
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label573"
                            Caption ="Total Cost"
                            LayoutCachedLeft =5610
                            LayoutCachedWidth =6630
                            LayoutCachedHeight =240
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
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =2
                    Left =9600
                    Top =29
                    Width =2985
                    Height =285
                    FontSize =10
                    BackColor =-2147483633
                    Name ="EnterMsg"
                    Caption ="▼ Type VMRSCode / Desc  HERE▼"
                    LayoutCachedLeft =9600
                    LayoutCachedTop =29
                    LayoutCachedWidth =12585
                    LayoutCachedHeight =314
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    Visible = NotDefault
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =10350
                    Top =240
                    Width =1485
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="DDClick"
                    Caption ="Double-Click to Add"
                    LayoutCachedLeft =10350
                    LayoutCachedTop =240
                    LayoutCachedWidth =11835
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =9
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            Height =4200
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =705
                    Width =5280
                    FontSize =10
                    TabIndex =2
                    Name ="wiDesc"
                    ControlSource ="wiDesc"

                    LayoutCachedLeft =60
                    LayoutCachedTop =705
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =945
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =3
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =1005
                    Width =5280
                    Height =690
                    FontSize =10
                    TabIndex =3
                    Name ="wiNotes"
                    ControlSource ="wiNotes"

                    LayoutCachedLeft =60
                    LayoutCachedTop =1005
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =1695
                    RowStart =3
                    RowEnd =4
                    ColumnEnd =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =60
                    Top =1860
                    Width =5454
                    Height =2280
                    FontSize =9
                    TabIndex =8
                    Name ="ItemList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT WorkOrderItems.wiLineItemID, WorkOrderItems.wiVMRSCode AS VMRS, WorkOrder"
                        "Items.WiDesc AS [Desc], CCur(Nz([wiPartsCost],0)+Nz([wiLaborCost],0)) AS Cost, W"
                        "orkOrderItems.wiStatus AS Status FROM WorkOrderItems WHERE (((WorkOrderItems.Wor"
                        "kOrderID)=[Forms]![UnitsForm]![WorkOrdersForm].[Form]![WorkOrderID]));"
                    ColumnWidths ="0;1152;2880;1008;864"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =1860
                    LayoutCachedWidth =5514
                    LayoutCachedHeight =4140
                    BorderThemeColorIndex =9
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =6660
                    Top =2670
                    Width =1260
                    Height =315
                    TabIndex =10
                    Name ="wiPartsCost"
                    ControlSource ="wiPartsCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    DefaultValue ="0"

                    LayoutCachedLeft =6660
                    LayoutCachedTop =2670
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =2985
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =2
                    BackTint =20.0
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =5640
                            Top =2670
                            Width =960
                            Height =315
                            FontSize =10
                            Name ="Label179"
                            Caption ="Part Cost"
                            LayoutCachedLeft =5640
                            LayoutCachedTop =2670
                            LayoutCachedWidth =6600
                            LayoutCachedHeight =2985
                            RowStart =5
                            RowEnd =5
                            ColumnStart =1
                            ColumnEnd =1
                            BackThemeColorIndex =4
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6420
                    Top =3360
                    Width =1380
                    Height =255
                    FontSize =10
                    TabIndex =12
                    Name ="DeleteMe"
                    Caption ="Delete Item"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6420
                    LayoutCachedTop =3360
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =3615
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =2
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    UseTheme =1
                    Shape =0
                    Gradient =0
                    BackColor =1643706
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6300
                    Top =1200
                    Width =1380
                    Height =255
                    FontSize =10
                    TabIndex =11
                    Name ="SaveMe"
                    Caption ="Update / Save "
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6300
                    LayoutCachedTop =1200
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =1455
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =2
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    UseTheme =1
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =2
                    BackTint =100.0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6660
                    Top =1920
                    Width =1260
                    Height =315
                    TabIndex =7
                    Name ="wiLaborHours"
                    ControlSource ="wiLaborHours"
                    DefaultValue ="0"

                    LayoutCachedLeft =6660
                    LayoutCachedTop =1920
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =2235
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    BackShade =95.0
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =5640
                            Top =1920
                            Width =960
                            Height =315
                            FontSize =10
                            Name ="Label287"
                            Caption ="Labor Hrs"
                            LayoutCachedLeft =5640
                            LayoutCachedTop =1920
                            LayoutCachedWidth =6600
                            LayoutCachedHeight =2235
                            RowStart =3
                            RowEnd =3
                            ColumnStart =1
                            ColumnEnd =1
                            BackThemeColorIndex =1
                            BackShade =50.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =6660
                    Top =2295
                    Width =1260
                    Height =315
                    TabIndex =9
                    BackColor =14151142
                    Name ="wiLaborCost"
                    ControlSource ="wiLaborCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    DefaultValue ="0"

                    LayoutCachedLeft =6660
                    LayoutCachedTop =2295
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =2610
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =5640
                            Top =2295
                            Width =960
                            Height =315
                            FontSize =10
                            BackColor =14151142
                            Name ="Label293"
                            Caption ="Labor Cost"
                            LayoutCachedLeft =5640
                            LayoutCachedTop =2295
                            LayoutCachedWidth =6600
                            LayoutCachedHeight =2610
                            RowStart =4
                            RowEnd =4
                            ColumnStart =1
                            ColumnEnd =1
                            ForeThemeColorIndex =9
                            ForeShade =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =6480
                    Left =4920
                    Top =375
                    Width =3000
                    Height =255
                    FontSize =10
                    TabIndex =4
                    Name ="wiVendorID"
                    ControlSource ="wiVendorID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorSNm, Vendors.VendorName FROM Vendors WHER"
                        "E (((Vendors.VendorObsolete)=False)) ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;1440;4320"

                    LayoutCachedLeft =4920
                    LayoutCachedTop =375
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =630
                    RowStart =6
                    RowEnd =6
                    ColumnEnd =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =360
                    Width =4485
                    Height =288
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="wiVMRSCode"
                    ControlSource ="wiVMRSCode"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Expense type is required"
                    ControlTipText ="Double click to add expsenst type"
                    OnDblClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="AddRmktExpForm"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"wiVMRSCode\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argumen"
                                "t Name=\"FormName\">AddRmktExpFor"
                        End
                        Begin
                            Comment ="_AXL:m</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =60
                    LayoutCachedTop =360
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =648
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =2
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =95
                    TextAlign =1
                    DecimalPlaces =2
                    IMESentenceMode =3
                    Left =6300
                    Top =705
                    Width =1620
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="wiStatus"
                    ControlSource ="wiStatus"
                    RowSourceType ="Value List"
                    RowSource ="\"Pending\";\"Complete\""
                    OnDblClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="AddRmktExpForm"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"wiStatus\" xmlns=\"http://schemas.microsoft.com/office/acces"
                                "sservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument "
                                "Name=\"FormName\">AddRmktExpForm<"
                        End
                        Begin
                            Comment ="_AXL:/Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =6300
                    LayoutCachedTop =705
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =945
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =95
                            TextAlign =2
                            Left =5340
                            Top =705
                            Width =960
                            Height =240
                            Name ="Label426"
                            Caption ="Status"
                            LayoutCachedLeft =5340
                            LayoutCachedTop =705
                            LayoutCachedWidth =6300
                            LayoutCachedHeight =945
                            RowStart =7
                            RowEnd =7
                            BackThemeColorIndex =4
                            BackShade =75.0
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7920
                    Top =360
                    Width =4554
                    Height =3180
                    FontSize =9
                    TabIndex =6
                    Name ="VMRSList"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="1008;3600"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =7920
                    LayoutCachedTop =360
                    LayoutCachedWidth =12474
                    LayoutCachedHeight =3540
                    BorderThemeColorIndex =9
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7920
                    Top =60
                    Width =4554
                    Height =300
                    FontSize =10
                    BackColor =13434879
                    Name ="VMRSSearch"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =7920
                    LayoutCachedTop =60
                    LayoutCachedWidth =12474
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =60
                    Width =4485
                    Height =240
                    FontSize =12
                    Name ="Label417"
                    Caption ="VMRS Code and Desc and Notes"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =300
                    ColumnEnd =2
                    ForeThemeColorIndex =9
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =4920
                    Top =60
                    Width =3000
                    Height =255
                    FontSize =10
                    Name ="Label430"
                    Caption ="Vendor"
                    LayoutCachedLeft =4920
                    LayoutCachedTop =60
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =315
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =1
                    ForeThemeColorIndex =9
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =8040
                    Top =900
                    Width =60
                    Height =2340
                    BackColor =12566463
                    Name ="Box535"
                    LayoutCachedLeft =8040
                    LayoutCachedTop =900
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =3240
                    BackShade =75.0
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
' See "WorkOrderItemsSub.cls"
