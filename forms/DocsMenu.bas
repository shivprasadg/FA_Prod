Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14400
    DatasheetFontHeight =11
    ItemSuffix =137
    Left =11700
    Top =3810
    Right =27435
    Bottom =14430
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x60bc96f9350be440
    End
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    OnClickEmMacro = Begin
        Version =196611
        ColumnsShown =8
        Begin
            Action ="Requery"
            Argument ="BLCDForm"
        End
    End
    OnCurrentEmMacro = Begin
        Version =196611
        ColumnsShown =8
        Begin
            Action ="Requery"
        End
    End
    OnGotFocusEmMacro = Begin
        Version =196611
        ColumnsShown =8
        Begin
            Action ="Requery"
            Argument ="\\"
        End
    End
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
            BorderColor =12632256
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =675
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Top =480
                    Width =14400
                    Height =195
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =480
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =675
                    TabIndex =2
                    LayoutGroup =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =14
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =3
                    Width =2535
                    Height =600
                    FontSize =18
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="Documentation"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedWidth =2535
                    LayoutCachedHeight =600
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =720
                    Top =375
                    Width =1815
                    Height =240
                    FontSize =8
                    Name ="Label227"
                    Caption =" MENU"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =720
                    LayoutCachedTop =375
                    LayoutCachedWidth =2535
                    LayoutCachedHeight =615
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =12750
                    Top =120
                    Width =750
                    Height =228
                    FontSize =8
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =12750
                    LayoutCachedTop =120
                    LayoutCachedWidth =13500
                    LayoutCachedHeight =348
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
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
                Begin CommandButton
                    OverlapFlags =215
                    Left =12195
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =1
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="FAMainMenu"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"GoHome\" xmlns=\"http://schemas.microsoft.com/office/accesss"
                                "ervices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument Na"
                                "me=\"FormName\">FAMainMenu</Argum"
                        End
                        Begin
                            Comment ="_AXL:ent></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12195
                    LayoutCachedTop =120
                    LayoutCachedWidth =12750
                    LayoutCachedHeight =348
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =9855
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =270
                    Top =1924
                    Width =4926
                    Height =375
                    TabIndex =4
                    Name ="Command5"
                    Caption ="All Clients && Targets"
                    ControlTipText ="Open Form"
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="Clients_view"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command5\" xmlns=\"http://schemas.microsoft.com/office/acces"
                                "sservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument "
                                "Name=\"FormName\">Clients_view</A"
                        End
                        Begin
                            Comment ="_AXL:rgument></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =270
                    LayoutCachedTop =1924
                    LayoutCachedWidth =5196
                    LayoutCachedHeight =2299
                    RowStart =4
                    RowEnd =4
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
                Begin CommandButton
                    OverlapFlags =93
                    Left =270
                    Top =574
                    Width =4926
                    Height =381
                    TabIndex =1
                    Name ="Command7"
                    Caption ="Bank Maintenance"
                    ControlTipText ="Open Form"
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="Bank_View"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command7\" xmlns=\"http://schemas.microsoft.com/office/acces"
                                "sservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument "
                                "Name=\"FormName\">Bank_View</Argu"
                        End
                        Begin
                            Comment ="_AXL:ment></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =270
                    LayoutCachedTop =574
                    LayoutCachedWidth =5196
                    LayoutCachedHeight =955
                    RowStart =1
                    RowEnd =1
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
                Begin CommandButton
                    OverlapFlags =93
                    Left =270
                    Top =2820
                    Width =4926
                    Height =375
                    TabIndex =6
                    Name ="OpenUnitLookup"
                    Caption ="Unit Lookup"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open Form"
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =270
                    LayoutCachedTop =2820
                    LayoutCachedWidth =5196
                    LayoutCachedHeight =3195
                    RowStart =6
                    RowEnd =6
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
                Begin CommandButton
                    OverlapFlags =95
                    Left =270
                    Top =1476
                    Width =4926
                    Height =375
                    TabIndex =3
                    Name ="Command11"
                    Caption ="Sales Tax Table"
                    ControlTipText ="Open Form"
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="StTaxf"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command11\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument"
                                " Name=\"FormName\">StTaxf</Argume"
                        End
                        Begin
                            Comment ="_AXL:nt></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =270
                    LayoutCachedTop =1476
                    LayoutCachedWidth =5196
                    LayoutCachedHeight =1851
                    RowStart =3
                    RowEnd =3
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
                Begin CommandButton
                    OverlapFlags =93
                    Left =270
                    Top =3690
                    Width =4926
                    Height =390
                    TabIndex =7
                    Name ="Command15"
                    Caption ="Customer Summary Delivery "
                    ControlTipText ="Open Form"
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenReport"
                            Argument ="DeliveryRptSum"
                            Argument ="2"
                            Argument =""
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command15\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenReport\"><Argume"
                                "nt Name=\"ReportName\">DeliveryRp"
                        End
                        Begin
                            Comment ="_AXL:tSum</Argument><Argument Name=\"View\">Print Preview</Argument></Action></S"
                                "tatements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =270
                    LayoutCachedTop =3690
                    LayoutCachedWidth =5196
                    LayoutCachedHeight =4080
                    RowStart =8
                    RowEnd =8
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
                Begin CommandButton
                    OverlapFlags =93
                    Left =270
                    Top =4593
                    Width =4926
                    Height =381
                    TabIndex =9
                    Name ="Command17"
                    Caption ="Delivery Report by Vendor"
                    ControlTipText ="Open Form"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenReport"
                            Argument ="DeliveryRptDet"
                            Argument ="2"
                            Argument =""
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command17\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenReport\"><Argume"
                                "nt Name=\"ReportName\">DeliveryRp"
                        End
                        Begin
                            Comment ="_AXL:tDet</Argument><Argument Name=\"View\">Print Preview</Argument></Action></S"
                                "tatements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =270
                    LayoutCachedTop =4593
                    LayoutCachedWidth =5196
                    LayoutCachedHeight =4974
                    RowStart =10
                    RowEnd =10
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
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =163
                    TextAlign =2
                    Left =2985
                    Top =3255
                    Width =4926
                    Height =375
                    FontSize =14
                    FontWeight =700
                    ForeColor =6108695
                    Name ="Label21"
                    Caption ="Reports"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2985
                    LayoutCachedTop =3255
                    LayoutCachedWidth =7911
                    LayoutCachedHeight =3630
                    RowStart =7
                    RowEnd =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =5775
                    Top =4593
                    Width =4926
                    Height =390
                    TabIndex =18
                    Name ="Command28"
                    Caption ="Bank Support Bill"
                    ControlTipText ="Open Form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="PD BillSel"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command28\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument"
                                " Name=\"FormName\">PD BillSel</Ar"
                        End
                        Begin
                            Comment ="_AXL:gument></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5775
                    LayoutCachedTop =4593
                    LayoutCachedWidth =10701
                    LayoutCachedHeight =4983
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
                Begin CommandButton
                    OverlapFlags =223
                    Left =270
                    Top =5040
                    Width =4926
                    Height =381
                    TabIndex =10
                    Name ="Command30"
                    Caption =" Outstanding CofAs"
                    ControlTipText ="Open Form"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenReport"
                            Argument ="OUTSTANDING CofAs"
                            Argument ="2"
                            Argument =""
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command30\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenReport\"><Argume"
                                "nt Name=\"ReportName\">OUTSTANDIN"
                        End
                        Begin
                            Comment ="_AXL:G CofAs</Argument><Argument Name=\"View\">Print Preview</Argument></Action>"
                                "</Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =270
                    LayoutCachedTop =5040
                    LayoutCachedWidth =5196
                    LayoutCachedHeight =5421
                    RowStart =11
                    RowEnd =11
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =5775
                    Top =3690
                    Width =4926
                    Height =390
                    TabIndex =17
                    Name ="ExipiringInsExcel"
                    Caption ="Expiring Insurance"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open Form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5775
                    LayoutCachedTop =3690
                    LayoutCachedWidth =10701
                    LayoutCachedHeight =4080
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
                Begin CommandButton
                    OverlapFlags =223
                    Left =5775
                    Top =5040
                    Width =4926
                    Height =390
                    TabIndex =19
                    Name ="Command36"
                    Caption ="Schedule Summary"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="SSumRpt"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command36\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument"
                                " Name=\"FormName\">SSumRpt</Argum"
                        End
                        Begin
                            Comment ="_AXL:ent></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5775
                    LayoutCachedTop =5040
                    LayoutCachedWidth =10701
                    LayoutCachedHeight =5430
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
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =75
                    Top =5445
                    Width =13935
                    Height =375
                    FontSize =14
                    FontWeight =700
                    ForeColor =6108695
                    Name ="Label40"
                    Caption ="Queries"
                    LayoutCachedLeft =75
                    LayoutCachedTop =5445
                    LayoutCachedWidth =14010
                    LayoutCachedHeight =5820
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =375
                    Top =6399
                    Width =1371
                    Height =366
                    TabIndex =22
                    Name ="Command41"
                    ControlTipText ="Open Form"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenQuery"
                            Argument ="CustomerUnitLeaseTermsq"
                            Argument ="0"
                            Argument ="2"
                        End
                        Begin
                            Action ="OutputTo"
                            Argument ="1"
                            Argument ="CustomerUnitLeaseTermsq"
                            Argument ="Excel97-Excel2003Workbook(*.xls)"
                            Argument =""
                            Argument ="0"
                            Argument =""
                            Argument ="0"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command41\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenQuery\"><Argumen"
                                "t Name=\"QueryName\">CustomerUnit"
                        End
                        Begin
                            Comment ="_AXL:LeaseTermsq</Argument><Argument Name=\"DataMode\">Read Only</Argument></Act"
                                "ion><Action Name=\"ExportWithFormatting\"><Argument Name=\"ObjectType\">Query</A"
                                "rgument><Argument Name=\"ObjectName\">CustomerUnitLeaseTermsq</Argument><Argumen"
                                "t Name=\"OutputFormat\""
                        End
                        Begin
                            Comment ="_AXL:>Excel97-Excel2003Workbook(*.xls)</Argument><Argument Name=\"Encoding\">0</"
                                "Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =375
                    LayoutCachedTop =6399
                    LayoutCachedWidth =1746
                    LayoutCachedHeight =6765
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            Left =1980
                            Top =6390
                            Width =4350
                            Height =375
                            FontSize =14
                            FontWeight =700
                            ForeColor =-2147483617
                            Name ="Label70"
                            Caption ="Customer Lease Terms Spreadsheet"
                            LayoutCachedLeft =1980
                            LayoutCachedTop =6390
                            LayoutCachedWidth =6330
                            LayoutCachedHeight =6765
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =163
                    Left =7965
                    Top =2655
                    Width =6240
                    Height =360
                    FontSize =10
                    FontWeight =700
                    ForeColor =5026082
                    Name ="Label42"
                    Caption ="Use Mgmt Menu for: Schedule Track, Syndication Inventory, Syndication YTD"
                    LayoutCachedLeft =7965
                    LayoutCachedTop =2655
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =3015
                End
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =2370
                    Top =5940
                    Width =3825
                    Height =375
                    TabIndex =20
                    Name ="SelCust"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Clients.clientID, Clients.ClientShNm FROM Clients INNER JOIN Mst"
                        "rLease ON Clients.clientID = MstrLease.Client WHERE (((Clients.ClientShNm) Not L"
                        "ike \"demo*\") AND ((Clients.clientType)=\"client\" Or (Clients.clientType)=\"Gu"
                        "arantor\") AND ((MstrLease.MLOrig)=\"fa\") AND ((MstrLease.MLNo) Not Like \"*Rem"
                        "arketint\")) ORDER BY Clients.ClientShNm;"
                    ColumnWidths ="0;2880"

                    LayoutCachedLeft =2370
                    LayoutCachedTop =5940
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =6315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =163
                            TextAlign =3
                            Left =615
                            Top =5940
                            Width =1680
                            Height =375
                            FontWeight =700
                            ForeColor =-2147483617
                            Name ="Label44"
                            Caption ="Select Customer:"
                            LayoutCachedLeft =615
                            LayoutCachedTop =5940
                            LayoutCachedWidth =2295
                            LayoutCachedHeight =6315
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7725
                    Top =5949
                    Width =1371
                    Height =366
                    TabIndex =21
                    Name ="Command47"
                    ControlTipText ="Open Form"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenQuery"
                            Argument ="Synagingq2"
                            Argument ="0"
                            Argument ="2"
                        End
                        Begin
                            Action ="OutputTo"
                            Argument ="1"
                            Argument ="Synagingq2"
                            Argument ="Excel97-Excel2003Workbook(*.xls)"
                            Argument =""
                            Argument ="0"
                            Argument =""
                            Argument ="0"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command47\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenQuery\"><Argumen"
                                "t Name=\"QueryName\">Synagingq2</"
                        End
                        Begin
                            Comment ="_AXL:Argument><Argument Name=\"DataMode\">Read Only</Argument></Action><Action N"
                                "ame=\"ExportWithFormatting\"><Argument Name=\"ObjectType\">Query</Argument><Argu"
                                "ment Name=\"ObjectName\">Synagingq2</Argument><Argument Name=\"OutputFormat\">Ex"
                                "cel97-Excel2003Workbook"
                        End
                        Begin
                            Comment ="_AXL:(*.xls)</Argument><Argument Name=\"Encoding\">0</Argument></Action></Statem"
                                "ents></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7725
                    LayoutCachedTop =5949
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =6315
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =9360
                    Top =6006
                    Width =4350
                    Height =375
                    FontSize =14
                    FontWeight =700
                    ForeColor =-2147483617
                    Name ="Label48"
                    Caption ="Non-Syndicated Lease Aging-Excel"
                    LayoutCachedLeft =9360
                    LayoutCachedTop =6006
                    LayoutCachedWidth =13710
                    LayoutCachedHeight =6381
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7725
                    Top =6399
                    Width =1371
                    Height =366
                    TabIndex =23
                    Name ="Command49"
                    ControlTipText ="Open Form"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenReport"
                            Argument ="SynagingRpt"
                            Argument ="2"
                            Argument =""
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command49\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenReport\"><Argume"
                                "nt Name=\"ReportName\">SynagingRp"
                        End
                        Begin
                            Comment ="_AXL:t</Argument><Argument Name=\"View\">Print Preview</Argument></Action></Stat"
                                "ements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7725
                    LayoutCachedTop =6399
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =6765
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =9360
                    Top =6396
                    Width =4350
                    Height =375
                    FontSize =14
                    FontWeight =700
                    ForeColor =-2147483617
                    Name ="Label50"
                    Caption ="Non-Syndicated Lease Aging-Report"
                    LayoutCachedLeft =9360
                    LayoutCachedTop =6396
                    LayoutCachedWidth =13710
                    LayoutCachedHeight =6771
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =270
                    Top =1028
                    Width =4926
                    Height =375
                    TabIndex =2
                    Name ="Command51"
                    Caption ="Search Clients && Targets"
                    ControlTipText ="Open Form"
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="Search_clients"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command51\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument"
                                " Name=\"FormName\">Search_clients"
                        End
                        Begin
                            Comment ="_AXL:</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =270
                    LayoutCachedTop =1028
                    LayoutCachedWidth =5196
                    LayoutCachedHeight =1403
                    RowStart =2
                    RowEnd =2
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
                Begin CommandButton
                    OverlapFlags =87
                    Left =270
                    Top =120
                    Width =4926
                    Height =381
                    Name ="ScheduleDashboard"
                    Caption =" Schedule Dashboard"
                    StatusBarText ="Schedule Dashboard"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Schedule Dashboard"
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff727272ff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff727272ff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xffffffff727272ff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffb3b3b3ffffffffffffffffffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffff727272ff727272ffffffffffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xffffffffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xffffffff727272ff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff727272ff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xffffffff727272ff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffb3b3b3ffffffffffffffffffffffffffffffffffffffffffb3b3b3ff ,
                        0xffffffff727272ff727272ffffffffffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xffffffffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xffffffff727272ff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff727272ffb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =270
                    LayoutCachedTop =120
                    LayoutCachedWidth =5196
                    LayoutCachedHeight =501
                    PictureCaptionArrangement =5
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
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =10605
                    Top =495
                    Width =3180
                    Height =315
                    TabIndex =11
                    BoundColumn =1
                    Name ="Combo59"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_Title_Transmital.clientID, qry_Title_Transmital.ClientShNm, qry_Title"
                        "_Transmital.clientCompanyName FROM qry_Title_Transmital GROUP BY qry_Title_Trans"
                        "mital.clientID, qry_Title_Transmital.ClientShNm, qry_Title_Transmital.clientComp"
                        "anyName HAVING (((qry_Title_Transmital.ClientShNm) Not Like \"*demo*\"));"
                    ColumnWidths ="0;2880;0;0;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =10605
                    LayoutCachedTop =495
                    LayoutCachedWidth =13785
                    LayoutCachedHeight =810
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =7470
                            Top =495
                            Width =2955
                            Height =390
                            FontSize =14
                            FontWeight =700
                            Name ="Label60"
                            Caption ="Select Customer/Lessee"
                            LayoutCachedLeft =7470
                            LayoutCachedTop =495
                            LayoutCachedWidth =10425
                            LayoutCachedHeight =885
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =10605
                    Top =945
                    Width =3180
                    Height =315
                    TabIndex =12
                    BoundColumn =1
                    Name ="Combo61"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_Title_Transmital.clientID, MstrLease.MLNo, MstrLease.MLOrig, qry_Titl"
                        "e_Transmital.ClientShNm, qry_Title_Transmital.clientCompanyName FROM qry_Title_T"
                        "ransmital INNER JOIN MstrLease ON qry_Title_Transmital.clientID = MstrLease.Clie"
                        "nt GROUP BY qry_Title_Transmital.clientID, MstrLease.MLNo, MstrLease.MLOrig, qry"
                        "_Title_Transmital.ClientShNm, qry_Title_Transmital.clientCompanyName HAVING (((M"
                        "strLease.MLOrig) Like \"FA\") AND ((qry_Title_Transmital.ClientShNm) Like [Forms"
                        "]![DocsMenu]![Combo59]) AND ((qry_Title_Transmital.clientCompanyName) Not Like \""
                        "*demo*\"));"
                    ColumnWidths ="0;2880;0;0"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Requery"
                            Argument ="combo61"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Combo61\" xmlns=\"http://schemas.microsoft.com/office/access"
                                "services/2009/11/application\"><Statements><Action Name=\"Requery\"><Argument Na"
                                "me=\"ControlName\">combo61</Argum"
                        End
                        Begin
                            Comment ="_AXL:ent></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =10605
                    LayoutCachedTop =945
                    LayoutCachedWidth =13785
                    LayoutCachedHeight =1260
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =7425
                            Top =945
                            Width =3000
                            Height =315
                            FontSize =14
                            FontWeight =700
                            Name ="Label62"
                            Caption ="Select Master Lease"
                            LayoutCachedLeft =7425
                            LayoutCachedTop =945
                            LayoutCachedWidth =10425
                            LayoutCachedHeight =1260
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =10605
                    Top =1410
                    Width =3180
                    Height =315
                    TabIndex =13
                    BoundColumn =1
                    Name ="Combo63"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MstrLease.MLNo, qry_Title_Transmital.SchNo, qry_Title_Transmital.clientID"
                        ", qry_Title_Transmital.ClientShNm, qry_Title_Transmital.MLOrig, qry_Title_Transm"
                        "ital.ClientShNm FROM qry_Title_Transmital INNER JOIN MstrLease ON qry_Title_Tran"
                        "smital.clientID = MstrLease.Client GROUP BY MstrLease.MLNo, qry_Title_Transmital"
                        ".SchNo, qry_Title_Transmital.clientID, qry_Title_Transmital.ClientShNm, qry_Titl"
                        "e_Transmital.MLOrig, qry_Title_Transmital.ClientShNm HAVING (((MstrLease.MLNo)=["
                        "Forms]![DocsMenu]![Combo61]) AND ((qry_Title_Transmital.ClientShNm) Not Like \"*"
                        "demo*\" And (qry_Title_Transmital.ClientShNm)=[Forms]![DocsMenu]![Combo59]) AND "
                        "((qry_Title_Transmital.MLOrig) Like \"fa\") AND ((qry_Title_Transmital.ClientShN"
                        "m)=[Forms]![DocsMenu]![Combo59]));"
                    ColumnWidths ="0;2880;0;0;0"
                    OnGotFocus ="[Event Procedure]"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Requery"
                            Argument ="combo63"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Combo63\" xmlns=\"http://schemas.microsoft.com/office/access"
                                "services/2009/11/application\"><Statements><Action Name=\"Requery\"><Argument Na"
                                "me=\"ControlName\">combo63</Argum"
                        End
                        Begin
                            Comment ="_AXL:ent></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =10605
                    LayoutCachedTop =1410
                    LayoutCachedWidth =13785
                    LayoutCachedHeight =1725
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =7650
                            Top =1410
                            Width =2775
                            Height =315
                            FontSize =14
                            FontWeight =700
                            Name ="Label64"
                            Caption ="Schedule Number"
                            LayoutCachedLeft =7650
                            LayoutCachedTop =1410
                            LayoutCachedWidth =10425
                            LayoutCachedHeight =1725
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    Left =8145
                    Top =60
                    Width =4980
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="Label65"
                    Caption ="Title Transmital Excel Export"
                    LayoutCachedLeft =8145
                    LayoutCachedTop =60
                    LayoutCachedWidth =13125
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12285
                    Top =2205
                    TabIndex =16
                    Name ="Command66"
                    Caption ="Export"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =12285
                    LayoutCachedTop =2205
                    LayoutCachedWidth =13725
                    LayoutCachedHeight =2565
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10665
                    Top =2205
                    TabIndex =15
                    Name ="Command67"
                    Caption ="Clear"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10665
                    LayoutCachedTop =2205
                    LayoutCachedWidth =12105
                    LayoutCachedHeight =2565
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10605
                    Top =1815
                    Width =3180
                    Height =315
                    TabIndex =14
                    Name ="Combo68"
                    RowSourceType ="Value List"
                    RowSource ="\"Equipment Files\";\"Data Files\""
                    ColumnWidths ="2880"
                    AllowValueListEdits =0

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =10605
                    LayoutCachedTop =1815
                    LayoutCachedWidth =13785
                    LayoutCachedHeight =2130
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =7650
                            Top =1815
                            Width =2775
                            Height =360
                            FontSize =14
                            FontWeight =700
                            Name ="Label69"
                            Caption ="File Type"
                            LayoutCachedLeft =7650
                            LayoutCachedTop =1815
                            LayoutCachedWidth =10425
                            LayoutCachedHeight =2175
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =270
                    Top =4137
                    Width =4926
                    Height =390
                    TabIndex =8
                    Name ="Command14"
                    Caption ="Delivery Report for Selected Customer"
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="CustDelSel"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command14\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument"
                                " Name=\"FormName\">CustDelSel</Ar"
                        End
                        Begin
                            Comment ="_AXL:gument></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =270
                    LayoutCachedTop =4137
                    LayoutCachedWidth =5196
                    LayoutCachedHeight =4527
                    RowStart =9
                    RowEnd =9
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
                Begin CommandButton
                    OverlapFlags =215
                    Left =270
                    Top =2372
                    Width =4926
                    Height =375
                    TabIndex =5
                    Name ="TitleInOutBtn"
                    Caption ="Title Check In/Out"
                    LeftPadding =0
                    TopPadding =15
                    RightPadding =0
                    BottomPadding =120
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="TitleCheckin"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"TitleInOutBtn\" xmlns=\"http://schemas.microsoft.com/office/"
                                "accessservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argu"
                                "ment Name=\"FormName\">TitleCheck"
                        End
                        Begin
                            Comment ="_AXL:in</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =270
                    LayoutCachedTop =2372
                    LayoutCachedWidth =5196
                    LayoutCachedHeight =2747
                    RowStart =5
                    RowEnd =5
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
            End
        End
        Begin FormFooter
            Height =105
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "DocsMenu.cls"
