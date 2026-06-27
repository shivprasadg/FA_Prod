Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18720
    DatasheetFontHeight =11
    ItemSuffix =170
    Left =5340
    Top =2775
    Right =26535
    Bottom =15420
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x8da3e92d8d7de540
    End
    RecordSource ="SELECT * FROM RMKT_Invoices WHERE (((RMKT_Invoices.[BuyerID])=0) AND ((RMKT_Invo"
        "ices.[InvoiceRequestID])=0));"
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
            Height =1290
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =1110
                    Width =18720
                    Height =180
                    Name ="Image142"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =1110
                    LayoutCachedWidth =18720
                    LayoutCachedHeight =1290
                    TabIndex =7
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =13410
                    Top =405
                    Width =1620
                    Height =270
                    FontSize =10
                    Name ="AddBill"
                    Caption ="Add an Invoice"
                    StatusBarText ="A Bill for any Remarketing Buyer can be added by clicking this button"
                    ControlTipText ="A Bill for any Remarketing Buyer can be added by clicking this button"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13410
                    LayoutCachedTop =405
                    LayoutCachedWidth =15030
                    LayoutCachedHeight =675
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
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =225
                    Top =855
                    Width =7155
                    Height =285
                    ColumnOrder =0
                    TabIndex =1
                    Name ="BuyerRefCBO"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Buyers.BuyerID, Buyers.BuyerName, Buyers.BuyerShortName, Count(Rmkt.UnitR"
                        "ef) AS Units FROM Rmkt RIGHT JOIN Buyers ON Rmkt.BuyerRef = Buyers.clientID GROU"
                        "P BY Buyers.BuyerID, Buyers.BuyerName, Buyers.BuyerShortName ORDER BY Buyers.Buy"
                        "erName;"
                    ColumnWidths ="0;3600;2448;576"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =225
                    LayoutCachedTop =855
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =1140
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =1
                            Left =210
                            Top =540
                            Width =2085
                            Height =315
                            FontWeight =700
                            Name ="Label14"
                            Caption ="Select A Buyer"
                            LayoutCachedLeft =210
                            LayoutCachedTop =540
                            LayoutCachedWidth =2295
                            LayoutCachedHeight =855
                            BackThemeColorIndex =1
                            BackShade =85.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7660
                    Top =420
                    Width =1049
                    Height =435
                    ColumnOrder =2
                    FontSize =20
                    FontWeight =700
                    TabIndex =2
                    BorderColor =1643706
                    Name ="UnitCount"
                    ControlSource ="=[ReMarketInvoiceSub].[Form]![UnitCount]"

                    LayoutCachedLeft =7660
                    LayoutCachedTop =420
                    LayoutCachedWidth =8709
                    LayoutCachedHeight =855
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =2
                            Left =7635
                            Top =210
                            Width =1110
                            Height =675
                            FontSize =9
                            FontWeight =700
                            BorderColor =-2147483615
                            ForeColor =16777215
                            Name ="UnitCountLabel"
                            Caption =" Units"
                            LayoutCachedLeft =7635
                            LayoutCachedTop =210
                            LayoutCachedWidth =8745
                            LayoutCachedHeight =885
                            BackThemeColorIndex =7
                            BackShade =75.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =11706
                    Top =405
                    Width =1671
                    Height =270
                    FontSize =10
                    TabIndex =3
                    Name ="BillLog"
                    Caption ="Print Invoice Log"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11706
                    LayoutCachedTop =405
                    LayoutCachedWidth =13377
                    LayoutCachedHeight =675
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
                Begin CommandButton
                    OverlapFlags =215
                    Left =10005
                    Top =405
                    Width =1671
                    Height =270
                    FontSize =10
                    TabIndex =4
                    Name ="ExportLog2Excel"
                    Caption ="Export Log"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddd87262ddddddd876262626dddddd6262626262262628 ,
                        0x2626262626fffff662ff6ff2626f62f2268f2f8626fffff6626fff62626f62f2 ,
                        0x2628f82626fffff66268f862626f62f2262fff2626fffff6628f6f82626f62f2 ,
                        0x26ff2ff626fffff662626262626f62f22626262626fffff6d872626262626268 ,
                        0xddddd87626dddddd000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10005
                    LayoutCachedTop =405
                    LayoutCachedWidth =11676
                    LayoutCachedHeight =675
                    PictureCaptionArrangement =5
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
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =120
                    Width =3915
                    Height =495
                    FontSize =16
                    Name ="Label465"
                    Caption ="Remarketing Buyers Orders"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedWidth =4035
                    LayoutCachedHeight =495
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =2265
                    Top =405
                    Width =915
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label227"
                    Caption ="Manager"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2265
                    LayoutCachedTop =405
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =645
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =17715
                    Top =150
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =5
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =17715
                    LayoutCachedTop =150
                    LayoutCachedWidth =18465
                    LayoutCachedHeight =378
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
                    Left =17160
                    Top =150
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =6
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

                    LayoutCachedLeft =17160
                    LayoutCachedTop =150
                    LayoutCachedWidth =17715
                    LayoutCachedHeight =378
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
            CanShrink = NotDefault
            Height =8790
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9570
                    Top =915
                    Height =315
                    TabIndex =7
                    BorderColor =-2147483615
                    Name ="SP"
                    ControlSource ="=DLookUp(\"SP\",\"RMKT_BuyersOrders\",\"BuyersOrderID = \" & [InvoiceRequestID] "
                        "& \"\")"
                    GroupTable =7
                    RightPadding =45
                    BottomPadding =45

                    LayoutCachedLeft =9570
                    LayoutCachedTop =915
                    LayoutCachedWidth =11010
                    LayoutCachedHeight =1230
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =8370
                            Top =915
                            Width =1140
                            Height =315
                            FontSize =10
                            FontWeight =700
                            BorderColor =-2147483615
                            Name ="Label154"
                            Caption ="Sales Person:"
                            GroupTable =7
                            BottomPadding =45
                            LayoutCachedLeft =8370
                            LayoutCachedTop =915
                            LayoutCachedWidth =9510
                            LayoutCachedHeight =1230
                            RowStart =2
                            RowEnd =2
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =7
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =90
                    Top =315
                    Width =3090
                    Height =1530
                    FontSize =9
                    TabIndex =8
                    ForeColor =-2147483615
                    BorderColor =-2147483615
                    Name ="InvoiceList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT RmktInvoiceRequestTbl.InvoiceRequestID AS Req, Nz(InvoiceNumber,'Request'"
                        ") AS Invoice, RmktInvoiceRequestTbl.InvoiceRequestDate AS RequestDate, RmktInvoi"
                        "ceRequestTbl.BuyerID, RmktInvoiceRequestTbl.Accepted FROM RmktInvoiceRequestTbl "
                        "WHERE (((RmktInvoiceRequestTbl.BuyerID)=4637) AND (Not (RmktInvoiceRequestTbl.Ac"
                        "cepted) Is Null)) ORDER BY RmktInvoiceRequestTbl.InvoiceRequestID DESC;"
                    ColumnWidths ="576;1008;720;0"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =90
                    LayoutCachedTop =315
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =1845
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =2
                            Left =435
                            Top =45
                            Width =2400
                            Height =345
                            FontSize =12
                            FontWeight =700
                            Name ="Label56"
                            Caption ="Buyers Order | Invoice"
                            LayoutCachedLeft =435
                            LayoutCachedTop =45
                            LayoutCachedWidth =2835
                            LayoutCachedHeight =390
                            ForeThemeColorIndex =2
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17295
                    Top =8550
                    Width =765
                    FontSize =9
                    TabIndex =19
                    ForeColor =5026082
                    Name ="ID"
                    ControlSource ="InvoiceRequestID"

                    LayoutCachedLeft =17295
                    LayoutCachedTop =8550
                    LayoutCachedWidth =18060
                    LayoutCachedHeight =8790
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4620
                    Top =120
                    Width =3690
                    Height =390
                    FontSize =12
                    FontWeight =700
                    BorderColor =-2147483615
                    Name ="InvoiceNumber"
                    ControlSource ="InvoiceNumber"
                    GroupTable =7

                    LayoutCachedLeft =4620
                    LayoutCachedTop =120
                    LayoutCachedWidth =8310
                    LayoutCachedHeight =510
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =3
                            Left =3540
                            Top =120
                            Width =1020
                            Height =390
                            FontSize =10
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =-2147483615
                            ForeColor =-2147483615
                            Name ="Label1"
                            Caption ="Invoice No:"
                            GroupTable =7
                            LayoutCachedLeft =3540
                            LayoutCachedTop =120
                            LayoutCachedWidth =4560
                            LayoutCachedHeight =510
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =7
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7110
                    Top =570
                    Width =1200
                    Height =285
                    TabIndex =3
                    BorderColor =-2147483615
                    Name ="Billdt"
                    ControlSource ="InvoiceDate"
                    Format ="Short Date"
                    DefaultValue ="Date()"
                    GroupTable =7

                    LayoutCachedLeft =7110
                    LayoutCachedTop =570
                    LayoutCachedWidth =8310
                    LayoutCachedHeight =855
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =3
                            Left =5880
                            Top =570
                            Width =1170
                            Height =285
                            FontSize =10
                            FontWeight =700
                            BorderColor =-2147483615
                            Name ="Label2"
                            Caption ="Invoice Date:"
                            GroupTable =7
                            LayoutCachedLeft =5880
                            LayoutCachedTop =570
                            LayoutCachedWidth =7050
                            LayoutCachedHeight =855
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =7
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7110
                    Top =915
                    Width =1200
                    Height =315
                    TabIndex =6
                    BorderColor =-2147483615
                    Name ="RBillDueDt"
                    ControlSource ="InvoiceDueDate"
                    Format ="Short Date"
                    StatusBarText ="date pmt due (defaults to today)"
                    ValidationRule =">Date()"
                    ValidationText ="Bill Due Dt is Required and must be greater than today"
                    DefaultValue ="Date()+5"
                    GroupTable =7
                    BottomPadding =45

                    LayoutCachedLeft =7110
                    LayoutCachedTop =915
                    LayoutCachedWidth =8310
                    LayoutCachedHeight =1230
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =3
                            Left =5880
                            Top =915
                            Width =1170
                            Height =315
                            FontSize =10
                            FontWeight =700
                            BorderColor =-2147483615
                            Name ="Label3"
                            Caption =" Due Date:"
                            GroupTable =7
                            BottomPadding =45
                            LayoutCachedLeft =5880
                            LayoutCachedTop =915
                            LayoutCachedWidth =7050
                            LayoutCachedHeight =1230
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =7
                        End
                    End
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OverlapFlags =223
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Left =3240
                    Top =1545
                    Width =7590
                    Height =315
                    TabIndex =15
                    BorderColor =-2147483615
                    ForeColor =1279872587
                    Name ="BillLink"
                    ControlSource ="InvoiceLink"
                    StatusBarText ="hyperlink to scan of billing package."

                    LayoutCachedLeft =3240
                    LayoutCachedTop =1545
                    LayoutCachedWidth =10830
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =1
                            Left =3240
                            Top =1305
                            Width =1680
                            Height =270
                            FontSize =10
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =-2147483615
                            ForeColor =-2147483615
                            Name ="Label11"
                            Caption ="Invoice Stored Link:"
                            LayoutCachedLeft =3240
                            LayoutCachedTop =1305
                            LayoutCachedWidth =4920
                            LayoutCachedHeight =1575
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4620
                    Top =570
                    Width =1200
                    Height =285
                    TabIndex =2
                    BorderColor =-2147483615
                    Name ="RequestDate"
                    ControlSource ="InvoiceRequestDate"
                    Format ="Short Date"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Remarketing  Invoice Request Date must be entered to save record."
                    GroupTable =7

                    LayoutCachedLeft =4620
                    LayoutCachedTop =570
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =855
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =3
                            Left =3540
                            Top =570
                            Width =1020
                            Height =285
                            FontSize =10
                            FontWeight =700
                            BorderColor =-2147483615
                            Name ="Label15"
                            Caption ="Requested:"
                            GroupTable =7
                            LayoutCachedLeft =3540
                            LayoutCachedTop =570
                            LayoutCachedWidth =4560
                            LayoutCachedHeight =855
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =7
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =14175
                    Top =1560
                    Width =1671
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =14
                    Name ="Print_Invoice"
                    Caption ="Print This  Invoice"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =14175
                    LayoutCachedTop =1560
                    LayoutCachedWidth =15846
                    LayoutCachedHeight =1860
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =3
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11265
                    Top =510
                    Width =990
                    Height =315
                    FontWeight =700
                    TabIndex =9
                    ForeColor =26163
                    Name ="UseTxRt"
                    ControlSource ="TaxRateToUse"
                    Format ="Percent"
                    ValidationRule =">=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =11265
                    LayoutCachedTop =510
                    LayoutCachedWidth =12255
                    LayoutCachedHeight =825
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =11220
                            Top =180
                            Width =1080
                            Height =735
                            BackColor =26163
                            ForeColor =16777215
                            Name ="Label33"
                            Caption ="TAX Rate:"
                            LayoutCachedLeft =11220
                            LayoutCachedTop =180
                            LayoutCachedWidth =12300
                            LayoutCachedHeight =915
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =105
                    Top =1995
                    Width =18555
                    Height =6480
                    TabIndex =17
                    Name ="ReMarketInvoiceSub"
                    SourceObject ="Form.ReMarketInvoiceSub"
                    LinkChildFields ="InvoiceRequestNum"
                    LinkMasterFields ="InvoiceRequestID"
                    VerticalAnchor =2

                    LayoutCachedLeft =105
                    LayoutCachedTop =1995
                    LayoutCachedWidth =18660
                    LayoutCachedHeight =8475
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15555
                    Top =8550
                    Width =765
                    FontSize =9
                    TabIndex =18
                    ForeColor =5026082
                    Name ="BuyerRef"
                    ControlSource ="BuyerID"
                    StatusBarText ="buyer"

                    LayoutCachedLeft =15555
                    LayoutCachedTop =8550
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =8790
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =3
                            Left =14775
                            Top =8550
                            Width =735
                            Height =240
                            FontSize =9
                            ForeColor =5026082
                            Name ="Label44"
                            Caption ="BuyerId:"
                            LayoutCachedLeft =14775
                            LayoutCachedTop =8550
                            LayoutCachedWidth =15510
                            LayoutCachedHeight =8790
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    Left =16365
                    Top =8550
                    Width =870
                    Height =240
                    FontSize =9
                    ForeColor =5026082
                    Name ="Label45"
                    Caption ="InvoiceId:"
                    LayoutCachedLeft =16365
                    LayoutCachedTop =8550
                    LayoutCachedWidth =17235
                    LayoutCachedHeight =8790
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =17475
                    Top =1065
                    Width =345
                    Height =270
                    TabIndex =11
                    Name ="PrtInv"
                    ControlTipText ="Prints Sales Purchase Order unless Print Invoice is checked"

                    LayoutCachedLeft =17475
                    LayoutCachedTop =1065
                    LayoutCachedWidth =17820
                    LayoutCachedHeight =1335
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =1
                            Left =16200
                            Top =1050
                            Width =1230
                            Height =315
                            FontSize =10
                            FontWeight =700
                            BorderColor =-2147483615
                            ForeColor =-2147483615
                            Name ="PINVLable"
                            Caption ="Print \"Invoice\""
                            ControlTipText ="Prints Sales Purchase Order unless Print Invoice is checked"
                            LayoutCachedLeft =16200
                            LayoutCachedTop =1050
                            LayoutCachedWidth =17430
                            LayoutCachedHeight =1365
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4620
                    Top =915
                    Width =1200
                    Height =315
                    TabIndex =5
                    BorderColor =-2147483615
                    Name ="InvoiceSentDate"
                    ControlSource ="InvoiceSentDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =7
                    BottomPadding =45

                    LayoutCachedLeft =4620
                    LayoutCachedTop =915
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =1230
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =238
                            TextAlign =3
                            Left =3540
                            Top =915
                            Width =1020
                            Height =315
                            FontSize =10
                            FontWeight =700
                            BorderColor =-2147483615
                            Name ="Label52"
                            Caption ="Sent:"
                            GroupTable =7
                            BottomPadding =45
                            LayoutCachedLeft =3540
                            LayoutCachedTop =915
                            LayoutCachedWidth =4560
                            LayoutCachedHeight =1230
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =7
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =12480
                    Top =1560
                    Width =1671
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =13
                    Name ="InvoiceReprint"
                    Caption ="Reprint Invoice"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Preview Report"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =12480
                    LayoutCachedTop =1560
                    LayoutCachedWidth =14151
                    LayoutCachedHeight =1860
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =12480
                    Top =930
                    Width =3366
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    Name ="Print_BillOfSale"
                    Caption ="Send - Bill of Sale - Delivery Package"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Preview Report"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =12480
                    LayoutCachedTop =930
                    LayoutCachedWidth =15846
                    LayoutCachedHeight =1230
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =10935
                    Top =1545
                    Width =1260
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =16
                    Name ="Makelink"
                    Caption ="Link Invoice"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10935
                    LayoutCachedTop =1545
                    LayoutCachedWidth =12195
                    LayoutCachedHeight =1845
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =4
                    BackShade =75.0
                    BorderThemeColorIndex =7
                    HoverThemeColorIndex =2
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =33
                    QuickStyleMask =-565
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9570
                    Top =120
                    Height =390
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BorderColor =-2147483615
                    Name ="InvoiceRequestID"
                    ControlSource ="InvoiceRequestID"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Invoice Number is Required"
                    DefaultValue ="\"tbd\""
                    GroupTable =7
                    RightPadding =45

                    LayoutCachedLeft =9570
                    LayoutCachedTop =120
                    LayoutCachedWidth =11010
                    LayoutCachedHeight =510
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =8370
                            Top =120
                            Width =1140
                            Height =390
                            FontSize =10
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BorderColor =-2147483615
                            ForeColor =-2147483615
                            Name ="Label76"
                            Caption ="Buyers Ord:"
                            GroupTable =7
                            LayoutCachedLeft =8370
                            LayoutCachedTop =120
                            LayoutCachedWidth =9510
                            LayoutCachedHeight =510
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =7
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9570
                    Top =570
                    Height =285
                    TabIndex =4
                    BorderColor =-2147483615
                    Name ="InvoicePaidDate"
                    ControlSource ="InvoicePaidDate"
                    Format ="Short Date"
                    GroupTable =7
                    RightPadding =45

                    LayoutCachedLeft =9570
                    LayoutCachedTop =570
                    LayoutCachedWidth =11010
                    LayoutCachedHeight =855
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =8370
                            Top =570
                            Width =1140
                            Height =285
                            FontSize =10
                            FontWeight =700
                            BorderColor =-2147483615
                            Name ="Label78"
                            Caption ="Paid:"
                            GroupTable =7
                            LayoutCachedLeft =8370
                            LayoutCachedTop =570
                            LayoutCachedWidth =9510
                            LayoutCachedHeight =855
                            RowStart =1
                            RowEnd =1
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =7
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =12480
                    Top =540
                    Width =3360
                    Height =315
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    Name ="MarkUnitsPaid"
                    Caption ="Mark Buyers Order Paid | Wire In"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Set the unit's Invoice Paid Date with the date above."
                    RightPadding =45
                    BottomPadding =45

                    LayoutCachedLeft =12480
                    LayoutCachedTop =540
                    LayoutCachedWidth =15840
                    LayoutCachedHeight =855
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =5
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =15900
                    Top =1560
                    Width =1671
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =20
                    Name ="ShowBillOfSale"
                    Caption ="Show Bill Of Sale"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =15900
                    LayoutCachedTop =1560
                    LayoutCachedWidth =17571
                    LayoutCachedHeight =1860
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =247
                    Left =3240
                    Width =14760
                    Height =1920
                    Name ="Cover1"
                    LayoutCachedLeft =3240
                    LayoutCachedWidth =18000
                    LayoutCachedHeight =1920
                    BorderThemeColorIndex =1
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =0
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "ReMarketInvoiceManager.cls"
