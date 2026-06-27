Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =24594
    DatasheetFontHeight =10
    ItemSuffix =738
    Left =7545
    Top =3060
    Right =32400
    Bottom =17340
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x36556d936db6e540
    End
    RecordSource ="Select * From Buyer"
    Caption ="Remarketing Buyers"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    NavigationCaption ="Company Navigation"
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin BoundObjectFrame
            SizeMode =3
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =-1800
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
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
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1140
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =17190
                    Top =450
                    Width =2010
                    Height =288
                    FontSize =9
                    TabIndex =2
                    Name ="AddNewBuyer"
                    Caption ="Add Buyer"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =17190
                    LayoutCachedTop =450
                    LayoutCachedWidth =19200
                    LayoutCachedHeight =738
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =10319446
                    BorderThemeColorIndex =1
                    HoverColor =12428174
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15690
                    Top =450
                    Height =288
                    FontSize =9
                    TabIndex =1
                    Name ="SaveMe"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =15690
                    LayoutCachedTop =450
                    LayoutCachedWidth =17130
                    LayoutCachedHeight =738
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =10319446
                    BorderThemeColorIndex =1
                    HoverColor =12428174
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    PressedForeThemeColorIndex =1
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =840
                    Top =600
                    Width =9825
                    Height =360
                    ColumnOrder =0
                    FontSize =14
                    FontWeight =700
                    BackColor =-2147483609
                    Name ="Text406"
                    ControlSource ="=\"Buyer:  \" & [BuyerName]"
                    FontName ="Calibri"

                    LayoutCachedLeft =840
                    LayoutCachedTop =600
                    LayoutCachedWidth =10665
                    LayoutCachedHeight =960
                    ForeThemeColorIndex =2
                End
                Begin Image
                    BackStyle =1
                    SizeMode =1
                    PictureAlignment =0
                    PictureType =2
                    Width =691
                    Height =780
                    Name ="Image1161"
                    Picture ="ATLAASGlobeLogo"
                    GroupTable =25
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =691
                    LayoutCachedHeight =780
                    TabIndex =5
                    RowEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =25
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =780
                    Width =7620
                    Height =780
                    FontSize =24
                    FontWeight =700
                    Name ="Label465"
                    Caption ="Remarketing Buyers && Contacts"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =780
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =780
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =960
                    Width =24594
                    Height =180
                    Name ="Image142"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =960
                    LayoutCachedWidth =24594
                    LayoutCachedHeight =1140
                    TabIndex =6
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =22860
                    Top =480
                    Width =555
                    Height =228
                    TabIndex =3
                    Name ="GoHome"
                    Caption ="Home"
                    FontName ="Calibri"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1
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

                    LayoutCachedLeft =22860
                    LayoutCachedTop =480
                    LayoutCachedWidth =23415
                    LayoutCachedHeight =708
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
                    PictureType =1
                    Left =23427
                    Top =480
                    Width =555
                    Height =228
                    TabIndex =4
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1

                    LayoutCachedLeft =23427
                    LayoutCachedTop =480
                    LayoutCachedWidth =23982
                    LayoutCachedHeight =708
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =13155
            Name ="Detail"
            Begin
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =11805
                    Top =720
                    Width =300
                    Height =270
                    TabIndex =22
                    BorderColor =6108695
                    Name ="Wholeseller"
                    ControlSource ="Wholeseller"
                    StatusBarText ="is remarketing Customer a wholesaler?"
                    Tag ="EditYN"

                    LayoutCachedLeft =11805
                    LayoutCachedTop =720
                    LayoutCachedWidth =12105
                    LayoutCachedHeight =990
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =247
                            TextAlign =1
                            Left =10455
                            Top =690
                            Width =1635
                            Height =240
                            FontSize =10
                            BorderColor =6108695
                            Name ="WHLabel"
                            Caption ="  Wholeseller?"
                            FontName ="Calibri"
                            LayoutCachedLeft =10455
                            LayoutCachedTop =690
                            LayoutCachedWidth =12090
                            LayoutCachedHeight =930
                            RowStart =1
                            RowEnd =1
                            ColumnStart =6
                            ColumnEnd =7
                            BackThemeColorIndex =1
                            BackShade =85.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =100
                    ListWidth =2304
                    Left =19620
                    Top =360
                    Width =2139
                    Height =285
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =18
                    Name ="SalesRep"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Emps.EmpInt, Emps.EmpNm FROM Emps WHERE (((Emps.Rmkt)=Yes) AND ((Emps.OBS"
                        "OLETE)=No)) OR (((Emps.EmpInt)=\"hs\")) ORDER BY Emps.EmpInt;"
                    ColumnWidths ="432;1872"
                    FontName ="Calibri"
                    InheritValueList =1

                    LayoutCachedLeft =19620
                    LayoutCachedTop =360
                    LayoutCachedWidth =21759
                    LayoutCachedHeight =645
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =4
                End
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =2460
                    Top =2160
                    Width =22110
                    Height =10995
                    FontSize =10
                    TabIndex =21
                    Name ="TabControls"
                    FontName ="Calibri"
                    VerticalAnchor =2

                    LayoutCachedLeft =2460
                    LayoutCachedTop =2160
                    LayoutCachedWidth =24570
                    LayoutCachedHeight =13155
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =2595
                            Top =2595
                            Width =21840
                            Height =10425
                            BorderColor =0
                            Name ="Contacts"
                            Caption ="Contacts"
                            GridlineColor =0
                            LayoutCachedLeft =2595
                            LayoutCachedTop =2595
                            LayoutCachedWidth =24435
                            LayoutCachedHeight =13020
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =2655
                                    Top =2640
                                    Width =16305
                                    Height =9645
                                    Name ="ContactDetails"
                                    SourceObject ="Form.RemarketingContact_sub"
                                    LinkChildFields ="ClientId"
                                    LinkMasterFields ="BuyerID"
                                    FilterOnEmptyMaster =0

                                    LayoutCachedLeft =2655
                                    LayoutCachedTop =2640
                                    LayoutCachedWidth =18960
                                    LayoutCachedHeight =12285
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =2595
                            Top =2595
                            Width =21840
                            Height =10425
                            BorderColor =0
                            Name ="Locations"
                            GridlineColor =0
                            LayoutCachedLeft =2595
                            LayoutCachedTop =2595
                            LayoutCachedWidth =24435
                            LayoutCachedHeight =13020
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =2655
                                    Top =2640
                                    Width =21660
                                    Height =9000
                                    Name ="RemarketingClientLocations"
                                    SourceObject ="Form.RemarketingClientLocations"
                                    LinkChildFields ="ClientREF"
                                    LinkMasterFields ="BuyerID"

                                    LayoutCachedLeft =2655
                                    LayoutCachedTop =2640
                                    LayoutCachedWidth =24315
                                    LayoutCachedHeight =11640
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =2595
                            Top =2595
                            Width =21840
                            Height =10425
                            BorderColor =0
                            Name ="Request Invoice"
                            EventProcPrefix ="Request_Invoice"
                            Caption ="Invoice Request | Buyers Order"
                            GridlineColor =0
                            LayoutCachedLeft =2595
                            LayoutCachedTop =2595
                            LayoutCachedWidth =24435
                            LayoutCachedHeight =13020
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =2610
                                    Top =2640
                                    Width =21585
                                    Height =10380
                                    Name ="RemarketingInvoiceRequest_Sub"
                                    SourceObject ="Form.RemarketingInvoiceRequest_Sub"
                                    LinkChildFields ="BuyerID"
                                    LinkMasterFields ="BuyerID"

                                    LayoutCachedLeft =2610
                                    LayoutCachedTop =2640
                                    LayoutCachedWidth =24195
                                    LayoutCachedHeight =13020
                                End
                            End
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =2460
                    Top =30
                    Width =2550
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label120"
                    Caption ="Company Info"
                    FontName ="Calibri"
                    LayoutCachedLeft =2460
                    LayoutCachedTop =30
                    LayoutCachedWidth =5010
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3540
                    Top =270
                    Width =3435
                    ColumnOrder =22
                    FontSize =9
                    TabIndex =1
                    BackColor =13434879
                    BorderColor =967423
                    Name ="BuyerName"
                    ControlSource ="BuyerName"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =3540
                    LayoutCachedTop =270
                    LayoutCachedWidth =6975
                    LayoutCachedHeight =510
                    ColumnStart =1
                    ColumnEnd =3
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =2445
                            Top =270
                            Width =1035
                            Height =240
                            BorderColor =967423
                            Name ="Label138"
                            Caption ="Company:"
                            FontName ="Calibri"
                            LayoutCachedLeft =2445
                            LayoutCachedTop =270
                            LayoutCachedWidth =3480
                            LayoutCachedHeight =510
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3540
                    Top =570
                    Width =3435
                    ColumnOrder =20
                    FontSize =9
                    TabIndex =4
                    Name ="BuyerDBA"
                    ControlSource ="BuyerDBA"
                    FontName ="Calibri"

                    LayoutCachedLeft =3540
                    LayoutCachedTop =570
                    LayoutCachedWidth =6975
                    LayoutCachedHeight =810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =3
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2445
                            Top =570
                            Width =1035
                            Height =240
                            Name ="Label139"
                            Caption ="DBA:"
                            FontName ="Calibri"
                            LayoutCachedLeft =2445
                            LayoutCachedTop =570
                            LayoutCachedWidth =3480
                            LayoutCachedHeight =810
                            RowStart =1
                            RowEnd =1
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3540
                    Top =870
                    Width =3435
                    ColumnOrder =19
                    FontSize =9
                    TabIndex =6
                    Name ="bAddress1"
                    ControlSource ="bAddress1"
                    FontName ="Calibri"

                    LayoutCachedLeft =3540
                    LayoutCachedTop =870
                    LayoutCachedWidth =6975
                    LayoutCachedHeight =1110
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =3
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2445
                            Top =870
                            Width =1035
                            Height =240
                            Name ="Label140"
                            Caption ="Address1:"
                            FontName ="Calibri"
                            LayoutCachedLeft =2445
                            LayoutCachedTop =870
                            LayoutCachedWidth =3480
                            LayoutCachedHeight =1110
                            RowStart =2
                            RowEnd =2
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3540
                    Top =1170
                    Width =3435
                    ColumnOrder =18
                    FontSize =9
                    TabIndex =8
                    Name ="bAddress2"
                    ControlSource ="bAddress2"
                    FontName ="Calibri"

                    LayoutCachedLeft =3540
                    LayoutCachedTop =1170
                    LayoutCachedWidth =6975
                    LayoutCachedHeight =1410
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =3
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2445
                            Top =1170
                            Width =1035
                            Height =240
                            Name ="Label141"
                            Caption ="Address2:"
                            FontName ="Calibri"
                            LayoutCachedLeft =2445
                            LayoutCachedTop =1170
                            LayoutCachedWidth =3480
                            LayoutCachedHeight =1410
                            RowStart =3
                            RowEnd =3
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3540
                    Top =1470
                    Width =1680
                    ColumnOrder =17
                    FontSize =9
                    TabIndex =10
                    Name ="bCity"
                    ControlSource ="bCity"
                    FontName ="Calibri"

                    LayoutCachedLeft =3540
                    LayoutCachedTop =1470
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =1710
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2445
                            Top =1470
                            Width =1035
                            Height =240
                            Name ="Label142"
                            Caption ="City State Zip:"
                            FontName ="Calibri"
                            LayoutCachedLeft =2445
                            LayoutCachedTop =1470
                            LayoutCachedWidth =3480
                            LayoutCachedHeight =1710
                            RowStart =4
                            RowEnd =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5280
                    Top =1470
                    Width =720
                    ColumnOrder =16
                    FontSize =9
                    TabIndex =11
                    Name ="bState"
                    ControlSource ="bState"
                    RowSourceType ="Table/Query"
                    RowSource ="lookup_States"
                    StatusBarText ="lookup states table"
                    FontName ="Calibri"

                    LayoutCachedLeft =5280
                    LayoutCachedTop =1470
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =1710
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6060
                    Top =1470
                    Width =915
                    ColumnOrder =15
                    FontSize =9
                    TabIndex =12
                    Name ="bZip"
                    ControlSource ="bZip"
                    FontName ="Calibri"

                    LayoutCachedLeft =6060
                    LayoutCachedTop =1470
                    LayoutCachedWidth =6975
                    LayoutCachedHeight =1710
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8295
                    Top =570
                    Width =1725
                    ColumnOrder =13
                    FontSize =9
                    TabIndex =5
                    Name ="bPhone"
                    ControlSource ="bPhone"
                    FontName ="Calibri"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =8295
                    LayoutCachedTop =570
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =810
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7035
                            Top =570
                            Width =1200
                            Height =240
                            Name ="Label27"
                            Caption ="Main Phone:"
                            FontName ="Calibri"
                            LayoutCachedLeft =7035
                            LayoutCachedTop =570
                            LayoutCachedWidth =8235
                            LayoutCachedHeight =810
                            RowStart =1
                            RowEnd =1
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8295
                    Top =870
                    Width =1725
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =7
                    Name ="bFax"
                    ControlSource ="bFax"
                    FontName ="Calibri"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =8295
                    LayoutCachedTop =870
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =1110
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7035
                            Top =870
                            Width =1200
                            Height =240
                            Name ="Label31"
                            Caption ="Fax:"
                            FontName ="Calibri"
                            LayoutCachedLeft =7035
                            LayoutCachedTop =870
                            LayoutCachedWidth =8235
                            LayoutCachedHeight =1110
                            RowStart =2
                            RowEnd =2
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3540
                    Top =1770
                    Width =3435
                    Height =255
                    ColumnOrder =11
                    FontSize =9
                    TabIndex =14
                    BorderColor =11830108
                    Name ="bWebsite"
                    ControlSource ="bWebsite"
                    FontName ="Calibri"

                    LayoutCachedLeft =3540
                    LayoutCachedTop =1770
                    LayoutCachedWidth =6975
                    LayoutCachedHeight =2025
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =3
                    BackThemeColorIndex =4
                    BackTint =20.0
                    ForeThemeColorIndex =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =3
                            Left =2445
                            Top =1770
                            Width =1035
                            Height =255
                            Name ="Label35"
                            Caption ="Website:"
                            FontName ="Calibri"
                            LayoutCachedLeft =2445
                            LayoutCachedTop =1770
                            LayoutCachedWidth =3480
                            LayoutCachedHeight =2025
                            RowStart =5
                            RowEnd =5
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8295
                    Top =1170
                    Width =1725
                    ColumnOrder =14
                    FontSize =9
                    TabIndex =9
                    Name ="LeaseSTofCorp"
                    ControlSource ="bStateOfCorp"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lookup_States.StateAbbrev FROM lookup_States ORDER BY lookup_States.State"
                        "Abbrev;"
                    FontName ="Calibri"

                    LayoutCachedLeft =8295
                    LayoutCachedTop =1170
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =1410
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    ThemeFontIndex =1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7035
                            Top =1170
                            Width =1200
                            Height =240
                            Name ="Label249"
                            Caption ="Incorporated in:"
                            FontName ="Calibri"
                            LayoutCachedLeft =7035
                            LayoutCachedTop =1170
                            LayoutCachedWidth =8235
                            LayoutCachedHeight =1410
                            RowStart =3
                            RowEnd =3
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8295
                    Top =1470
                    Width =1725
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =13
                    Name ="bTaxID"
                    ControlSource ="bTaxID"
                    FontName ="Calibri"

                    LayoutCachedLeft =8295
                    LayoutCachedTop =1470
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =1710
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7035
                            Top =1470
                            Width =1200
                            Height =240
                            Name ="Label248"
                            Caption ="Tax ID:"
                            FontName ="Calibri"
                            LayoutCachedLeft =7035
                            LayoutCachedTop =1470
                            LayoutCachedWidth =8235
                            LayoutCachedHeight =1710
                            RowStart =4
                            RowEnd =4
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8295
                    Top =270
                    Width =1725
                    ColumnWidth =1425
                    ColumnOrder =21
                    FontSize =9
                    TabIndex =2
                    BackColor =13434879
                    BorderColor =967423
                    Name ="BuyerShortName"
                    ControlSource ="BuyerShortName"
                    StatusBarText ="Short Name for Reporting"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Sort Nm for Reports is required"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =8295
                    LayoutCachedTop =270
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =510
                    ColumnStart =5
                    ColumnEnd =5
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =3
                            Left =7035
                            Top =270
                            Width =1200
                            Height =240
                            BorderColor =967423
                            Name ="Label358"
                            Caption ="Short Name:"
                            FontName ="Calibri"
                            LayoutCachedLeft =7035
                            LayoutCachedTop =270
                            LayoutCachedWidth =8235
                            LayoutCachedHeight =510
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10455
                    Top =360
                    Width =1635
                    ColumnOrder =6
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    BorderColor =11830108
                    Name ="clientType"
                    ControlSource ="bType"
                    ValidationRule ="Is Not Null"
                    ValidationText ="CLIENT TYPE REQUIRED"
                    DefaultValue ="\"Remarketing\""
                    FontName ="Calibri"

                    LayoutCachedLeft =10455
                    LayoutCachedTop =360
                    LayoutCachedWidth =12090
                    LayoutCachedHeight =600
                    ColumnStart =7
                    ColumnEnd =8
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8295
                    Top =1770
                    Width =1725
                    Height =255
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =15
                    Name ="ResaleExpire"
                    ControlSource ="ResaleExpire"
                    Format ="Short Date"
                    StatusBarText ="Remarketing - Dealer or Resale Expriation date."
                    FontName ="Calibri"

                    LayoutCachedLeft =8295
                    LayoutCachedTop =1770
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =2025
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =5
                    BorderThemeColorIndex =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =7035
                            Top =1770
                            Width =1200
                            Height =255
                            Name ="Label440"
                            Caption ="Dealer Lic Expires:"
                            FontName ="Calibri"
                            LayoutCachedLeft =7035
                            LayoutCachedTop =1770
                            LayoutCachedWidth =8235
                            LayoutCachedHeight =2025
                            RowStart =5
                            RowEnd =5
                            ColumnStart =4
                            ColumnEnd =4
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =238
                    Left =6255
                    Top =90
                    Width =3705
                    Height =255
                    FontWeight =700
                    ForeColor =-2147483611
                    Name ="Label485"
                    Caption ="┌─────────── double click to unlock ──────────┐"
                    FontName ="Calibri"
                    LayoutCachedLeft =6255
                    LayoutCachedTop =90
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =345
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =30
                    Top =555
                    Width =2355
                    Height =12465
                    ColumnOrder =0
                    FontSize =8
                    TabIndex =17
                    Name ="BuyerList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Buyer.BuyerID, Buyer.BuyerShortName AS Buyer FROM Buyer ORDER BY Buyer.Bu"
                        "yerShortName;"
                    ColumnWidths ="0;2016"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =30
                    LayoutCachedTop =555
                    LayoutCachedWidth =2385
                    LayoutCachedHeight =13020
                    BorderThemeColorIndex =2
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =45
                    Top =210
                    Width =2340
                    Height =270
                    ColumnOrder =23
                    FontSize =9
                    BackColor =13952764
                    BorderColor =967423
                    Name ="SearchBuyer"
                    FontName ="Calibri"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =45
                    LayoutCachedTop =210
                    LayoutCachedWidth =2385
                    LayoutCachedHeight =480
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =45
                            Width =2340
                            Height =240
                            FontWeight =700
                            BorderColor =967423
                            Name ="Label489"
                            Caption ="Type Buyer Search HERE:"
                            FontName ="Calibri"
                            LayoutCachedLeft =45
                            LayoutCachedWidth =2385
                            LayoutCachedHeight =240
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12435
                    Top =495
                    Width =3720
                    Height =1365
                    TabIndex =16
                    ForeColor =-2147483617
                    Name ="AddressPreview"
                    ControlSource ="=[BuyerName] & Chr(13) & Chr(10) & IIf(IsNull([BuyerDBA]),Null,\"dba:\" & [Buyer"
                        "DBA] & Chr(13) & Chr(10)) & IIf(IsNull([bAddress1]),Null,[bAddress1] & Chr(13) &"
                        " Chr(10)) & IIf(IsNull([bAddress2]),Null,[bAddress2] & Chr(13) & Chr(10)) & [bCi"
                        "ty] & \" \" & [bState] & \" \" & [bZip]"

                    LayoutCachedLeft =12435
                    LayoutCachedTop =495
                    LayoutCachedWidth =16155
                    LayoutCachedHeight =1860
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =255
                    Left =12360
                    Top =300
                    Width =3900
                    Height =1650
                    Name ="Box656"
                    LayoutCachedLeft =12360
                    LayoutCachedTop =300
                    LayoutCachedWidth =16260
                    LayoutCachedHeight =1950
                    BorderThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =12420
                    Top =180
                    Width =1455
                    Height =240
                    FontWeight =700
                    Name ="Label657"
                    Caption ="Address Preview:"
                    FontName ="Calibri"
                    LayoutCachedLeft =12420
                    LayoutCachedTop =180
                    LayoutCachedWidth =13875
                    LayoutCachedHeight =420
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =18390
                    Top =300
                    Width =990
                    ColumnOrder =5
                    TabIndex =19
                    ForeColor =5026082
                    Name ="BuyerID"
                    ControlSource ="BuyerID"
                    FontName ="Calibri"

                    LayoutCachedLeft =18390
                    LayoutCachedTop =300
                    LayoutCachedWidth =19380
                    LayoutCachedHeight =540
                    BorderThemeColorIndex =6
                    BorderTint =60.0
                    ThemeFontIndex =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =17685
                            Top =315
                            Width =675
                            Height =240
                            ForeColor =5026082
                            Name ="BuyerIDLbl"
                            Caption ="BuyerID:"
                            FontName ="Calibri"
                            LayoutCachedLeft =17685
                            LayoutCachedTop =315
                            LayoutCachedWidth =18360
                            LayoutCachedHeight =555
                            ThemeFontIndex =1
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =18390
                    Top =510
                    Width =990
                    TabIndex =20
                    ForeColor =5026082
                    Name ="ContactID"
                    FontName ="Calibri"

                    LayoutCachedLeft =18390
                    LayoutCachedTop =510
                    LayoutCachedWidth =19380
                    LayoutCachedHeight =750
                    BorderThemeColorIndex =6
                    BorderTint =60.0
                    ThemeFontIndex =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =17475
                            Top =525
                            Width =885
                            Height =240
                            ForeColor =5026082
                            Name ="Label667"
                            Caption ="ContactID:"
                            FontName ="Calibri"
                            LayoutCachedLeft =17475
                            LayoutCachedTop =525
                            LayoutCachedWidth =18360
                            LayoutCachedHeight =765
                            ThemeFontIndex =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10455
                    Top =1350
                    Width =1635
                    FontSize =9
                    FontWeight =700
                    TabIndex =23
                    BorderColor =967423
                    Name ="bStatus"
                    ControlSource ="bStatus"
                    FontName ="Calibri"

                    LayoutCachedLeft =10455
                    LayoutCachedTop =1350
                    LayoutCachedWidth =12090
                    LayoutCachedHeight =1590
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =9
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =10455
                    Top =1050
                    Width =1635
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label720"
                    Caption ="Buyer Status"
                    FontName ="Calibri"
                    LayoutCachedLeft =10455
                    LayoutCachedTop =1050
                    LayoutCachedWidth =12090
                    LayoutCachedHeight =1290
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "BuyersForm.cls"
