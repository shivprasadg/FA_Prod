Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ViewsAllowed =1
    TabularCharSet =238
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =20448
    DatasheetFontHeight =11
    ItemSuffix =233
    Left =4860
    Top =3645
    Right =25560
    Bottom =14235
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xc83239e4ef8ce440
    End
    RecordSource ="SELECT Vendors.* FROM Vendors WHERE (((Vendors.VendorObsolete)=No)) ORDER BY Ven"
        "dors.VendorObsolete DESC , Vendors.VendorSNm, Vendors.VendorName;"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
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
        Begin Tab
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =660
            BackColor =8210719
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =12540
                    Top =120
                    Width =1560
                    Height =300
                    FontSize =9
                    FontWeight =700
                    Name ="SaveMe"
                    Caption ="Save  / Update"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =12540
                    LayoutCachedTop =120
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =9
                    QuickStyleMask =-193
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10740
                    Top =120
                    Width =1560
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="AddVendor"
                    Caption ="Add Vendor"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10740
                    LayoutCachedTop =120
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =4
                    BorderColor =16777215
                    HoverThemeColorIndex =4
                    HoverTint =80.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =9
                    QuickStyleMask =-193
                End
                Begin Label
                    OverlapFlags =85
                    Width =2160
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="Label104"
                    Caption ="Vendors"
                    FontName ="Segoe UI"
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =360
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Image
                    PictureTiling = NotDefault
                    BackStyle =1
                    PictureType =2
                    Top =480
                    Width =20448
                    Height =180
                    BorderColor =6108695
                    Name ="Image142"
                    Picture ="1_DkBlue_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =480
                    LayoutCachedWidth =20448
                    LayoutCachedHeight =660
                    TabIndex =4
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =127
                    Left =18660
                    Top =180
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =2
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

                    LayoutCachedLeft =18660
                    LayoutCachedTop =180
                    LayoutCachedWidth =19215
                    LayoutCachedHeight =408
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
                    OverlapFlags =247
                    PictureType =1
                    Left =19231
                    Top =180
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =3
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =19231
                    LayoutCachedTop =180
                    LayoutCachedWidth =19981
                    LayoutCachedHeight =408
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
            Height =10035
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3855
                    Top =240
                    Width =4755
                    Height =315
                    BackColor =13434879
                    BorderColor =5676533
                    Name ="VendorName"
                    ControlSource ="VendorName"
                    ValidationRule ="Is Not Null"
                    ValidationText ="BANK NAME IS A REQUIRED FIELD"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =3855
                    LayoutCachedTop =240
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =555
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =3
                            Left =2580
                            Top =240
                            Width =1185
                            Height =315
                            FontWeight =700
                            Name ="Label0"
                            Caption ="Name:"
                            LayoutCachedLeft =2580
                            LayoutCachedTop =240
                            LayoutCachedWidth =3765
                            LayoutCachedHeight =555
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3840
                    Top =1110
                    Width =3720
                    Height =315
                    TabIndex =4
                    BorderColor =14270637
                    Name ="VendorAddress"
                    ControlSource ="VendorAddress"

                    LayoutCachedLeft =3840
                    LayoutCachedTop =1110
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =1425
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            Left =2865
                            Top =1110
                            Width =885
                            Height =315
                            FontWeight =700
                            Name ="Label2"
                            Caption ="Address:"
                            LayoutCachedLeft =2865
                            LayoutCachedTop =1110
                            LayoutCachedWidth =3750
                            LayoutCachedHeight =1425
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3825
                    Top =1545
                    Width =2700
                    Height =315
                    TabIndex =5
                    BorderColor =14270637
                    Name ="VendorCity"
                    ControlSource ="VendorCity"

                    LayoutCachedLeft =3825
                    LayoutCachedTop =1545
                    LayoutCachedWidth =6525
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            Left =3240
                            Top =1545
                            Width =495
                            Height =315
                            FontWeight =700
                            Name ="Label3"
                            Caption ="City:"
                            LayoutCachedLeft =3240
                            LayoutCachedTop =1545
                            LayoutCachedWidth =3735
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListRows =8
                    ListWidth =1440
                    Left =7380
                    Top =1545
                    Width =780
                    Height =315
                    TabIndex =6
                    BorderColor =14270637
                    Name ="VendorState"
                    ControlSource ="VendorState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lookup_States.StateAbbrev FROM lookup_States ORDER BY [StateAbbrev]; "
                    ColumnWidths ="1440"
                    StatusBarText ="lookup_states"
                    AllowValueListEdits =0

                    LayoutCachedLeft =7380
                    LayoutCachedTop =1545
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            Left =6720
                            Top =1545
                            Width =420
                            Height =315
                            FontWeight =700
                            Name ="Label4"
                            Caption ="St:"
                            LayoutCachedLeft =6720
                            LayoutCachedTop =1545
                            LayoutCachedWidth =7140
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9300
                    Top =1545
                    Height =315
                    TabIndex =7
                    BorderColor =14270637
                    Name ="VendorZip"
                    ControlSource ="VendorZip"
                    InputMask ="00000\\-9999;0;_"

                    LayoutCachedLeft =9300
                    LayoutCachedTop =1545
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            Left =8220
                            Top =1545
                            Width =945
                            Height =315
                            FontWeight =700
                            Name ="Label5"
                            Caption ="Zip Code:"
                            LayoutCachedLeft =8220
                            LayoutCachedTop =1545
                            LayoutCachedWidth =9165
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3840
                    Top =1920
                    Height =315
                    TabIndex =8
                    BorderColor =14270637
                    Name ="VendorPhone"
                    ControlSource ="VendorPhone"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =3840
                    LayoutCachedTop =1920
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =2235
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            Left =3015
                            Top =1920
                            Width =735
                            Height =315
                            FontWeight =700
                            Name ="Label7"
                            Caption ="Phone:"
                            LayoutCachedLeft =3015
                            LayoutCachedTop =1920
                            LayoutCachedWidth =3750
                            LayoutCachedHeight =2235
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6015
                    Top =1920
                    Height =315
                    TabIndex =9
                    BorderColor =14270637
                    Name ="VendorFax"
                    ControlSource ="VendorFax"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =6015
                    LayoutCachedTop =1920
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =2235
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            Left =5460
                            Top =1920
                            Width =450
                            Height =315
                            FontWeight =700
                            Name ="Label8"
                            Caption ="Fax:"
                            LayoutCachedLeft =5460
                            LayoutCachedTop =1920
                            LayoutCachedWidth =5910
                            LayoutCachedHeight =2235
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3840
                    Top =705
                    Width =4680
                    Height =315
                    TabIndex =2
                    BorderColor =14270637
                    Name ="VENDORDBA"
                    ControlSource ="VENDORDBA"

                    LayoutCachedLeft =3840
                    LayoutCachedTop =705
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =1020
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =2430
                            Top =705
                            Width =1320
                            Height =315
                            FontWeight =700
                            Name ="Label13"
                            Caption ="D/B/A:"
                            LayoutCachedLeft =2430
                            LayoutCachedTop =705
                            LayoutCachedWidth =3750
                            LayoutCachedHeight =1020
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =13590
                    Top =240
                    TabIndex =11
                    BorderColor =13434879
                    Name ="VendorObsolete"
                    ControlSource ="VendorObsolete"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =13590
                    LayoutCachedTop =240
                    LayoutCachedWidth =13850
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =12360
                            Top =240
                            Width =1155
                            Height =315
                            Name ="Label14"
                            Caption ="Obsolete:"
                            LayoutCachedLeft =12360
                            LayoutCachedTop =240
                            LayoutCachedWidth =13515
                            LayoutCachedHeight =555
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14940
                    Top =240
                    Width =1455
                    Height =315
                    TabIndex =12
                    BackColor =13434879
                    BorderColor =5676533
                    Name ="ObsoleteDate"
                    ControlSource ="VensorODt"
                    Format ="Short Date"
                    OnDblClick ="[Event Procedure]"
                    ShowDatePicker =0

                    LayoutCachedLeft =14940
                    LayoutCachedTop =240
                    LayoutCachedWidth =16395
                    LayoutCachedHeight =555
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =14175
                            Top =240
                            Width =750
                            Height =315
                            Name ="Label15"
                            Caption ="As of :"
                            LayoutCachedLeft =14175
                            LayoutCachedTop =240
                            LayoutCachedWidth =14925
                            LayoutCachedHeight =555
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =10080
                    Top =240
                    Width =2025
                    Height =315
                    TabIndex =1
                    BackColor =13434879
                    BorderColor =5676533
                    Name ="VendorSNm"
                    ControlSource ="VendorSNm"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =10080
                    LayoutCachedTop =240
                    LayoutCachedWidth =12105
                    LayoutCachedHeight =555
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =204
                            TextAlign =3
                            Left =8640
                            Top =240
                            Width =1335
                            Height =315
                            FontWeight =700
                            Name ="Label35"
                            Caption ="Short Name:"
                            LayoutCachedLeft =8640
                            LayoutCachedTop =240
                            LayoutCachedWidth =9975
                            LayoutCachedHeight =555
                        End
                    End
                End
                Begin Tab
                    OverlapFlags =85
                    Left =2520
                    Top =2805
                    Width =17880
                    Height =7230
                    TabIndex =15
                    Name ="TabCtl37"

                    LayoutCachedLeft =2520
                    LayoutCachedTop =2805
                    LayoutCachedWidth =20400
                    LayoutCachedHeight =10035
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =2655
                            Top =3285
                            Width =17610
                            Height =6615
                            Name ="Contacts"
                            ImageData = Begin
                                0x00000000
                            End
                            LayoutCachedLeft =2655
                            LayoutCachedTop =3285
                            LayoutCachedWidth =20265
                            LayoutCachedHeight =9900
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =2670
                                    Top =3330
                                    Width =17475
                                    Height =6465
                                    Name ="VendorContact_sub"
                                    SourceObject ="Form.VendorContact_sub"
                                    LinkChildFields ="VendorRef"
                                    LinkMasterFields ="VendorID"

                                    LayoutCachedLeft =2670
                                    LayoutCachedTop =3330
                                    LayoutCachedWidth =20145
                                    LayoutCachedHeight =9795
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =2655
                            Top =3285
                            Width =17610
                            Height =6620
                            Name ="Payment Info | Notes"
                            EventProcPrefix ="Payment_Info___Notes"
                            LayoutCachedLeft =2655
                            LayoutCachedTop =3285
                            LayoutCachedWidth =20265
                            LayoutCachedHeight =9905
                            Begin
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =4470
                                    Top =4080
                                    Width =1350
                                    Height =315
                                    FontSize =9
                                    BorderColor =11830108
                                    Name ="VendorPmtType"
                                    ControlSource ="VendorPmtType"
                                    RowSourceType ="Value List"
                                    RowSource ="\"WIRE\";\"ACH\";\"CHECK\";\"CASHIER CK\""
                                    ValidationRule ="Is Not Null"
                                    ValidationText ="Pmt Method is required"
                                    FontName ="Segoe UI"
                                    AllowValueListEdits =0

                                    LayoutCachedLeft =4470
                                    LayoutCachedTop =4080
                                    LayoutCachedWidth =5820
                                    LayoutCachedHeight =4395
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =3060
                                            Top =4080
                                            Width =1350
                                            Height =315
                                            FontSize =10
                                            Name ="Label19"
                                            Caption ="Pmt Method: "
                                            LayoutCachedLeft =3060
                                            LayoutCachedTop =4080
                                            LayoutCachedWidth =4410
                                            LayoutCachedHeight =4395
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =4488
                                    Top =4500
                                    Width =5892
                                    Height =315
                                    FontSize =9
                                    TabIndex =1
                                    BorderColor =11830108
                                    Name ="VENDORPAYEE"
                                    ControlSource ="VENDORPAYEE"
                                    StatusBarText ="IF DIFFERENT FROM VENDOR NAME"
                                    FontName ="Segoe UI"
                                    GroupTable =9

                                    LayoutCachedLeft =4488
                                    LayoutCachedTop =4500
                                    LayoutCachedWidth =10380
                                    LayoutCachedHeight =4815
                                    ColumnStart =1
                                    ColumnEnd =5
                                    LayoutGroup =4
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =9
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =3420
                                            Top =4500
                                            Width =990
                                            Height =315
                                            FontSize =10
                                            Name ="Label20"
                                            Caption ="Payee\015\012"
                                            GroupTable =9
                                            LayoutCachedLeft =3420
                                            LayoutCachedTop =4500
                                            LayoutCachedWidth =4410
                                            LayoutCachedHeight =4815
                                            LayoutGroup =4
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =9
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =4488
                                    Top =4884
                                    Width =5892
                                    Height =312
                                    FontSize =9
                                    TabIndex =2
                                    BorderColor =11830108
                                    Name ="VENDORBADD1"
                                    ControlSource ="VENDORBADD1"
                                    StatusBarText ="STREET ADDRESS"
                                    FontName ="Segoe UI"
                                    GroupTable =9

                                    LayoutCachedLeft =4488
                                    LayoutCachedTop =4884
                                    LayoutCachedWidth =10380
                                    LayoutCachedHeight =5196
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =5
                                    LayoutGroup =4
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =9
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =3420
                                            Top =4884
                                            Width =990
                                            Height =312
                                            FontSize =10
                                            Name ="Label24"
                                            Caption ="Address 1"
                                            GroupTable =9
                                            LayoutCachedLeft =3420
                                            LayoutCachedTop =4884
                                            LayoutCachedWidth =4410
                                            LayoutCachedHeight =5196
                                            RowStart =1
                                            RowEnd =1
                                            LayoutGroup =4
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =9
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =4488
                                    Top =5268
                                    Width =5892
                                    Height =312
                                    FontSize =9
                                    TabIndex =3
                                    BorderColor =11830108
                                    Name ="VEDNORBADD2"
                                    ControlSource ="VEDNORBADD2"
                                    StatusBarText ="SUITE ETC"
                                    FontName ="Segoe UI"
                                    GroupTable =9

                                    LayoutCachedLeft =4488
                                    LayoutCachedTop =5268
                                    LayoutCachedWidth =10380
                                    LayoutCachedHeight =5580
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =5
                                    LayoutGroup =4
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =9
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =3420
                                            Top =5268
                                            Width =990
                                            Height =312
                                            FontSize =10
                                            Name ="Label25"
                                            Caption ="Address 2"
                                            GroupTable =9
                                            LayoutCachedLeft =3420
                                            LayoutCachedTop =5268
                                            LayoutCachedWidth =4410
                                            LayoutCachedHeight =5580
                                            RowStart =2
                                            RowEnd =2
                                            LayoutGroup =4
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =9
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =4488
                                    Top =5652
                                    Width =2556
                                    Height =312
                                    FontSize =9
                                    TabIndex =4
                                    BorderColor =11830108
                                    Name ="VENDORBCITY"
                                    ControlSource ="VENDORBCITY"
                                    FontName ="Segoe UI"
                                    GroupTable =9
                                    RightPadding =0

                                    LayoutCachedLeft =4488
                                    LayoutCachedTop =5652
                                    LayoutCachedWidth =7044
                                    LayoutCachedHeight =5964
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =4
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =9
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =3420
                                            Top =5652
                                            Width =990
                                            Height =312
                                            FontSize =10
                                            Name ="Label26"
                                            Caption ="City"
                                            GroupTable =9
                                            LayoutCachedLeft =3420
                                            LayoutCachedTop =5652
                                            LayoutCachedWidth =4410
                                            LayoutCachedHeight =5964
                                            RowStart =3
                                            RowEnd =3
                                            LayoutGroup =4
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =9
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ListRows =8
                                    Left =7380
                                    Top =5652
                                    Width =876
                                    Height =312
                                    FontSize =9
                                    TabIndex =5
                                    BorderColor =11830108
                                    Name ="VENDORBST"
                                    ControlSource ="VENDORBST"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT lookup_States.StateAbbrev FROM lookup_States WHERE (((lookup_States.State"
                                        "Abbrev)=\"A\"));"
                                    FontName ="Segoe UI"
                                    GroupTable =9
                                    LeftPadding =0

                                    LayoutCachedLeft =7380
                                    LayoutCachedTop =5652
                                    LayoutCachedWidth =8256
                                    LayoutCachedHeight =5964
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =4
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =9
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =2
                                            Left =7044
                                            Top =5652
                                            Width =300
                                            Height =312
                                            FontSize =10
                                            Name ="Label27"
                                            Caption ="ST"
                                            GroupTable =9
                                            LeftPadding =0
                                            LayoutCachedLeft =7044
                                            LayoutCachedTop =5652
                                            LayoutCachedWidth =7344
                                            LayoutCachedHeight =5964
                                            RowStart =3
                                            RowEnd =3
                                            ColumnStart =2
                                            ColumnEnd =2
                                            LayoutGroup =4
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =9
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =8724
                                    Top =5652
                                    Width =1656
                                    Height =312
                                    FontSize =9
                                    TabIndex =6
                                    BorderColor =11830108
                                    Name ="VENDORBZIP"
                                    ControlSource ="VENDORBZIP"
                                    FontName ="Segoe UI"
                                    InputMask ="00000\\-9999;0;_"
                                    GroupTable =9
                                    LeftPadding =0

                                    LayoutCachedLeft =8724
                                    LayoutCachedTop =5652
                                    LayoutCachedWidth =10380
                                    LayoutCachedHeight =5964
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =5
                                    ColumnEnd =5
                                    LayoutGroup =4
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =9
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =2
                                            Left =8292
                                            Top =5652
                                            Width =396
                                            Height =312
                                            FontSize =10
                                            Name ="Label28"
                                            Caption ="ZIP"
                                            GroupTable =9
                                            LeftPadding =0
                                            LayoutCachedLeft =8292
                                            LayoutCachedTop =5652
                                            LayoutCachedWidth =8688
                                            LayoutCachedHeight =5964
                                            RowStart =3
                                            RowEnd =3
                                            ColumnStart =4
                                            ColumnEnd =4
                                            LayoutGroup =4
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =9
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =12228
                                    Top =4080
                                    Width =2976
                                    Height =315
                                    FontSize =9
                                    TabIndex =7
                                    BorderColor =11830108
                                    Name ="VENDORBANK"
                                    ControlSource ="VENDORBANK"
                                    FontName ="Segoe UI"
                                    GroupTable =5

                                    LayoutCachedLeft =12228
                                    LayoutCachedTop =4080
                                    LayoutCachedWidth =15204
                                    LayoutCachedHeight =4395
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =5
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =10980
                                            Top =4080
                                            Width =1170
                                            Height =315
                                            FontSize =10
                                            FontWeight =700
                                            Name ="Label23"
                                            Caption ="Bank  Name"
                                            GroupTable =5
                                            LayoutCachedLeft =10980
                                            LayoutCachedTop =4080
                                            LayoutCachedWidth =12150
                                            LayoutCachedHeight =4395
                                            LayoutGroup =3
                                            BackThemeColorIndex =2
                                            ForeThemeColorIndex =1
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =5
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =12228
                                    Top =4464
                                    Width =2976
                                    Height =288
                                    FontSize =9
                                    TabIndex =8
                                    BorderColor =11830108
                                    Name ="VENDORBKACCT"
                                    ControlSource ="VENDORBKACCT"
                                    StatusBarText ="BANK ACCOUNT NUMBER"
                                    FontName ="Segoe UI"
                                    GroupTable =5

                                    LayoutCachedLeft =12228
                                    LayoutCachedTop =4464
                                    LayoutCachedWidth =15204
                                    LayoutCachedHeight =4752
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =5
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =10980
                                            Top =4464
                                            Width =1170
                                            Height =288
                                            FontSize =10
                                            FontWeight =700
                                            Name ="Label32"
                                            Caption ="Bank Acct No"
                                            GroupTable =5
                                            LayoutCachedLeft =10980
                                            LayoutCachedTop =4464
                                            LayoutCachedWidth =12150
                                            LayoutCachedHeight =4752
                                            RowStart =1
                                            RowEnd =1
                                            LayoutGroup =3
                                            BackThemeColorIndex =2
                                            ForeThemeColorIndex =1
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =5
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =12228
                                    Top =4824
                                    Width =2976
                                    Height =312
                                    FontSize =9
                                    TabIndex =9
                                    BorderColor =11830108
                                    Name ="VENDORABA"
                                    ControlSource ="VENDORABA"
                                    StatusBarText ="ABA ROUTING NUMBER"
                                    FontName ="Segoe UI"
                                    GroupTable =5

                                    LayoutCachedLeft =12228
                                    LayoutCachedTop =4824
                                    LayoutCachedWidth =15204
                                    LayoutCachedHeight =5136
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =5
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =10980
                                            Top =4824
                                            Width =1170
                                            Height =312
                                            FontSize =10
                                            FontWeight =700
                                            Name ="Label22"
                                            Caption ="ABA No"
                                            GroupTable =5
                                            LayoutCachedLeft =10980
                                            LayoutCachedTop =4824
                                            LayoutCachedWidth =12150
                                            LayoutCachedHeight =5136
                                            RowStart =2
                                            RowEnd =2
                                            LayoutGroup =3
                                            BackThemeColorIndex =2
                                            ForeThemeColorIndex =1
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =5
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =12228
                                    Top =5208
                                    Width =2976
                                    Height =312
                                    FontSize =9
                                    TabIndex =10
                                    BorderColor =11830108
                                    Name ="VENDORACH"
                                    ControlSource ="VENDORACH"
                                    StatusBarText ="ACH ROUTING NUMBER"
                                    FontName ="Segoe UI"
                                    GroupTable =5

                                    LayoutCachedLeft =12228
                                    LayoutCachedTop =5208
                                    LayoutCachedWidth =15204
                                    LayoutCachedHeight =5520
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =5
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =10980
                                            Top =5208
                                            Width =1170
                                            Height =312
                                            FontSize =10
                                            FontWeight =700
                                            Name ="Label30"
                                            Caption ="ACH"
                                            GroupTable =5
                                            LayoutCachedLeft =10980
                                            LayoutCachedTop =5208
                                            LayoutCachedWidth =12150
                                            LayoutCachedHeight =5520
                                            RowStart =3
                                            RowEnd =3
                                            LayoutGroup =3
                                            BackThemeColorIndex =2
                                            ForeThemeColorIndex =1
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =5
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =12228
                                    Top =5592
                                    Width =2976
                                    Height =540
                                    FontSize =9
                                    TabIndex =11
                                    BorderColor =11830108
                                    Name ="VENDORPMTVER"
                                    ControlSource ="VENDORPMTVER"
                                    Format ="Short Date"
                                    StatusBarText ="LAST TIME VERIFIED VENDOR INSTRUCTIONS"
                                    FontName ="Segoe UI"
                                    GroupTable =5

                                    LayoutCachedLeft =12228
                                    LayoutCachedTop =5592
                                    LayoutCachedWidth =15204
                                    LayoutCachedHeight =6132
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =5
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =10980
                                            Top =5592
                                            Width =1170
                                            Height =540
                                            FontSize =10
                                            FontWeight =700
                                            Name ="Label31"
                                            Caption ="Instructions \015\012Verified"
                                            GroupTable =5
                                            LayoutCachedLeft =10980
                                            LayoutCachedTop =5592
                                            LayoutCachedWidth =12150
                                            LayoutCachedHeight =6132
                                            RowStart =4
                                            RowEnd =4
                                            LayoutGroup =3
                                            BackThemeColorIndex =2
                                            ForeThemeColorIndex =1
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =5
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =2880
                                    Top =3720
                                    Width =7620
                                    Height =2520
                                    Name ="Box195"
                                    LayoutCachedLeft =2880
                                    LayoutCachedTop =3720
                                    LayoutCachedWidth =10500
                                    LayoutCachedHeight =6240
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =10620
                                    Top =3720
                                    Width =4920
                                    Height =2520
                                    Name ="Box194"
                                    LayoutCachedLeft =10620
                                    LayoutCachedTop =3720
                                    LayoutCachedWidth =15540
                                    LayoutCachedHeight =6240
                                End
                                Begin Label
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =7260
                                    Top =4140
                                    Width =3060
                                    Height =270
                                    FontSize =9
                                    Name ="Label21"
                                    Caption ="if other than Vendor name or D/B/A"
                                    LayoutCachedLeft =7260
                                    LayoutCachedTop =4140
                                    LayoutCachedWidth =10320
                                    LayoutCachedHeight =4410
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =10680
                                    Top =3540
                                    Width =1500
                                    Height =375
                                    FontSize =12
                                    FontWeight =700
                                    ForeColor =8210719
                                    Name ="Label193"
                                    Caption ="Wire Details:"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =10680
                                    LayoutCachedTop =3540
                                    LayoutCachedWidth =12180
                                    LayoutCachedHeight =3915
                                    BackThemeColorIndex =1
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =3060
                                    Top =3540
                                    Width =2520
                                    Height =375
                                    FontSize =12
                                    FontWeight =700
                                    ForeColor =8210719
                                    Name ="Label29"
                                    Caption ="Payment Instructions:"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =3060
                                    LayoutCachedTop =3540
                                    LayoutCachedWidth =5580
                                    LayoutCachedHeight =3915
                                    BackThemeColorIndex =1
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =3900
                                    Top =6600
                                    Width =10080
                                    Height =3180
                                    TabIndex =12
                                    Name ="notes"
                                    SourceObject ="Form.VendorNotes"
                                    LinkChildFields ="vnVendorID"
                                    LinkMasterFields ="VendorID"

                                    LayoutCachedLeft =3900
                                    LayoutCachedTop =6600
                                    LayoutCachedWidth =13980
                                    LayoutCachedHeight =9780
                                End
                                Begin Rectangle
                                    OverlapFlags =223
                                    Left =2880
                                    Top =6540
                                    Width =12660
                                    Height =3300
                                    Name ="Box231"
                                    LayoutCachedLeft =2880
                                    LayoutCachedTop =6540
                                    LayoutCachedWidth =15540
                                    LayoutCachedHeight =9840
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    Left =3060
                                    Top =6300
                                    Width =780
                                    Height =375
                                    FontSize =12
                                    FontWeight =700
                                    ForeColor =8210719
                                    Name ="Label232"
                                    Caption ="Notes:"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =3060
                                    LayoutCachedTop =6300
                                    LayoutCachedWidth =3840
                                    LayoutCachedHeight =6675
                                    BackThemeColorIndex =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =2655
                            Top =3285
                            Width =17610
                            Height =6615
                            Name ="Units"
                            LayoutCachedLeft =2655
                            LayoutCachedTop =3285
                            LayoutCachedWidth =20265
                            LayoutCachedHeight =9900
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =2700
                                    Top =3330
                                    Width =17460
                                    Height =5550
                                    Name ="Vendors_UnitsSub"
                                    SourceObject ="Form.Vendors_UnitsSub"
                                    LinkChildFields ="VendorId"
                                    LinkMasterFields ="VendorID"

                                    LayoutCachedLeft =2700
                                    LayoutCachedTop =3330
                                    LayoutCachedWidth =20160
                                    LayoutCachedHeight =8880
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10080
                    Top =690
                    Height =315
                    TabIndex =3
                    BorderColor =14270637
                    Name ="VendorTaxID"
                    ControlSource ="VendorTaxID"
                    InputMask ="00\\-0000000;;_"

                    LayoutCachedLeft =10080
                    LayoutCachedTop =690
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =1005
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =8640
                            Top =690
                            Width =1335
                            Height =315
                            FontWeight =700
                            Name ="Label63"
                            Caption ="TaxID:"
                            LayoutCachedLeft =8640
                            LayoutCachedTop =690
                            LayoutCachedWidth =9975
                            LayoutCachedHeight =1005
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =12855
                    Top =1320
                    TabIndex =13
                    Name ="OEM"
                    ControlSource ="OEM"

                    LayoutCachedLeft =12855
                    LayoutCachedTop =1320
                    LayoutCachedWidth =13115
                    LayoutCachedHeight =1560
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =13085
                            Top =1290
                            Width =510
                            Height =315
                            Name ="Label68"
                            Caption ="OEM"
                            LayoutCachedLeft =13085
                            LayoutCachedTop =1290
                            LayoutCachedWidth =13595
                            LayoutCachedHeight =1605
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =12855
                    Top =1710
                    TabIndex =14
                    Name ="NotEquip"
                    ControlSource ="NotEquip"
                    StatusBarText ="i.e title and registration"

                    LayoutCachedLeft =12855
                    LayoutCachedTop =1710
                    LayoutCachedWidth =13115
                    LayoutCachedHeight =1950
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13110
                            Top =1680
                            Width =2550
                            Height =315
                            Name ="Label69"
                            Caption ="Not an Equipment Vendor"
                            LayoutCachedLeft =13110
                            LayoutCachedTop =1680
                            LayoutCachedWidth =15660
                            LayoutCachedHeight =1995
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =10095
                    Top =1110
                    TabIndex =16
                    Name ="VendorisB"
                    ControlSource ="VendorisB"
                    StatusBarText ="Vendor is a bank"

                    LayoutCachedLeft =10095
                    LayoutCachedTop =1110
                    LayoutCachedWidth =10355
                    LayoutCachedHeight =1350
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =10325
                            Top =1080
                            Width =1560
                            Height =315
                            Name ="Label80"
                            Caption ="Is a bank"
                            LayoutCachedLeft =10325
                            LayoutCachedTop =1080
                            LayoutCachedWidth =11885
                            LayoutCachedHeight =1395
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3840
                    Top =2340
                    Width =5760
                    Height =315
                    TabIndex =10
                    BorderColor =14270637
                    Name ="VendorEmail"
                    ControlSource ="VendorEmail"
                    ControlTipText ="Generic E-mail for the company. Not for a specfici contact."

                    LayoutCachedLeft =3840
                    LayoutCachedTop =2340
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =2655
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =163
                            TextAlign =3
                            Left =3000
                            Top =2340
                            Width =750
                            Height =315
                            FontWeight =700
                            Name ="VE-mailLBL"
                            Caption ="E-mail:"
                            EventProcPrefix ="VE_mailLBL"
                            ControlTipText ="Generic E-mail for the company. Not for a specfici contact."
                            LayoutCachedLeft =3000
                            LayoutCachedTop =2340
                            LayoutCachedWidth =3750
                            LayoutCachedHeight =2655
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =60
                    Top =465
                    Width =2229
                    Height =9555
                    FontSize =9
                    TabIndex =17
                    ForeColor =-2147483615
                    Name ="VendorList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorSNm, Vendors.VendorName FROM Vendors WHER"
                        "E (((Vendors.VendorObsolete)=No)) ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;2160;2880"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =465
                    LayoutCachedWidth =2289
                    LayoutCachedHeight =10020
                    BorderThemeColorIndex =1
                    BorderShade =75.0
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =60
                    Width =1575
                    Height =210
                    FontSize =8
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label58"
                    Caption ="Vendor Search"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =1635
                    LayoutCachedHeight =210
                End
                Begin TextBox
                    OverlapFlags =215
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =60
                    Top =195
                    Width =2229
                    FontSize =9
                    TabIndex =18
                    BackColor =13434879
                    BorderColor =62207
                    Name ="SearchVendor"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =195
                    LayoutCachedWidth =2289
                    LayoutCachedHeight =435
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19560
                    Top =48
                    Width =576
                    Height =192
                    FontSize =8
                    TabIndex =19
                    Name ="VendorID"
                    ControlSource ="VendorID"
                    StatusBarText ="link to Clients table"
                    DefaultValue ="3840"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19560
                    LayoutCachedTop =48
                    LayoutCachedWidth =20136
                    LayoutCachedHeight =240
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =18792
                            Top =48
                            Width =756
                            Height =192
                            FontSize =8
                            Name ="Label103"
                            Caption ="VendorId"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18792
                            LayoutCachedTop =48
                            LayoutCachedWidth =19548
                            LayoutCachedHeight =240
                            LayoutGroup =1
                            BorderThemeColorIndex =1
                            BorderShade =65.0
                            ForeThemeColorIndex =1
                            ForeShade =65.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10320
                    Top =2340
                    Width =2325
                    Height =315
                    TabIndex =20
                    Name ="ExcelVendorsUnits"
                    Caption =" This Vendor's Units"
                    OnClick ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =10320
                    LayoutCachedTop =2340
                    LayoutCachedWidth =12645
                    LayoutCachedHeight =2655
                    PictureCaptionArrangement =5
                    ColumnEnd =1
                    LayoutGroup =2
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
                    GroupTable =3
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                End
                Begin Label
                    OverlapFlags =119
                    TextAlign =3
                    Left =5760
                    Width =9600
                    Height =240
                    FontSize =8
                    Name ="Label1802"
                    Caption ="┌───────── Double-Click to Un-Lock───────────────┬───────────────────┬──────────"
                        "────────┐"
                    FontName ="Tahoma"
                    LayoutCachedLeft =5760
                    LayoutCachedWidth =15360
                    LayoutCachedHeight =240
                    ForeThemeColorIndex =4
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =14658304
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "VendorsForm.cls"
