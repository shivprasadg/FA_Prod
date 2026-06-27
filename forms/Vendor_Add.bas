Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10275
    DatasheetFontHeight =11
    ItemSuffix =15
    Left =2955
    Top =2295
    Right =13965
    Bottom =8190
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xdeb8b2d19893e340
    End
    RecordSource ="SELECT Vendors.VendorID, Vendors.VendorName, Vendors.VendorContact, Vendors.Vend"
        "orAddress, Vendors.VendorCity, Vendors.VendorState, Vendors.VendorZip, Vendors.V"
        "endorPhone, Vendors.VendorFax, Vendors.VendorPytInstructions, Vendors.VendorEmai"
        "l FROM Vendors; "
    Caption ="Add or Edit a Vendor"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    NavigationCaption ="Vendors"
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16777215
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Line
            BorderLineStyle =0
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
        Begin FormHeader
            Height =555
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8325
                    Height =315
                    FontSize =10
                    ForeColor =1643706
                    Name ="VendorID"
                    ControlSource ="VendorID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =8325
                    LayoutCachedWidth =9765
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =2910
                    Height =435
                    FontSize =14
                    FontWeight =700
                    Name ="Label13"
                    Caption ="Add/Edit Vendors"
                    FontName ="Tahoma"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =2970
                    LayoutCachedHeight =495
                End
            End
        End
        Begin Section
            Height =3480
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =120
                    Width =5025
                    Height =315
                    FontSize =10
                    Name ="VendorName"
                    ControlSource ="VendorName"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =120
                    LayoutCachedWidth =6705
                    LayoutCachedHeight =435
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =150
                            Top =120
                            Width =1440
                            Height =315
                            FontSize =10
                            Name ="Label1"
                            Caption ="Vendor Name:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =150
                            LayoutCachedTop =120
                            LayoutCachedWidth =1590
                            LayoutCachedHeight =435
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =515
                    Width =5025
                    Height =315
                    FontSize =10
                    TabIndex =1
                    Name ="VendorContact"
                    ControlSource ="VendorContact"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =515
                    LayoutCachedWidth =6705
                    LayoutCachedHeight =830
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =150
                            Top =515
                            Width =1440
                            Height =315
                            FontSize =10
                            Name ="Label2"
                            Caption ="Contact Name:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =150
                            LayoutCachedTop =515
                            LayoutCachedWidth =1590
                            LayoutCachedHeight =830
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =910
                    Width =5025
                    Height =315
                    FontSize =10
                    TabIndex =2
                    Name ="VendorAddress"
                    ControlSource ="VendorAddress"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =910
                    LayoutCachedWidth =6705
                    LayoutCachedHeight =1225
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =150
                            Top =910
                            Width =1440
                            Height =315
                            FontSize =10
                            Name ="Label3"
                            Caption ="Address:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =150
                            LayoutCachedTop =910
                            LayoutCachedWidth =1590
                            LayoutCachedHeight =1225
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1305
                    Width =1710
                    Height =315
                    FontSize =10
                    TabIndex =3
                    Name ="VendorCity"
                    ControlSource ="VendorCity"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =1305
                    LayoutCachedWidth =3390
                    LayoutCachedHeight =1620
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =150
                            Top =1305
                            Width =1440
                            Height =315
                            FontSize =10
                            Name ="Label4"
                            Caption ="City:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =150
                            LayoutCachedTop =1305
                            LayoutCachedWidth =1590
                            LayoutCachedHeight =1620
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ListRows =8
                    ListWidth =1440
                    Left =4080
                    Top =1320
                    Width =465
                    Height =315
                    FontSize =10
                    TabIndex =4
                    Name ="VendorState"
                    ControlSource ="VendorState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lookup_States.StateAbbrev FROM lookup_States ORDER BY [StateAbbrev]; "
                    ColumnWidths ="1440"
                    FontName ="Tahoma"
                    AllowValueListEdits =0

                    LayoutCachedLeft =4080
                    LayoutCachedTop =1320
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =1635
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3465
                            Top =1320
                            Width =615
                            Height =315
                            FontSize =10
                            Name ="Label5"
                            Caption ="State:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =3465
                            LayoutCachedTop =1320
                            LayoutCachedWidth =4080
                            LayoutCachedHeight =1635
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5610
                    Top =1275
                    Width =1125
                    Height =315
                    FontSize =10
                    TabIndex =5
                    Name ="VendorZip"
                    ControlSource ="VendorZip"
                    FontName ="Tahoma"

                    LayoutCachedLeft =5610
                    LayoutCachedTop =1275
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =1590
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4620
                            Top =1275
                            Width =945
                            Height =315
                            FontSize =10
                            Name ="Label6"
                            Caption ="Zip Code:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =4620
                            LayoutCachedTop =1275
                            LayoutCachedWidth =5565
                            LayoutCachedHeight =1590
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7845
                    Top =105
                    Width =2400
                    Height =315
                    FontSize =10
                    TabIndex =6
                    Name ="VendorPhone"
                    ControlSource ="VendorPhone"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =7845
                    LayoutCachedTop =105
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7005
                            Top =105
                            Width =735
                            Height =315
                            FontSize =10
                            Name ="Label7"
                            Caption ="Phone:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =7005
                            LayoutCachedTop =105
                            LayoutCachedWidth =7740
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7845
                    Top =500
                    Width =2400
                    Height =315
                    FontSize =10
                    TabIndex =8
                    Name ="VendorFax"
                    ControlSource ="VendorFax"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =7845
                    LayoutCachedTop =500
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =815
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7005
                            Top =500
                            Width =735
                            Height =315
                            FontSize =10
                            Name ="Label8"
                            Caption ="Fax:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =7005
                            LayoutCachedTop =500
                            LayoutCachedWidth =7740
                            LayoutCachedHeight =815
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1710
                    Top =1725
                    Width =8475
                    Height =1008
                    FontSize =10
                    TabIndex =9
                    Name ="VendorPytInstructions"
                    ControlSource ="VendorPytInstructions"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1710
                    LayoutCachedTop =1725
                    LayoutCachedWidth =10185
                    LayoutCachedHeight =2733
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =165
                            Top =1725
                            Width =1440
                            Height =675
                            FontSize =10
                            Name ="Label9"
                            Caption ="Payment Instructions:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =165
                            LayoutCachedTop =1725
                            LayoutCachedWidth =1605
                            LayoutCachedHeight =2400
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7845
                    Top =915
                    Width =2400
                    Height =315
                    FontSize =10
                    TabIndex =7
                    Name ="VendorEmail"
                    ControlSource ="VendorEmail"
                    FontName ="Tahoma"

                    LayoutCachedLeft =7845
                    LayoutCachedTop =915
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7005
                            Top =915
                            Width =735
                            Height =315
                            FontSize =10
                            Name ="Label10"
                            Caption ="E-mail:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =7005
                            LayoutCachedTop =915
                            LayoutCachedWidth =7740
                            LayoutCachedHeight =1230
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =85
                    SpecialEffect =2
                    Left =6840
                    Top =120
                    Width =15
                    Height =1500
                    Name ="Line11"
                    LayoutCachedLeft =6840
                    LayoutCachedTop =120
                    LayoutCachedWidth =6855
                    LayoutCachedHeight =1620
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =3000
                    Width =1320
                    Height =330
                    FontSize =9
                    TabIndex =10
                    ForeColor =2366701
                    Name ="CloseFormBtn"
                    Caption ="Save && Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =120
                    LayoutCachedTop =3000
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =3330
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =1620
                    Top =3000
                    Width =1395
                    Height =330
                    FontSize =9
                    TabIndex =11
                    ForeColor =5026082
                    Name ="EditVendorBtn"
                    Caption ="Edit a Vendor"
                    FontName ="Tahoma"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="ShowAllRecords"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"EditVendorBtn\" xmlns=\"http://schemas.microsoft.com/office/"
                                "accessservices/2009/11/application\"><Statements><Action Name=\"ShowAllRecords\""
                                "/></Statements></UserInterfaceM"
                        End
                        Begin
                            Comment ="_AXL:acro>"
                        End
                    End

                    LayoutCachedLeft =1620
                    LayoutCachedTop =3000
                    LayoutCachedWidth =3015
                    LayoutCachedHeight =3330
                End
            End
        End
        Begin FormFooter
            Height =405
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "Vendor_Add.cls"
