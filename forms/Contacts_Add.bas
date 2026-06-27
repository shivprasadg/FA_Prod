Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    DatasheetFontHeight =11
    ItemSuffix =37
    Left =4380
    Top =3930
    Right =14325
    Bottom =9255
    Filter ="([ContactsQuery].[clientRef]=4133)"
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xc24f305bb1a7e340
    End
    RecordSource ="ContactsQuery"
    Caption ="Add/Edit Contacts"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    NavigationCaption ="Contacts"
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
            CanGrow = NotDefault
            Height =990
            BackColor =16249583
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6000
                    Top =660
                    Width =720
                    ColumnWidth =1260
                    FontSize =8
                    ForeColor =255
                    Name ="clientRef"
                    ControlSource ="clientRef"
                    StatusBarText ="link to Clients table"
                    DefaultValue ="=[Forms]![Deals_view]![clientRef]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =660
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =900
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =120
                    Width =6900
                    Height =390
                    FontSize =14
                    TabIndex =1
                    Name ="CompanyAndDBA"
                    ControlSource ="CompanyAndDBA"
                    FontName ="Tahoma"

                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =510
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2280
                    Top =600
                    Width =1425
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    ForeColor =255
                    Name ="CloseFormBtn"
                    Caption ="Save && Close"
                    FontName ="Tahoma"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Close"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"CloseFormBtn\" xmlns=\"http://schemas.microsoft.com/office/a"
                                "ccessservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></"
                                "Statements></UserInterfaceMacro"
                        End
                        Begin
                            Comment ="_AXL:>"
                        End
                    End

                    LayoutCachedLeft =2280
                    LayoutCachedTop =600
                    LayoutCachedWidth =3705
                    LayoutCachedHeight =960
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =600
                    Width =1830
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    ForeColor =11830108
                    Name ="AddContactBtn"
                    Caption ="Add New Contact"
                    FontName ="Tahoma"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Action ="OnError"
                            Argument ="0"
                        End
                        Begin
                            Action ="GoToRecord"
                            Argument ="-1"
                            Argument =""
                            Argument ="5"
                        End
                        Begin
                            Condition ="[MacroError]<>0"
                            Action ="MsgBox"
                            Argument ="=[MacroError].[Description]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"AddContactBtn\" xmlns=\"http://schemas.microsoft.com/office/"
                                "accessservices/2009/11/application\"><Statements><Action Name=\"OnError\"/><Acti"
                                "on Name=\"GoToRecord\"><Argument "
                        End
                        Begin
                            Comment ="_AXL:Name=\"Record\">New</Argument></Action><ConditionalBlock><If><Condition>[Ma"
                                "croError]&lt;&gt;0</Condition><Statements><Action Name=\"MessageBox\"><Argument "
                                "Name=\"Message\">=[MacroError].[Description]</Argument></Action></Statements></I"
                                "f></ConditionalBloc"
                        End
                        Begin
                            Comment ="_AXL:k></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =120
                    LayoutCachedTop =600
                    LayoutCachedWidth =1950
                    LayoutCachedHeight =960
                End
            End
        End
        Begin Section
            Height =3060
            BackColor =15523798
            Name ="Detail"
            AlternateBackColor =14270637
            Begin
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListRows =8
                    ListWidth =1440
                    Left =1199
                    Top =60
                    Width =2880
                    Height =284
                    ColumnWidth =1650
                    FontSize =8
                    Name ="contactClassification"
                    ControlSource ="contactClassification"
                    RowSourceType ="Value List"
                    RowSource ="\"Financial\";\"Operations\";\"Primary\""
                    ColumnWidths ="1440"
                    FontName ="Tahoma"
                    AllowValueListEdits =0
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =1199
                    LayoutCachedTop =60
                    LayoutCachedWidth =4079
                    LayoutCachedHeight =344
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =60
                            Width =1018
                            Height =284
                            FontSize =8
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =6050636
                            Name ="Label3"
                            Caption ="Classification"
                            FontName ="Tahoma"
                            LayoutCachedLeft =120
                            LayoutCachedTop =60
                            LayoutCachedWidth =1138
                            LayoutCachedHeight =344
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1199
                    Top =419
                    Width =840
                    Height =284
                    ColumnWidth =975
                    FontSize =8
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="contactPrefix"
                    ControlSource ="contactPrefix"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1199
                    LayoutCachedTop =419
                    LayoutCachedWidth =2039
                    LayoutCachedHeight =703
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =120
                            Top =419
                            Width =1018
                            Height =284
                            FontSize =8
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =6050636
                            Name ="Label5"
                            Caption ="Name"
                            FontName ="Tahoma"
                            LayoutCachedLeft =120
                            LayoutCachedTop =419
                            LayoutCachedWidth =1138
                            LayoutCachedHeight =703
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2040
                    Top =420
                    Width =1620
                    Height =284
                    ColumnWidth =1200
                    FontSize =8
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="contactFirst"
                    ControlSource ="contactFirst"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =2040
                    LayoutCachedTop =420
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =704
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3720
                    Top =420
                    Width =780
                    Height =284
                    ColumnWidth =570
                    FontSize =8
                    TabIndex =3
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="contactMiddle"
                    ControlSource ="contactMiddle"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =3720
                    LayoutCachedTop =420
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =704
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4560
                    Top =420
                    Height =284
                    ColumnWidth =1245
                    FontSize =8
                    TabIndex =4
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="contactLast"
                    ControlSource ="contactLast"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =420
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =704
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6060
                    Top =420
                    Width =480
                    Height =284
                    ColumnWidth =960
                    FontSize =8
                    TabIndex =5
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="contactSuffix"
                    ControlSource ="contactSuffix"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =6060
                    LayoutCachedTop =420
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =704
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1199
                    Top =1020
                    Width =2880
                    Height =284
                    ColumnWidth =840
                    FontSize =8
                    TabIndex =6
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="contactTitle"
                    ControlSource ="contactTitle"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1199
                    LayoutCachedTop =1020
                    LayoutCachedWidth =4079
                    LayoutCachedHeight =1304
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =1020
                            Width =1018
                            Height =284
                            FontSize =8
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =6050636
                            Name ="Label15"
                            Caption ="Title"
                            FontName ="Tahoma"
                            LayoutCachedLeft =120
                            LayoutCachedTop =1020
                            LayoutCachedWidth =1138
                            LayoutCachedHeight =1304
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1199
                    Top =1380
                    Width =1380
                    Height =284
                    FontSize =8
                    TabIndex =7
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="contactPrimaryPhone"
                    ControlSource ="contactPrimaryPhone"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =1199
                    LayoutCachedTop =1380
                    LayoutCachedWidth =2579
                    LayoutCachedHeight =1664
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =1380
                            Width =1154
                            Height =284
                            FontSize =8
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =6050636
                            Name ="Label17"
                            Caption ="Primary Phone"
                            FontName ="Tahoma"
                            LayoutCachedTop =1380
                            LayoutCachedWidth =1154
                            LayoutCachedHeight =1664
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2999
                    Top =1380
                    Width =720
                    Height =284
                    FontSize =8
                    TabIndex =8
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="contactPhoneExt"
                    ControlSource ="contactPhoneExt"
                    FontName ="Tahoma"

                    LayoutCachedLeft =2999
                    LayoutCachedTop =1380
                    LayoutCachedWidth =3719
                    LayoutCachedHeight =1664
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2640
                            Top =1380
                            Width =374
                            Height =284
                            FontSize =8
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =6050636
                            Name ="Label19"
                            Caption ="x"
                            FontName ="Tahoma"
                            LayoutCachedLeft =2640
                            LayoutCachedTop =1380
                            LayoutCachedWidth =3014
                            LayoutCachedHeight =1664
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1199
                    Top =1740
                    Width =1380
                    Height =284
                    FontSize =8
                    TabIndex =10
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="contactCell"
                    ControlSource ="contactCell"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =1199
                    LayoutCachedTop =1740
                    LayoutCachedWidth =2579
                    LayoutCachedHeight =2024
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =1740
                            Width =1154
                            Height =284
                            FontSize =8
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =6050636
                            Name ="Label21"
                            Caption ="Cellular"
                            FontName ="Tahoma"
                            LayoutCachedTop =1740
                            LayoutCachedWidth =1154
                            LayoutCachedHeight =2024
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1199
                    Top =2100
                    Width =3900
                    Height =284
                    FontSize =8
                    TabIndex =11
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="contactEmail"
                    ControlSource ="contactEmail"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1199
                    LayoutCachedTop =2100
                    LayoutCachedWidth =5099
                    LayoutCachedHeight =2384
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =600
                            Top =2100
                            Width =554
                            Height =284
                            FontSize =8
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =6050636
                            Name ="Label23"
                            Caption ="E-mail"
                            FontName ="Tahoma"
                            LayoutCachedLeft =600
                            LayoutCachedTop =2100
                            LayoutCachedWidth =1154
                            LayoutCachedHeight =2384
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6014
                    Top =2100
                    Width =345
                    Height =299
                    TabIndex =12
                    Name ="contactOptOut"
                    ControlSource ="contactOptOut"
                    StatusBarText ="email opt out"

                    LayoutCachedLeft =6014
                    LayoutCachedTop =2100
                    LayoutCachedWidth =6359
                    LayoutCachedHeight =2399
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5160
                            Top =2100
                            Width =749
                            Height =284
                            FontSize =8
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =6050636
                            Name ="Label25"
                            Caption ="Opt Out "
                            FontName ="Tahoma"
                            LayoutCachedLeft =5160
                            LayoutCachedTop =2100
                            LayoutCachedWidth =5909
                            LayoutCachedHeight =2384
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4859
                    Top =1380
                    Height =315
                    FontSize =8
                    TabIndex =9
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="contactPhone2"
                    ControlSource ="contactPhone2"
                    FontName ="Tahoma"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    LayoutCachedLeft =4859
                    LayoutCachedTop =1380
                    LayoutCachedWidth =6299
                    LayoutCachedHeight =1695
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4081
                            Top =1380
                            Width =779
                            Height =284
                            FontSize =8
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =6050636
                            Name ="Label27"
                            Caption ="Phone 2:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =4081
                            LayoutCachedTop =1380
                            LayoutCachedWidth =4860
                            LayoutCachedHeight =1664
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =2460
                    Width =5040
                    Height =483
                    FontSize =8
                    TabIndex =13
                    Name ="contactNote"
                    ControlSource ="contactNote"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1200
                    LayoutCachedTop =2460
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =2943
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =361
                            Top =2460
                            Width =779
                            Height =284
                            FontSize =8
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =6050636
                            Name ="Label28"
                            Caption ="Notes:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =361
                            LayoutCachedTop =2460
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =2744
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =87
                    Left =1140
                    Top =705
                    Width =840
                    Height =255
                    FontSize =8
                    ForeColor =10855845
                    Name ="Label29"
                    Caption ="salutation"
                    FontName ="Tahoma"
                    LayoutCachedLeft =1140
                    LayoutCachedTop =705
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =960
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =87
                    Left =2040
                    Top =705
                    Width =780
                    Height =255
                    FontSize =8
                    ForeColor =10855845
                    Name ="Label30"
                    Caption ="first"
                    FontName ="Tahoma"
                    LayoutCachedLeft =2040
                    LayoutCachedTop =705
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =960
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =87
                    Left =3720
                    Top =705
                    Width =780
                    Height =255
                    FontSize =8
                    ForeColor =10855845
                    Name ="Label31"
                    Caption ="middle"
                    FontName ="Tahoma"
                    LayoutCachedLeft =3720
                    LayoutCachedTop =705
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =960
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =87
                    Left =4560
                    Top =705
                    Width =780
                    Height =255
                    FontSize =8
                    ForeColor =10855845
                    Name ="Label32"
                    Caption ="last"
                    FontName ="Tahoma"
                    LayoutCachedLeft =4560
                    LayoutCachedTop =705
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =960
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =87
                    Left =6060
                    Top =705
                    Width =480
                    Height =255
                    FontSize =8
                    ForeColor =10855845
                    Name ="Label33"
                    Caption ="suffix"
                    FontName ="Tahoma"
                    LayoutCachedLeft =6060
                    LayoutCachedTop =705
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =960
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
' See "Contacts_Add.cls"
