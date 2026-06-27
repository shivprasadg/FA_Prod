Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13800
    DatasheetFontHeight =10
    ItemSuffix =66
    Left =4680
    Top =6270
    Right =18465
    Bottom =12435
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xaa0184acb1a7e340
    End
    RecordSource ="Documentation"
    Caption ="Documentation"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    NavigationCaption ="Documentation"
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
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
            Height =435
            BackColor =15523798
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =1875
                    Height =330
                    FontSize =9
                    FontWeight =700
                    Name ="AddContactBtn"
                    Caption ="Add/Edit Contacts"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1935
                    LayoutCachedHeight =390
                End
            End
        End
        Begin Section
            Height =1890
            BackColor =15523798
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =2310
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Title"
                    Caption ="Lessee Information"
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =4020
                    Top =120
                    Width =5415
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label30"
                    Caption ="-----------------  Incumbent Signers --------------"
                    LayoutCachedLeft =4020
                    LayoutCachedTop =120
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =405
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =4080
                    Top =450
                    Width =2820
                    Height =240
                    ForeColor =9605778
                    Name ="Label31"
                    Caption ="Signer"
                    LayoutCachedLeft =4080
                    LayoutCachedTop =450
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =690
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =7035
                    Top =450
                    Width =2400
                    Height =240
                    ForeColor =9605778
                    Name ="Label32"
                    Caption ="Title"
                    LayoutCachedLeft =7035
                    LayoutCachedTop =450
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =690
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =1050
                    Width =2400
                    TabIndex =4
                    BorderColor =14270637
                    Name ="docRALSigTitle"
                    ControlSource ="docRALSigTitle"

                    LayoutCachedLeft =7020
                    LayoutCachedTop =1050
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =1290
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =750
                    Width =2400
                    TabIndex =1
                    BorderColor =14270637
                    Name ="docScheduleSigTitle"
                    ControlSource ="docSchedSigTitle"

                    LayoutCachedLeft =7020
                    LayoutCachedTop =750
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =990
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =1080
                    Top =450
                    Width =1125
                    Height =240
                    ForeColor =9605778
                    Name ="Label45"
                    Caption ="Contact"
                    LayoutCachedLeft =1080
                    LayoutCachedTop =450
                    LayoutCachedWidth =2205
                    LayoutCachedHeight =690
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =180
                    Top =1050
                    Width =870
                    Height =240
                    Name ="Label46"
                    Caption ="RAL:"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1050
                    LayoutCachedWidth =1050
                    LayoutCachedHeight =1290
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =180
                    Top =750
                    Width =870
                    Height =240
                    Name ="Label47"
                    Caption ="Schedule:"
                    LayoutCachedLeft =180
                    LayoutCachedTop =750
                    LayoutCachedWidth =1050
                    LayoutCachedHeight =990
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10080
                    Width =840
                    TabIndex =9
                    ForeColor =255
                    Name ="DocID"
                    ControlSource ="DocID"

                    LayoutCachedLeft =10080
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =240
                End
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =2
                    Left =9480
                    Top =720
                    Width =15
                    Height =915
                    Name ="Line57"
                    LayoutCachedLeft =9480
                    LayoutCachedTop =720
                    LayoutCachedWidth =9495
                    LayoutCachedHeight =1635
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =4080
                    Top =750
                    Width =2820
                    Height =255
                    FontSize =8
                    BorderColor =14270637
                    Name ="docScheduleSigner"
                    ControlSource ="docSchedSigREF"
                    RowSourceType ="Table/Query"
                    RowSource ="lookup_Contacts"
                    ColumnWidths ="0;2160;0;0"
                    FontName ="Tahoma"

                    LayoutCachedLeft =4080
                    LayoutCachedTop =750
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =1005
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =4080
                    Top =1050
                    Width =2820
                    Height =255
                    FontSize =8
                    TabIndex =3
                    BorderColor =14270637
                    Name ="docRALSigner"
                    ControlSource ="docRALSigREF"
                    RowSourceType ="Table/Query"
                    RowSource ="lookup_Contacts"
                    ColumnWidths ="0;2160;0;0"
                    FontName ="Tahoma"

                    LayoutCachedLeft =4080
                    LayoutCachedTop =1050
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =1305
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =1080
                    Top =750
                    Width =2880
                    Height =255
                    FontSize =8
                    TabIndex =2
                    BorderColor =14270637
                    Name ="docSchedContact"
                    ControlSource ="docSchedContactREF"
                    RowSourceType ="Table/Query"
                    RowSource ="lookup_Contacts"
                    ColumnWidths ="0;2160;0;0"
                    StatusBarText ="drop down"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1080
                    LayoutCachedTop =750
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =1005
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =1080
                    Top =1050
                    Width =2880
                    Height =255
                    FontSize =8
                    TabIndex =5
                    BorderColor =14270637
                    Name ="docRALContact"
                    ControlSource ="docRALContactRef"
                    RowSourceType ="Table/Query"
                    RowSource ="lookup_Contacts"
                    ColumnWidths ="0;2160;0;0"
                    StatusBarText ="drop down"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1080
                    LayoutCachedTop =1050
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =1305
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10980
                    Width =660
                    Height =255
                    TabIndex =10
                    ForeColor =2366701
                    Name ="dealRef"
                    ControlSource ="dealRef"
                    StatusBarText ="link to deals table"

                    LayoutCachedLeft =10980
                    LayoutCachedWidth =11640
                    LayoutCachedHeight =255
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =180
                    Top =1350
                    Width =870
                    Height =240
                    Name ="Label59"
                    Caption ="COA:"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1350
                    LayoutCachedWidth =1050
                    LayoutCachedHeight =1590
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =1080
                    Top =1335
                    Width =2880
                    Height =255
                    FontSize =8
                    TabIndex =7
                    BorderColor =14270637
                    Name ="docCOAContactREF"
                    ControlSource ="docCOAContactREF"
                    RowSourceType ="Table/Query"
                    RowSource ="lookup_Contacts"
                    ColumnWidths ="0;2160;0;0"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1080
                    LayoutCachedTop =1335
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =1590
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =1350
                    Width =2400
                    TabIndex =8
                    BorderColor =14270637
                    Name ="docCOASigTitle"
                    ControlSource ="docCOASigTitle"

                    LayoutCachedLeft =7020
                    LayoutCachedTop =1350
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =1590
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =4080
                    Top =1335
                    Width =2820
                    Height =255
                    FontSize =8
                    TabIndex =6
                    BorderColor =14270637
                    Name ="docCOASigREF"
                    ControlSource ="docCOASigREF"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Contacts.contactID, Contacts.contactFullName, Contacts.contactTitle, Cont"
                        "acts.clientRef FROM Contacts WHERE (((Contacts.clientRef)=[Forms]![Deals_view]!["
                        "clientRef]));"
                    ColumnWidths ="0;2160;0;0"
                    StatusBarText ="link to contacts table"
                    FontName ="Tahoma"

                    LayoutCachedLeft =4080
                    LayoutCachedTop =1335
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =1590
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =15523798
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "Documentation_sub.cls"
