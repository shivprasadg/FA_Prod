Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9960
    DatasheetFontHeight =11
    ItemSuffix =20
    Left =6735
    Top =3855
    Right =16695
    Bottom =6780
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xa7a066e1fa92e340
    End
    Caption ="Attach a Document"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
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
            Visible = NotDefault
            Height =480
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4440
                    Width =480
                    FontSize =8
                    ForeColor =1643706
                    Name ="clientREF"
                    FontName ="Tahoma"

                    LayoutCachedLeft =4440
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =240
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3600
                            Width =795
                            Height =240
                            FontSize =8
                            Name ="Label5"
                            Caption ="ClientRef:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =3600
                            LayoutCachedWidth =4395
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5640
                    Width =540
                    FontSize =8
                    TabIndex =1
                    ForeColor =1643706
                    Name ="DealRef"
                    FontName ="Tahoma"

                    LayoutCachedLeft =5640
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =240
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4980
                            Width =705
                            Height =240
                            FontSize =8
                            Name ="Label9"
                            Caption ="DealRef:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =4980
                            LayoutCachedWidth =5685
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7080
                    Width =240
                    FontSize =8
                    TabIndex =2
                    ForeColor =1643706
                    Name ="GroupRef"
                    FontName ="Tahoma"

                    LayoutCachedLeft =7080
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =240
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6240
                            Width =825
                            Height =240
                            FontSize =8
                            Name ="Label7"
                            Caption ="GroupRef:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =6240
                            LayoutCachedWidth =7065
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8100
                    Width =300
                    FontSize =8
                    TabIndex =3
                    ForeColor =1643706
                    Name ="UnitRef"
                    FontName ="Tahoma"

                    LayoutCachedLeft =8100
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =240
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7380
                            Width =675
                            Height =240
                            FontSize =8
                            Name ="Label11"
                            Caption ="UnitRef:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =7380
                            LayoutCachedWidth =8055
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9420
                    Width =420
                    FontSize =8
                    TabIndex =4
                    ForeColor =1643706
                    Name ="PricingRef"
                    FontName ="Tahoma"

                    LayoutCachedLeft =9420
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =240
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8580
                            Width =825
                            Height =240
                            FontSize =8
                            Name ="Label13"
                            Caption ="PrcingRef:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =8580
                            LayoutCachedWidth =9405
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =1890
                    Height =315
                    FontWeight =700
                    Name ="Label18"
                    Caption ="Add a Document"
                    FontName ="Tahoma"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1950
                    LayoutCachedHeight =375
                End
            End
        End
        Begin Section
            Height =2760
            BackColor =15523798
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =60
                    Width =7200
                    Height =270
                    FontSize =9
                    Name ="txtfilename"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =60
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =330
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =60
                            Width =960
                            Height =285
                            FontSize =9
                            Name ="Label1"
                            Caption ="File Path:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =120
                            LayoutCachedTop =60
                            LayoutCachedWidth =1080
                            LayoutCachedHeight =345
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8460
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8210719
                    Name ="bbrowse"
                    Caption ="Browse"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =8460
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =360
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1140
                    Top =420
                    Width =2820
                    Height =270
                    FontSize =9
                    TabIndex =2
                    Name ="AttachedFileTypeRef"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lookup_FileTypes.FileTypeID, lookup_FileTypes.FileType FROM lookup_FileTy"
                        "pes ORDER BY lookup_FileTypes.FileType; "
                    ColumnWidths ="0;1440"
                    StatusBarText ="Link to lookup_fileTypes table"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =420
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =690
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =780
                    Width =6120
                    Height =270
                    FontSize =9
                    TabIndex =3
                    Name ="AttachedFileDescShort"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =780
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =1050
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =1140
                    Width =8580
                    Height =945
                    FontSize =9
                    TabIndex =4
                    Name ="AttachedFileDescLong"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =1140
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =2085
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =120
                    Top =420
                    Width =960
                    Height =285
                    FontSize =9
                    Name ="Label14"
                    Caption ="File Type:"
                    FontName ="Tahoma"
                    LayoutCachedLeft =120
                    LayoutCachedTop =420
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =705
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =120
                    Top =780
                    Width =960
                    Height =285
                    FontSize =9
                    Name ="Label15"
                    Caption ="Name:"
                    FontName ="Tahoma"
                    LayoutCachedLeft =120
                    LayoutCachedTop =780
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =1065
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =120
                    Top =1140
                    Width =960
                    Height =285
                    FontSize =9
                    Name ="Label16"
                    Caption ="Long Desc:"
                    FontName ="Tahoma"
                    LayoutCachedLeft =120
                    LayoutCachedTop =1140
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =1425
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =2280
                    Width =1980
                    Height =390
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    ForeColor =-2147483615
                    Name ="AddFileBtn"
                    Caption ="Save/Add File"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =120
                    LayoutCachedTop =2280
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =2670
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4020
                    Top =420
                    Width =1860
                    Height =315
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    ForeColor =8210719
                    Name ="AddFiletypeBtn"
                    Caption ="Add a file type"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =420
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =735
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2460
                    Top =2280
                    Width =1980
                    Height =390
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    ForeColor =2366701
                    Name ="CancelBtn"
                    Caption ="Cancel/Close Form"
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
                                "nterfaceMacro For=\"CancelBtn\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></Sta"
                                "tements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =2460
                    LayoutCachedTop =2280
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =2670
                End
            End
        End
        Begin FormFooter
            Height =180
            BackColor =16249583
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "Attach_File.cls"
