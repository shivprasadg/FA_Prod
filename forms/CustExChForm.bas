Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5490
    DatasheetFontHeight =11
    ItemSuffix =7
    Left =4395
    Top =3615
    Right =10200
    Bottom =6270
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x9774e3ff4ef7e340
    End
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    OnCurrentEmMacro = Begin
        Version =196611
        ColumnsShown =12
        Begin
            Action ="MoveSize"
            Argument ="1440"
            Argument ="1440"
            Argument ="7200"
            Argument ="2880"
        End
    End
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
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
            Height =615
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =163
                    Left =15
                    Top =15
                    Width =4245
                    Height =390
                    FontSize =16
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Label0"
                    Caption ="Customer EXchangeIt Monitor"
                    LayoutCachedLeft =15
                    LayoutCachedTop =15
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =405
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =163
                    Left =4665
                    Top =30
                    Width =660
                    Height =390
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="CLExCustBtn"
                    Caption ="Close"
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
                                "nterfaceMacro For=\"CLExCustBtn\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></S"
                                "tatements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =4665
                    LayoutCachedTop =30
                    LayoutCachedWidth =5325
                    LayoutCachedHeight =420
                End
            End
        End
        Begin Section
            Height =1101
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2310
                    Top =30
                    Width =1800
                    Height =315
                    Name ="exCLient"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Clients.ClientShNm FROM Clients WHERE (((Clients.ClientShNm) Not Like \"d"
                        "emo*\") AND ((Clients.clientType) Like \"c*\")) ORDER BY Clients.ClientShNm;"

                    LayoutCachedLeft =2310
                    LayoutCachedTop =30
                    LayoutCachedWidth =4110
                    LayoutCachedHeight =345
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =900
                            Top =30
                            Width =1305
                            Height =315
                            Name ="Label2"
                            Caption ="Select Client:"
                            LayoutCachedLeft =900
                            LayoutCachedTop =30
                            LayoutCachedWidth =2205
                            LayoutCachedHeight =345
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =900
                    Top =465
                    Width =3201
                    Height =411
                    TabIndex =1
                    ForeColor =-2147483615
                    Name ="Command5"
                    Caption ="Run Report"
                    ControlTipText ="Preview Report"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenReport"
                            Argument ="CustExItMonitor"
                            Argument ="2"
                            Argument =""
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command5\" xmlns=\"http://schemas.microsoft.com/office/acces"
                                "sservices/2009/11/application\"><Statements><Action Name=\"OpenReport\"><Argumen"
                                "t Name=\"ReportName\">CustExItMon"
                        End
                        Begin
                            Comment ="_AXL:itor</Argument><Argument Name=\"View\">Print Preview</Argument></Action></S"
                                "tatements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =900
                    LayoutCachedTop =465
                    LayoutCachedWidth =4101
                    LayoutCachedHeight =876
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
        End
    End
End
