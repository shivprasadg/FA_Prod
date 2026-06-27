Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =20055
    DatasheetFontHeight =11
    ItemSuffix =91
    Left =3450
    Top =8520
    Right =23505
    Bottom =13185
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xb51a12a73853e440
    End
    RecordSource ="SELECT Clients.clientCompanyName, Clients.ClientShNm, RmktBills.[Invoice No], Rm"
        "ktBills.Billdt, RmktBills.RBillDueDt, RmktBills.BillSent, RmktBills.RevNo, RmktB"
        "ills.BillLink, RmktBills.RevReason, RmktBills.UseTxRt, RmktBills.ReqDt FROM Rmkt"
        "Bills INNER JOIN Clients ON RmktBills.BuyerRef = Clients.clientID WHERE (((Clien"
        "ts.clientCompanyName)=[forms]![RmktBillsel]![Selbuyer])) ORDER BY Clients.client"
        "CompanyName;"
    DatasheetFontName ="Calibri"
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
        Begin FormHeader
            Height =1320
            BackColor =15707392
            Name ="FormHeader"
            Begin
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =161
                    Left =1650
                    Top =705
                    Width =1770
                    Height =375
                    FontSize =12
                    FontWeight =700
                    BackColor =15707392
                    ForeColor =16777215
                    Name ="Label18"
                    Caption ="Invoice No"
                    LayoutCachedLeft =1650
                    LayoutCachedTop =705
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =1080
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =161
                    Left =5250
                    Top =720
                    Width =1035
                    Height =375
                    FontSize =12
                    FontWeight =700
                    BackColor =15707392
                    ForeColor =16777215
                    Name ="Label20"
                    Caption ="Bill Dt"
                    LayoutCachedLeft =5250
                    LayoutCachedTop =720
                    LayoutCachedWidth =6285
                    LayoutCachedHeight =1095
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =161
                    TextAlign =2
                    Left =8880
                    Top =735
                    Width =1320
                    Height =375
                    FontSize =12
                    FontWeight =700
                    BackColor =15707392
                    ForeColor =16777215
                    Name ="Label23"
                    Caption ="Bill Due "
                    LayoutCachedLeft =8880
                    LayoutCachedTop =735
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =1110
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =161
                    TextAlign =2
                    Left =7230
                    Top =720
                    Width =1440
                    Height =375
                    FontSize =12
                    FontWeight =700
                    BackColor =15707392
                    ForeColor =16777215
                    Name ="Label24"
                    Caption ="Bill Sent"
                    LayoutCachedLeft =7230
                    LayoutCachedTop =720
                    LayoutCachedWidth =8670
                    LayoutCachedHeight =1095
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =14655
                    Width =885
                    Height =390
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Command32"
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
                                "nterfaceMacro For=\"Command32\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></Sta"
                                "tements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =14655
                    LayoutCachedWidth =15540
                    LayoutCachedHeight =390
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =13695
                    Width =885
                    Height =390
                    FontWeight =700
                    TabIndex =1
                    ForeColor =-2147483615
                    Name ="AddBillBtn"
                    Caption ="Add Bill"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =12
                        Begin
                            Action ="OpenForm"
                            Argument ="NewBill"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="0"
                            Argument ="0"
                        End
                        Begin
                            Action ="SetValue"
                            Argument ="[Forms]![NewBill]![BillSchRef]"
                            Argument ="[Forms]![pd billsel]![selsch]"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"AddBillBtn\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argumen"
                                "t Name=\"FormName\">NewBill</Argu"
                        End
                        Begin
                            Comment ="_AXL:ment><Argument Name=\"DataMode\">Add</Argument></Action><Action Name=\"SetV"
                                "alue\"><Argument Name=\"Item\">[Forms]![NewBill]![BillSchRef]</Argument><Argumen"
                                "t Name=\"Expression\">[Forms]![pd billsel]![selsch]</Argument></Action></Stateme"
                                "nts></UserInterfaceMa"
                        End
                        Begin
                            Comment ="_AXL:cro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =13695
                    LayoutCachedWidth =14580
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    Left =105
                    Top =60
                    Width =12675
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    BackColor =15707392
                    BorderColor =0
                    ForeColor =16777215
                    Name ="Text54"
                    ControlSource ="=UCase([Forms]![RmktBillSel]![Selbuyer])"

                    LayoutCachedLeft =105
                    LayoutCachedTop =60
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =510
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =161
                    Left =10545
                    Top =735
                    Width =630
                    Height =375
                    FontSize =12
                    FontWeight =700
                    BackColor =15707392
                    ForeColor =16777215
                    Name ="Label63"
                    Caption ="Rev"
                    LayoutCachedLeft =10545
                    LayoutCachedTop =735
                    LayoutCachedWidth =11175
                    LayoutCachedHeight =1110
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =161
                    Left =11310
                    Top =765
                    Width =4725
                    Height =375
                    FontSize =12
                    FontWeight =700
                    BackColor =15707392
                    ForeColor =16777215
                    Name ="Label64"
                    Caption ="Reason"
                    LayoutCachedLeft =11310
                    LayoutCachedTop =765
                    LayoutCachedWidth =16035
                    LayoutCachedHeight =1140
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =161
                    TextAlign =2
                    Left =135
                    Top =705
                    Width =1395
                    Height =375
                    FontSize =12
                    FontWeight =700
                    BackColor =15707392
                    ForeColor =16777215
                    Name ="Label83"
                    Caption ="Request Dt"
                    LayoutCachedLeft =135
                    LayoutCachedTop =705
                    LayoutCachedWidth =1530
                    LayoutCachedHeight =1080
                End
                Begin Label
                    FontUnderline = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextFontCharSet =161
                    Left =3630
                    Top =705
                    Width =1440
                    Height =375
                    FontSize =12
                    FontWeight =700
                    BackColor =15707392
                    ForeColor =16777215
                    Name ="Label88"
                    Caption ="ReqDt:"
                    LayoutCachedLeft =3630
                    LayoutCachedTop =705
                    LayoutCachedWidth =5070
                    LayoutCachedHeight =1080
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =1035
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =161
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15435
                    Width =435
                    Height =45
                    FontSize =10
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="BillID"
                    ControlSource ="BillID"

                    LayoutCachedLeft =15435
                    LayoutCachedWidth =15870
                    LayoutCachedHeight =45
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =161
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14055
                    Width =1320
                    Height =15
                    FontSize =9
                    TabIndex =1
                    BorderColor =0
                    ForeColor =5026082
                    Name ="BillSchRef"
                    ControlSource ="BillSchRef"
                    StatusBarText ="link to schedule table"

                    LayoutCachedLeft =14055
                    LayoutCachedWidth =15375
                    LayoutCachedHeight =15
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =18240
                    Top =480
                    Width =780
                    Height =315
                    TabIndex =2
                    Name ="Text72"
                    ControlSource ="BillID"

                    LayoutCachedLeft =18240
                    LayoutCachedTop =480
                    LayoutCachedWidth =19020
                    LayoutCachedHeight =795
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            Left =17460
                            Top =480
                            Width =630
                            Height =315
                            Name ="Label73"
                            Caption ="BillID:"
                            LayoutCachedLeft =17460
                            LayoutCachedTop =480
                            LayoutCachedWidth =18090
                            LayoutCachedHeight =795
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1665
                    Width =1770
                    Height =315
                    TabIndex =3
                    Name ="Invoice No"
                    ControlSource ="Invoice No"
                    EventProcPrefix ="Invoice_No"

                    LayoutCachedLeft =1665
                    LayoutCachedWidth =3435
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5220
                    Height =315
                    TabIndex =4
                    Name ="Billdt"
                    ControlSource ="Billdt"
                    Format ="mm/dd/yy"

                    LayoutCachedLeft =5220
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7260
                    Height =315
                    TabIndex =5
                    Name ="BillSent"
                    ControlSource ="BillSent"
                    Format ="mm/dd/yy"

                    LayoutCachedLeft =7260
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8925
                    Width =1305
                    Height =315
                    TabIndex =6
                    Name ="RBillDueDt"
                    ControlSource ="RBillDueDt"
                    Format ="mm/dd/yy"

                    LayoutCachedLeft =8925
                    LayoutCachedWidth =10230
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =11310
                    Width =4725
                    Height =315
                    TabIndex =7
                    Name ="RevReason"
                    ControlSource ="RevReason"

                    LayoutCachedLeft =11310
                    LayoutCachedWidth =16035
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10500
                    Width =690
                    Height =315
                    TabIndex =8
                    Name ="RevNo"
                    ControlSource ="RevNo"

                    LayoutCachedLeft =10500
                    LayoutCachedWidth =11190
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =7455
                    Top =435
                    Width =10725
                    Height =363
                    TabIndex =9
                    Name ="BillLink"
                    ControlSource ="BillLink"

                    LayoutCachedLeft =7455
                    LayoutCachedTop =435
                    LayoutCachedWidth =18180
                    LayoutCachedHeight =798
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6150
                            Top =420
                            Width =855
                            Height =315
                            Name ="Label81"
                            Caption ="Bill Link:"
                            LayoutCachedLeft =6150
                            LayoutCachedTop =420
                            LayoutCachedWidth =7005
                            LayoutCachedHeight =735
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =195
                    Width =1335
                    Height =315
                    TabIndex =10
                    Name ="ReqDt"
                    ControlSource ="ReqDt"

                    LayoutCachedLeft =195
                    LayoutCachedWidth =1530
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3630
                    Height =315
                    TabIndex =11
                    Name ="Text87"
                    ControlSource ="ReqDt"

                    LayoutCachedLeft =3630
                    LayoutCachedWidth =5070
                    LayoutCachedHeight =315
                End
                Begin Line
                    OverlapFlags =85
                    Left =45
                    Top =840
                    Width =19935
                    Name ="Line90"
                    LayoutCachedLeft =45
                    LayoutCachedTop =840
                    LayoutCachedWidth =19980
                    LayoutCachedHeight =840
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
' See "RmktBillUpdt.cls"
