Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
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
    Width =7935
    DatasheetFontHeight =11
    ItemSuffix =17
    Left =8010
    Top =2460
    Right =16560
    Bottom =9240
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x1a72cf5c32f1e340
    End
    Caption ="CUSTOMER DELIVERY REPORT SELECTION"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    OnLoadEmMacro = Begin
        Version =196611
        ColumnsShown =8
        Begin
            Action ="MoveSize"
            Argument ="1440"
            Argument ="1440"
            Argument ="8640"
            Argument ="7200"
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
            Height =645
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =15
                    Top =150
                    Width =5085
                    Height =360
                    FontSize =14
                    FontWeight =700
                    ForeColor =10040879
                    Name ="Label6"
                    Caption ="CUSTOMER DELIVERY REPORT SELECTION"
                    LayoutCachedLeft =15
                    LayoutCachedTop =150
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =510
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =6270
                    Top =165
                    Width =855
                    Height =390
                    FontWeight =700
                    ForeColor =10040879
                    Name ="Command12"
                    Caption ="Close"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Action ="Close"
                            Argument ="2"
                            Argument ="CustDelSel"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command12\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"><Argum"
                                "ent Name=\"ObjectType\">Form</Arg"
                        End
                        Begin
                            Comment ="_AXL:ument><Argument Name=\"ObjectName\">CustDelSel</Argument></Action></Stateme"
                                "nts></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6270
                    LayoutCachedTop =165
                    LayoutCachedWidth =7125
                    LayoutCachedHeight =555
                End
            End
        End
        Begin Section
            Height =3810
            Name ="Detail"
            Begin
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5295
                    Top =465
                    Width =2505
                    Height =315
                    Name ="selcust"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Clients.ClientShNm, Clients.[client group] FROM (((Clients INNER"
                        " JOIN MstrLease ON Clients.clientID = MstrLease.Client) INNER JOIN Schedule ON M"
                        "strLease.MLID = Schedule.SchMLRef) INNER JOIN SchGrp ON Schedule.SchID = SchGrp."
                        "SGrpSchID) INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF WHERE (((Clients.["
                        "client group])<>\"client demo\") AND ((Clients.clientType)=\"CLIENT\") AND ((Uni"
                        "ts.unitstatus)=\"b\")) OR (((Clients.[client group])<>\"client demo\") AND ((Cli"
                        "ents.clientType)=\"CLIENT\") AND ((Units.unitstatus)=\"a\") AND ((Units.unitinse"
                        "rvind)<>\"a\")) ORDER BY Clients.ClientShNm;"
                    Format =">"

                    LayoutCachedLeft =5295
                    LayoutCachedTop =465
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =780
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =465
                            Width =5055
                            Height =315
                            Name ="Label1"
                            Caption ="or for Specific Customer (leave Customer group blank)"
                            LayoutCachedTop =465
                            LayoutCachedWidth =5055
                            LayoutCachedHeight =780
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5295
                    Width =2505
                    Height =315
                    TabIndex =1
                    Name ="selgrp"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Clients.[CLIENT GROUP] FROM (((Clients INNER JOIN MstrLease ON C"
                        "lients.clientID = MstrLease.Client) INNER JOIN Schedule ON MstrLease.MLID = Sche"
                        "dule.SchMLRef) INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID) INNER JOI"
                        "N Units ON SchGrp.SGrpID = Units.GroupREF WHERE (((Clients.[CLIENT GROUP]) Is No"
                        "t Null And (Clients.[CLIENT GROUP])<>\"client demo\") AND ((Units.unitstatus)=\""
                        "b\") AND ((Clients.clientType)=\"client\")) OR (((Clients.[CLIENT GROUP]) Is Not"
                        " Null And (Clients.[CLIENT GROUP])<>\"client demo\") AND ((Units.unitstatus)=\"a"
                        "\") AND ((Clients.clientType)=\"client\") AND ((Units.unitinservind)<>\"a\")) OR"
                        "DER BY Clients.[CLIENT GROUP];"
                    Format =">"

                    LayoutCachedLeft =5295
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =210
                            Top =45
                            Width =4860
                            Height =315
                            Name ="Label8"
                            Caption ="for Customer Group (lease Specific Customer blank)"
                            LayoutCachedLeft =210
                            LayoutCachedTop =45
                            LayoutCachedWidth =5070
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =5295
                    Top =1380
                    Width =2505
                    Height =390
                    FontWeight =700
                    TabIndex =2
                    ForeColor =10040879
                    Name ="Command9"
                    Caption ="Preview Report"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Condition ="[Forms]![custdelsel]![selgrp] Is Not Null And [Forms]![custdelsel]![selcust] Is "
                                "Not Null"
                            Action ="MsgBox"
                            Argument ="Select Group or Specific Customer not both"
                            Argument ="-1"
                            Argument ="1"
                            Argument ="Selection Error"
                        End
                        Begin
                            Condition ="..."
                            Action ="StopMacro"
                        End
                        Begin
                            Condition ="[Forms]![custdelsel]![selprov] Is Not Null"
                            Action ="OpenReport"
                            Argument ="DeliveryRptMainesProv"
                            Argument ="2"
                            Argument =""
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Condition ="[Forms]![custdelsel]![selprov] Is Null"
                            Action ="OpenReport"
                            Argument ="DeliveryRptDetMainesA"
                            Argument ="2"
                            Argument =""
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command9\" xmlns=\"http://schemas.microsoft.com/office/acces"
                                "sservices/2009/11/application\"><Statements><ConditionalBlock><If><Condition>[Fo"
                                "rms]![custdelsel]![selgrp] Is"
                        End
                        Begin
                            Comment ="_AXL: Not Null And [Forms]![custdelsel]![selcust] Is Not Null</Condition><Statem"
                                "ents><Action Name=\"MessageBox\"><Argument Name=\"Message\">Select Group or Spec"
                                "ific Customer not both</Argument><Argument Name=\"Type\">Critical</Argument><Arg"
                                "ument Name=\"Title\">"
                        End
                        Begin
                            Comment ="_AXL:Selection Error</Argument></Action><Action Name=\"StopMacro\"/></Statements"
                                "></If></ConditionalBlock><ConditionalBlock><If><Condition>[Forms]![custdelsel]!["
                                "selprov] Is Not Null</Condition><Statements><Action Name=\"OpenReport\"><Argumen"
                                "t Name=\"ReportNam"
                        End
                        Begin
                            Comment ="_AXL:e\">DeliveryRptMainesProv</Argument><Argument Name=\"View\">Print Preview</"
                                "Argument></Action></Statements></If></ConditionalBlock><ConditionalBlock><If><Co"
                                "ndition>[Forms]![custdelsel]![selprov] Is Null</Condition><Statements><Action Na"
                                "me=\"OpenReport\">"
                        End
                        Begin
                            Comment ="_AXL:<Argument Name=\"ReportName\">DeliveryRptDetMainesA</Argument><Argument Nam"
                                "e=\"View\">Print Preview</Argument></Action></Statements></If></ConditionalBlock"
                                "></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =5295
                    LayoutCachedTop =1380
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =1770
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =5295
                    Top =1830
                    Width =2505
                    Height =390
                    FontWeight =700
                    TabIndex =3
                    ForeColor =10040879
                    Name ="Command10"
                    Caption ="Export to Excel"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5295
                    LayoutCachedTop =1830
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =2220
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    Top =2490
                    Width =5085
                    Height =360
                    FontSize =14
                    FontWeight =700
                    ForeColor =10040879
                    Name ="Label13"
                    Caption ="Maines Provider Fleet Report"
                    LayoutCachedTop =2490
                    LayoutCachedWidth =5085
                    LayoutCachedHeight =2850
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5205
                    Top =3015
                    Width =2505
                    Height =315
                    TabIndex =4
                    Name ="MAINTPROV"
                    RowSourceType ="Value List"
                    RowSource ="AIM;Penske"
                    AllowValueListEdits =0

                    LayoutCachedLeft =5205
                    LayoutCachedTop =3015
                    LayoutCachedWidth =7710
                    LayoutCachedHeight =3330
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3405
                            Top =3015
                            Width =1545
                            Height =315
                            Name ="Label15"
                            Caption ="Select Provider:"
                            LayoutCachedLeft =3405
                            LayoutCachedTop =3015
                            LayoutCachedWidth =4950
                            LayoutCachedHeight =3330
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =5205
                    Top =3420
                    Width =2505
                    Height =390
                    FontWeight =700
                    TabIndex =5
                    ForeColor =10040879
                    Name ="MaineProvq"
                    Caption ="Export to Excel"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenQuery"
                            Argument ="MainProvUnits"
                            Argument ="0"
                            Argument ="1"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"MaineProvq\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenQuery\"><Argume"
                                "nt Name=\"QueryName\">MainProvUni"
                        End
                        Begin
                            Comment ="_AXL:ts</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5205
                    LayoutCachedTop =3420
                    LayoutCachedWidth =7710
                    LayoutCachedHeight =3810
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
' See "CustDelSel.cls"
