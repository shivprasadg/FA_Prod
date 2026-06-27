Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10755
    DatasheetFontHeight =11
    ItemSuffix =26
    Left =4770
    Top =4320
    Right =15270
    Bottom =7005
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xf80cbbca5255e440
    End
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    OnCloseEmMacro = Begin
        Version =196611
        ColumnsShown =8
        Begin
            Action ="Close"
            Argument ="2"
            Argument ="bosubform"
            Argument ="0"
        End
    End
    OnLoadEmMacro = Begin
        Version =196611
        ColumnsShown =8
        Begin
            Action ="MoveSize"
            Argument ="0"
            Argument ="4320"
            Argument ="10800"
            Argument ="3600"
        End
    End
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin FormHeader
            Height =630
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =8730
                    Top =90
                    Width =660
                    Height =390
                    Name ="clrscsbtn"
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
                                "nterfaceMacro For=\"clrscsbtn\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></Sta"
                                "tements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =8730
                    LayoutCachedTop =90
                    LayoutCachedWidth =9390
                    LayoutCachedHeight =480
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =1791
            Name ="Detail"
            Begin
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2160
                    Left =1260
                    Top =60
                    Width =3600
                    Height =315
                    Name ="SelComp"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Clients.clientID, Clients.ClientShNm FROM Clients INNER JOIN Mst"
                        "rLease ON Clients.clientID = MstrLease.Client WHERE (((Clients.ClientShNm) Not L"
                        "ike \"demo*\") AND ((Clients.clientType)=\"client\" Or (Clients.clientType)=\"gu"
                        "arantor\" Or (Clients.clientType) Like \"*target\")) ORDER BY Clients.ClientShNm"
                        ";"
                    ColumnWidths ="0;2160"
                    ValidationRule ="Is Not Null"
                    OnClick ="Requery"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1260
                    LayoutCachedTop =60
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =375
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =163
                            TextAlign =3
                            Left =60
                            Top =60
                            Width =1125
                            Height =315
                            FontWeight =700
                            Name ="Label1"
                            Caption ="Company:"
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =1185
                            LayoutCachedHeight =375
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1260
                    Top =450
                    Width =3600
                    Height =315
                    TabIndex =1
                    Name ="Lease"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MstrLease.MLID, MstrLease.MLNo, MstrLease.MLOrig FROM MstrLease WHERE ((("
                        "MstrLease.Client)=[forms]![movegrp]![SelComp]));"
                    ColumnWidths ="0;2160;720"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    OnClick ="Requery"
                    OnExitEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Requery"
                            Argument ="schedule"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Lease\" xmlns=\"http://schemas.microsoft.com/office/accessse"
                                "rvices/2009/11/application\"><Statements><Action Name=\"Requery\"><Argument Name"
                                "=\"ControlName\">schedule</Argume"
                        End
                        Begin
                            Comment ="_AXL:nt></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1260
                    LayoutCachedTop =450
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =765
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =163
                            TextAlign =3
                            Left =75
                            Top =450
                            Width =1080
                            Height =315
                            FontWeight =700
                            Name ="Label3"
                            Caption ="Lease:"
                            LayoutCachedLeft =75
                            LayoutCachedTop =450
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =765
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1260
                    Top =1320
                    Width =3600
                    Height =315
                    TabIndex =2
                    Name ="group"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT SchGrp.SGrpID, SchGrp.SGrpNO FROM SchGrp WHERE (((SchGrp.SGrpSchID)=[form"
                        "s]![movegrp]![schedule]));"
                    ColumnWidths ="0;2160;720"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    OnClick ="Requery"
                    OnExitEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Requery"
                            Argument ="newsch"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"group\" xmlns=\"http://schemas.microsoft.com/office/accessse"
                                "rvices/2009/11/application\"><Statements><Action Name=\"Requery\"><Argument Name"
                                "=\"ControlName\">newsch</Argument"
                        End
                        Begin
                            Comment ="_AXL:></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1260
                    LayoutCachedTop =1320
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =1635
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =163
                            TextAlign =3
                            Left =15
                            Top =1320
                            Width =1140
                            Height =315
                            FontWeight =700
                            Name ="Label20"
                            Caption ="Sch. Group:"
                            LayoutCachedLeft =15
                            LayoutCachedTop =1320
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =1635
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1260
                    Top =885
                    Width =3600
                    Height =315
                    TabIndex =3
                    Name ="schedule"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Schedule.SchID, Schedule.SchNo FROM Schedule WHERE (((Schedule.SchMLRef)="
                        "[forms]![movegrp]![lease])) ORDER BY Schedule.SCHNo1;"
                    ColumnWidths ="0;2880"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    OnClick ="Requery"
                    OnExitEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Requery"
                            Argument ="group"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"schedule\" xmlns=\"http://schemas.microsoft.com/office/acces"
                                "sservices/2009/11/application\"><Statements><Action Name=\"Requery\"><Argument N"
                                "ame=\"ControlName\">group</Argume"
                        End
                        Begin
                            Comment ="_AXL:nt></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1260
                    LayoutCachedTop =885
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =1200
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =163
                            TextAlign =3
                            Left =75
                            Top =885
                            Width =1080
                            Height =315
                            FontWeight =700
                            Name ="Label22"
                            Caption ="Schedule:"
                            LayoutCachedLeft =75
                            LayoutCachedTop =885
                            LayoutCachedWidth =1155
                            LayoutCachedHeight =1200
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5175
                    Top =1260
                    Width =3600
                    Height =315
                    TabIndex =4
                    Name ="NewSch"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Schedule.SchID, Schedule.SchNo FROM Schedule WHERE (((Schedule.SchMLRef)="
                        "[forms]![movegrp]![lease])) ORDER BY Schedule.SCHNo1;"
                    ColumnWidths ="0;2880"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Lease is Required"
                    OnClick ="Requery"
                    OnExitEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Requery"
                            Argument ="group"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"NewSch\" xmlns=\"http://schemas.microsoft.com/office/accesss"
                                "ervices/2009/11/application\"><Statements><Action Name=\"Requery\"><Argument Nam"
                                "e=\"ControlName\">group</Argument"
                        End
                        Begin
                            Comment ="_AXL:></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =5175
                    LayoutCachedTop =1260
                    LayoutCachedWidth =8775
                    LayoutCachedHeight =1575
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =163
                            TextAlign =3
                            Left =5115
                            Top =795
                            Width =1470
                            Height =315
                            FontWeight =700
                            Name ="Label24"
                            Caption ="New Schedule:"
                            LayoutCachedLeft =5115
                            LayoutCachedTop =795
                            LayoutCachedWidth =6585
                            LayoutCachedHeight =1110
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9015
                    Top =1215
                    Height =390
                    TabIndex =5
                    Name ="mgrpbtn"
                    Caption ="Click to Move"
                    ControlTipText ="Run Query"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenQuery"
                            Argument ="moveschgrpq"
                            Argument ="0"
                            Argument ="1"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"mgrpbtn\" xmlns=\"http://schemas.microsoft.com/office/access"
                                "services/2009/11/application\"><Statements><Action Name=\"OpenQuery\"><Argument "
                                "Name=\"QueryName\">moveschgrpq</A"
                        End
                        Begin
                            Comment ="_AXL:rgument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =9015
                    LayoutCachedTop =1215
                    LayoutCachedWidth =10455
                    LayoutCachedHeight =1605
                End
            End
        End
        Begin FormFooter
            Height =270
            Name ="FormFooter"
        End
    End
End
