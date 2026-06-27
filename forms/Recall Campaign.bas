Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13305
    DatasheetFontHeight =11
    ItemSuffix =10
    Left =3525
    Top =3465
    Right =13395
    Bottom =7005
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x0bcb01979709e440
    End
    RecordSource ="RecallCampaign"
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
            Height =465
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =675
                    Top =15
                    Width =5295
                    Height =345
                    FontSize =14
                    Name ="Label7"
                    Caption ="Set up Recall"
                    LayoutCachedLeft =675
                    LayoutCachedTop =15
                    LayoutCachedWidth =5970
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =6405
                    Top =45
                    Width =660
                    Height =390
                    FontWeight =700
                    ForeColor =6108695
                    Name ="rcfclbbtn"
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
                                "nterfaceMacro For=\"rcfclbbtn\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></Sta"
                                "tements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =6405
                    LayoutCachedTop =45
                    LayoutCachedWidth =7065
                    LayoutCachedHeight =435
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =7335
                    Top =45
                    Width =1125
                    Height =390
                    FontWeight =700
                    TabIndex =1
                    ForeColor =6108695
                    Name ="addrecallbtn"
                    Caption ="Add Recall"
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
                                "nterfaceMacro For=\"addrecallbtn\" xmlns=\"http://schemas.microsoft.com/office/a"
                                "ccessservices/2009/11/application\"><Statements><Action Name=\"OnError\"/><Actio"
                                "n Name=\"GoToRecord\"><Argument N"
                        End
                        Begin
                            Comment ="_AXL:ame=\"Record\">New</Argument></Action><ConditionalBlock><If><Condition>[Mac"
                                "roError]&lt;&gt;0</Condition><Statements><Action Name=\"MessageBox\"><Argument N"
                                "ame=\"Message\">=[MacroError].[Description]</Argument></Action></Statements></If"
                                "></ConditionalBlock"
                        End
                        Begin
                            Comment ="_AXL:></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =7335
                    LayoutCachedTop =45
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =435
                End
            End
        End
        Begin Section
            Height =2130
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1800
                    Top =420
                    Width =2010
                    Height =315
                    Name ="RecallNo"
                    ControlSource ="RecallNo"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =420
                    LayoutCachedWidth =3810
                    LayoutCachedHeight =735
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =735
                            Top =405
                            Width =1005
                            Height =315
                            Name ="Label0"
                            Caption ="No:"
                            LayoutCachedLeft =735
                            LayoutCachedTop =405
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =720
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =735
                    Width =3195
                    Height =315
                    TabIndex =1
                    Name ="Type"
                    ControlSource ="Type"
                    RowSourceType ="Value List"
                    RowSource ="Recall;Safety Recall;Warranty;Other"
                    AllowValueListEdits =0

                    LayoutCachedLeft =735
                    LayoutCachedWidth =3930
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =15
                            Width =585
                            Height =315
                            Name ="Label1"
                            Caption ="Type:"
                            LayoutCachedLeft =15
                            LayoutCachedWidth =600
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1800
                    Top =795
                    Width =7320
                    Height =315
                    TabIndex =2
                    Name ="RecallDesc"
                    ControlSource ="RecallDesc"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =795
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =1110
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =540
                            Top =825
                            Width =1200
                            Height =315
                            Name ="Label2"
                            Caption ="Description:"
                            LayoutCachedLeft =540
                            LayoutCachedTop =825
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =1140
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4770
                    Top =420
                    Width =1035
                    Height =315
                    TabIndex =3
                    Name ="RecallDt"
                    ControlSource ="RecallDt"
                    Format ="mm/dd/yy"

                    LayoutCachedLeft =4770
                    LayoutCachedTop =420
                    LayoutCachedWidth =5805
                    LayoutCachedHeight =735
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4035
                            Top =420
                            Width =465
                            Height =315
                            Name ="Label3"
                            Caption ="Dt:"
                            LayoutCachedLeft =4035
                            LayoutCachedTop =420
                            LayoutCachedWidth =4500
                            LayoutCachedHeight =735
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1770
                    Top =1245
                    Height =315
                    TabIndex =4
                    Name ="RecallRec"
                    ControlSource ="RecallRec"
                    Format ="mm/dd/yy"

                    LayoutCachedLeft =1770
                    LayoutCachedTop =1245
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =1560
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =1260
                            Width =1650
                            Height =315
                            Name ="Label4"
                            Caption ="Notice Received:"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1260
                            LayoutCachedWidth =1710
                            LayoutCachedHeight =1575
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5550
                    Top =15
                    Width =4545
                    Height =315
                    TabIndex =5
                    Name ="Otherdesc"
                    ControlSource ="Otherdesc"

                    LayoutCachedLeft =5550
                    LayoutCachedTop =15
                    LayoutCachedWidth =10095
                    LayoutCachedHeight =330
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4305
                            Width =1140
                            Height =315
                            Name ="Label5"
                            Caption ="Other desc:"
                            LayoutCachedLeft =4305
                            LayoutCachedWidth =5445
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =1770
                    Top =1650
                    Width =4320
                    Height =315
                    TabIndex =6
                    ForeColor =1279872587
                    Name ="RecallPDFLIng"
                    ControlSource ="RecallPDFLIng"

                    LayoutCachedLeft =1770
                    LayoutCachedTop =1650
                    LayoutCachedWidth =6090
                    LayoutCachedHeight =1965
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =375
                            Top =1650
                            Width =1335
                            Height =315
                            Name ="Label6"
                            Caption ="PDF Link:"
                            LayoutCachedLeft =375
                            LayoutCachedTop =1650
                            LayoutCachedWidth =1710
                            LayoutCachedHeight =1965
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =210
            Name ="FormFooter"
        End
    End
End
