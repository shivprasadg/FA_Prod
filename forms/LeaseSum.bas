Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    ViewsAllowed =1
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11280
    DatasheetFontHeight =11
    ItemSuffix =26
    Left =4575
    Top =7845
    Right =16545
    Bottom =12690
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x617138f38da0e440
    End
    RecordSource ="SELECT MstrLease.MLID, MstrLease.Client, MstrLease.MLOrig, MstrLease.MLNo, MstrL"
        "ease.MLDt, MstrLease.MLLessor FROM MstrLease WHERE (((MstrLease.Client)=[forms]!"
        "[ClientsView]![clientid])) ORDER BY MstrLease.MLOrig, MstrLease.MLNo;"
    Caption ="LeaseList"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
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
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin FormHeader
            Height =285
            BackColor =15707392
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10170
                    Width =885
                    Name ="Client"
                    ControlSource ="Client"
                    StatusBarText ="link to client file"

                    LayoutCachedLeft =10170
                    LayoutCachedWidth =11055
                    LayoutCachedHeight =240
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8220
                    Width =1455
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =-2147483615
                    Name ="OpenAddschfbtn"
                    Caption ="Add Schedule"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =8220
                    LayoutCachedWidth =9675
                    LayoutCachedHeight =285
                End
            End
        End
        Begin Section
            Height =405
            BackColor =15070679
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =6135
                    Width =1665
                    Height =285
                    FontSize =9
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="LeaseSchBtn"
                    Caption ="Lease Schedules"
                    FontName ="Tahoma"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="schedules_sub"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"LeaseSchBtn\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argume"
                                "nt Name=\"FormName\">schedules_su"
                        End
                        Begin
                            Comment ="_AXL:b</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =6135
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =270
                    Height =315
                    TabIndex =1
                    Name ="MLID"
                    ControlSource ="MLID"

                    LayoutCachedWidth =270
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1515
                    Width =1095
                    Height =315
                    TabIndex =2
                    Name ="MLOrig"
                    ControlSource ="MLOrig"

                    LayoutCachedLeft =1515
                    LayoutCachedWidth =2610
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =600
                            Width =795
                            Height =315
                            Name ="Label21"
                            Caption ="Origin"
                            LayoutCachedLeft =600
                            LayoutCachedWidth =1395
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3525
                    Width =2220
                    Height =315
                    TabIndex =3
                    Name ="MLNo"
                    ControlSource ="MLNo"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"

                    LayoutCachedLeft =3525
                    LayoutCachedWidth =5745
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2715
                            Width =720
                            Height =315
                            Name ="Label22"
                            Caption ="ML No:"
                            LayoutCachedLeft =2715
                            LayoutCachedWidth =3435
                            LayoutCachedHeight =315
                        End
                    End
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
' See "LeaseSum.cls"
