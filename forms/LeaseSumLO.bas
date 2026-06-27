Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
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
    Width =8970
    DatasheetFontHeight =11
    ItemSuffix =13
    Left =4620
    Top =7575
    Right =13590
    Bottom =9450
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xc0b38c1e7906e440
    End
    RecordSource ="SELECT MstrLease.MLID, MstrLease.Client, MstrLease.MLNo, MstrLease.MLOrig, MstrL"
        "ease.MLDt, MstrLease.MLLessor FROM MstrLease ORDER BY MstrLease.MLOrig;"
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
            Height =315
            BackColor =15707392
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =2010
                    Width =1320
                    Height =315
                    ForeColor =16777215
                    Name ="Label3"
                    Caption ="Origin:"
                    LayoutCachedLeft =2010
                    LayoutCachedWidth =3330
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =85
                    Left =3450
                    Width =990
                    Height =315
                    ForeColor =16777215
                    Name ="Label4"
                    Caption ="Lessor:"
                    LayoutCachedLeft =3450
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =85
                    Width =675
                    Height =315
                    ForeColor =16777215
                    Name ="Label8"
                    Caption ="MLNo:"
                    LayoutCachedWidth =675
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6300
                    Height =315
                    Name ="Client"
                    ControlSource ="Client"
                    StatusBarText ="link to client file"

                    LayoutCachedLeft =6300
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =315
                End
            End
        End
        Begin Section
            Height =360
            BackColor =15070679
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1500
                    Height =0
                    Name ="MLID"
                    ControlSource ="MLID"

                    LayoutCachedLeft =1500
                    LayoutCachedWidth =2940
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3450
                    Height =285
                    TabIndex =2
                    Name ="MLLessor"
                    ControlSource ="MLLessor"
                    Format =">"
                    StatusBarText ="If other name of lessor"

                    LayoutCachedLeft =3450
                    LayoutCachedWidth =4890
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1950
                    Width =1500
                    Height =285
                    TabIndex =1
                    Name ="MLOrig"
                    ControlSource ="MLOrig"
                    Format =">"

                    LayoutCachedLeft =1950
                    LayoutCachedWidth =3450
                    LayoutCachedHeight =285
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5550
                    Top =60
                    Width =1665
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    ForeColor =-2147483615
                    Name ="LeaseSchBtn"
                    Caption ="Lease Schedules"
                    FontName ="Tahoma"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="schedules_sublo"
                            Argument ="0"
                            Argument =""
                            Argument ="=\"[MLno]=\" & \"'\" & [MLno] & \"'\""
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
                            Comment ="_AXL:blo</Argument><Argument Name=\"WhereCondition\">=\"[MLno]=\" &amp; \"'\" &a"
                                "mp; [MLno] &amp; \"'\"</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =5550
                    LayoutCachedTop =60
                    LayoutCachedWidth =7215
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Width =1845
                    Height =285
                    TabIndex =4
                    Name ="MLNo"
                    ControlSource ="MLNo"
                    Format =">"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"

                    LayoutCachedWidth =1845
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4965
                    Width =435
                    Height =315
                    TabIndex =5
                    Name ="v"
                    ControlSource ="MLID"

                    LayoutCachedLeft =4965
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =315
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
        End
    End
End
