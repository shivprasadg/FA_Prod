Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10200
    DatasheetFontHeight =11
    ItemSuffix =11
    Left =9585
    Top =2295
    Right =20040
    Bottom =12945
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x3229a145cefde340
    End
    RecordSource ="SELECT SgrpExp.*, MstrLease.MLNo, Schedule.SchNo, SchGrp.SGrpID, SchGrp.SGrpNO F"
        "ROM ((MstrLease INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef) INNER"
        " JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID) INNER JOIN SgrpExp ON SchGrp."
        "SGrpID = SgrpExp.SgrpExpSGRef WHERE (((SchGrp.SGrpID)=[forms]![AcctMenu]![SelGrp"
        "])) ORDER BY SchGrp.SGrpNO, SgrpExp.SgrpExpType;"
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
        Begin FormHeader
            Height =405
            BackColor =15707392
            Name ="FormHeader"
            AutoHeight =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =204
                    BackStyle =0
                    IMESentenceMode =3
                    Top =75
                    Width =7245
                    Height =315
                    FontSize =14
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Text5"
                    ControlSource ="=\"Expenses for \" & [MLNo] & \" Sch \" & [SChNo] & \" Grp \" & [SgrpNO]"

                    LayoutCachedTop =75
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =390
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =9000
                    Width =1050
                    Height =390
                    FontWeight =700
                    TabIndex =1
                    ForeColor =-2147483615
                    Name ="ClGrpExpBtn"
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
                                "nterfaceMacro For=\"ClGrpExpBtn\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></S"
                                "tatements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =9000
                    LayoutCachedWidth =10050
                    LayoutCachedHeight =390
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =7425
                    Width =1050
                    Height =390
                    FontWeight =700
                    TabIndex =2
                    ForeColor =-2147483615
                    Name ="AddNExpBtn"
                    Caption ="Add New"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="NewSGExpForm"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="0"
                            Argument ="0"
                        End
                        Begin
                            Action ="SetValue"
                            Argument ="[Forms]![NewSGExpForm]![SgrpExpSGRef]"
                            Argument ="[Forms]![AcctMenu]![SelGrp]"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"AddNExpBtn\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argumen"
                                "t Name=\"FormName\">NewSGExpForm<"
                        End
                        Begin
                            Comment ="_AXL:/Argument><Argument Name=\"DataMode\">Add</Argument></Action><Action Name=\""
                                "SetValue\"><Argument Name=\"Item\">[Forms]![NewSGExpForm]![SgrpExpSGRef]</Argume"
                                "nt><Argument Name=\"Expression\">[Forms]![AcctMenu]![SelGrp]</Argument></Action>"
                                "</Statements></UserI"
                        End
                        Begin
                            Comment ="_AXL:nterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =7425
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =390
                End
            End
        End
        Begin Section
            Height =1260
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7170
                    Top =105
                    Height =315
                    Name ="SgrpExpId"
                    ControlSource ="SgrpExpId"
                    StatusBarText ="id for expenses associated with a schedule group which are not calculate like i."
                        "e. interest or Peoplenet expense, Orig Lease refunds"

                    LayoutCachedLeft =7170
                    LayoutCachedTop =105
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5820
                            Top =105
                            Width =1050
                            Height =315
                            Name ="Label0"
                            Caption ="SgrpExpId:"
                            LayoutCachedLeft =5820
                            LayoutCachedTop =105
                            LayoutCachedWidth =6870
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7170
                    Top =500
                    Height =315
                    TabIndex =1
                    Name ="SgrpExpSGRef"
                    ControlSource ="SgrpExpSGRef"
                    StatusBarText ="reference to schedule group which had this expense"

                    LayoutCachedLeft =7170
                    LayoutCachedTop =500
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =815
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5445
                            Top =500
                            Width =1425
                            Height =315
                            Name ="Label1"
                            Caption ="SgrpExpSGRef:"
                            LayoutCachedLeft =5445
                            LayoutCachedTop =500
                            LayoutCachedWidth =6870
                            LayoutCachedHeight =815
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2325
                    Top =135
                    Width =2010
                    Height =315
                    TabIndex =2
                    Name ="SgrpExp"
                    ControlSource ="SgrpExp"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="expenses all entered as positive nubmers"

                    LayoutCachedLeft =2325
                    LayoutCachedTop =135
                    LayoutCachedWidth =4335
                    LayoutCachedHeight =450
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =405
                            Top =105
                            Width =1710
                            Height =315
                            Name ="Label2"
                            Caption ="Expense Amount:"
                            LayoutCachedLeft =405
                            LayoutCachedTop =105
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2310
                    Top =920
                    Width =4950
                    Height =315
                    TabIndex =4
                    Name ="SgrpExpOthExp"
                    ControlSource ="SgrpExpOthExp"
                    StatusBarText ="explain expense type of other"

                    LayoutCachedLeft =2310
                    LayoutCachedTop =920
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =1235
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =945
                            Width =1995
                            Height =315
                            Name ="Label4"
                            Caption ="Explanation if Other:"
                            LayoutCachedLeft =120
                            LayoutCachedTop =945
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =1260
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2325
                    Top =530
                    Width =2025
                    Height =315
                    TabIndex =3
                    Name ="SgrpExpType"
                    ControlSource ="SgrpExpType"
                    RowSourceType ="Value List"
                    RowSource ="Attorney Fee;Depreciation;Other"

                    LayoutCachedLeft =2325
                    LayoutCachedTop =530
                    LayoutCachedWidth =4350
                    LayoutCachedHeight =845
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =720
                            Top =525
                            Width =1395
                            Height =315
                            Name ="Label3"
                            Caption ="Espense Type:"
                            LayoutCachedLeft =720
                            LayoutCachedTop =525
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =840
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =186
                    IMESentenceMode =3
                    Left =8115
                    Top =945
                    Width =870
                    Height =255
                    FontSize =9
                    TabIndex =5
                    ForeColor =5167783
                    Name ="Text9"
                    ControlSource ="SgrpExpId"
                    StatusBarText ="id for expenses associated with a schedule group which are not calculate like i."
                        "e. interest or Peoplenet expense, Orig Lease refunds"

                    LayoutCachedLeft =8115
                    LayoutCachedTop =945
                    LayoutCachedWidth =8985
                    LayoutCachedHeight =1200
                End
            End
        End
        Begin FormFooter
            Height =180
            BackColor =15707392
            Name ="FormFooter"
        End
    End
End
