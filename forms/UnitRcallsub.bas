Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15909
    DatasheetFontHeight =11
    ItemSuffix =33
    Left =5325
    Top =13410
    Right =21600
    Bottom =18975
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x7ae40fe6b597e540
    End
    OnClose ="[Event Procedure]"
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
            Height =555
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    Top =240
                    Width =1860
                    Height =315
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Recall #:"
                    LayoutCachedTop =240
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =555
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =1905
                    Top =240
                    Width =585
                    Height =315
                    FontWeight =700
                    Name ="Label1"
                    Caption ="Type:"
                    LayoutCachedLeft =1905
                    LayoutCachedTop =240
                    LayoutCachedWidth =2490
                    LayoutCachedHeight =555
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =3390
                    Top =240
                    Width =1575
                    Height =315
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Description:"
                    LayoutCachedLeft =3390
                    LayoutCachedTop =240
                    LayoutCachedWidth =4965
                    LayoutCachedHeight =555
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    Left =6375
                    Top =240
                    Width =1275
                    Height =315
                    FontWeight =700
                    Name ="Label3"
                    Caption ="Recall Date:"
                    LayoutCachedLeft =6375
                    LayoutCachedTop =240
                    LayoutCachedWidth =7650
                    LayoutCachedHeight =555
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    Left =7725
                    Width =1260
                    Height =555
                    FontWeight =700
                    Name ="Label4"
                    Caption ="FA \015\012Received:"
                    LayoutCachedLeft =7725
                    LayoutCachedWidth =8985
                    LayoutCachedHeight =555
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    Left =10185
                    Top =240
                    Width =1065
                    Height =315
                    FontWeight =700
                    Name ="Label22"
                    Caption ="Inspected:"
                    LayoutCachedLeft =10185
                    LayoutCachedTop =240
                    LayoutCachedWidth =11250
                    LayoutCachedHeight =555
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    Left =11319
                    Width =1005
                    Height =555
                    FontWeight =700
                    Name ="Label24"
                    Caption ="Repair \015\012Needed:"
                    LayoutCachedLeft =11319
                    LayoutCachedWidth =12324
                    LayoutCachedHeight =555
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    Left =12369
                    Width =990
                    Height =555
                    FontWeight =700
                    Name ="Label26"
                    Caption ="Repair \015\012Made:"
                    LayoutCachedLeft =12369
                    LayoutCachedWidth =13359
                    LayoutCachedHeight =555
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =13434
                    Top =210
                    Width =1710
                    Height =315
                    FontWeight =700
                    Name ="Label28"
                    Caption ="Contact:"
                    LayoutCachedLeft =13434
                    LayoutCachedTop =210
                    LayoutCachedWidth =15144
                    LayoutCachedHeight =525
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    Left =9090
                    Width =990
                    Height =555
                    FontWeight =700
                    Name ="Label32"
                    Caption ="Recall\015\012Sent"
                    LayoutCachedLeft =9090
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =555
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =315
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =1830
                    Height =270
                    Name ="RecallNo"
                    ControlSource ="RecallNo"

                    LayoutCachedWidth =1830
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =3330
                    Width =2940
                    Height =270
                    ColumnWidth =5565
                    FontSize =10
                    TabIndex =2
                    Name ="Descsription"
                    ControlSource ="Description"

                    LayoutCachedLeft =3330
                    LayoutCachedWidth =6270
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6390
                    Width =1260
                    Height =270
                    TabIndex =3
                    Name ="Recall Dt"
                    ControlSource ="RecallDate"
                    Format ="mm/dd/yy"
                    InputMask ="99/99/00;"
                    EventProcPrefix ="Recall_Dt"

                    LayoutCachedLeft =6390
                    LayoutCachedWidth =7650
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7725
                    Width =1260
                    Height =270
                    TabIndex =4
                    Name ="RA Received"
                    ControlSource ="ReCallRec"
                    Format ="mm/dd/yy"
                    InputMask ="99/99/00;"
                    EventProcPrefix ="RA_Received"

                    LayoutCachedLeft =7725
                    LayoutCachedWidth =8985
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1890
                    Height =270
                    TabIndex =1
                    Name ="Type"
                    ControlSource ="Type"

                    LayoutCachedLeft =1890
                    LayoutCachedWidth =3330
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10185
                    Width =1065
                    Height =270
                    ColumnWidth =1950
                    TabIndex =5
                    Name ="Text21"
                    ControlSource ="Inspected"
                    Format ="mm/dd/yy"
                    InputMask ="99/99/00;"

                    LayoutCachedLeft =10185
                    LayoutCachedWidth =11250
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12369
                    Width =990
                    Height =270
                    ColumnWidth =1890
                    TabIndex =6
                    Name ="Text25"
                    ControlSource ="RepairMade"
                    InputMask ="99/99/00;"

                    LayoutCachedLeft =12369
                    LayoutCachedWidth =13359
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =13419
                    Width =2370
                    Height =270
                    ColumnWidth =2250
                    TabIndex =7
                    Name ="Text27"
                    ControlSource ="contactlName"

                    LayoutCachedLeft =13419
                    LayoutCachedWidth =15789
                    LayoutCachedHeight =270
                End
                Begin CheckBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =11619
                    Top =45
                    Width =360
                    Height =270
                    TabIndex =8
                    Name ="Check29"
                    ControlSource ="Repair Needed (Y/N)"

                    LayoutCachedLeft =11619
                    LayoutCachedTop =45
                    LayoutCachedWidth =11979
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    SpecialEffect =2
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =13950
                    Width =45
                    Height =270
                    TabIndex =9
                    Name ="UnitID"
                    ControlSource ="UnitID"

                    LayoutCachedLeft =13950
                    LayoutCachedWidth =13995
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9120
                    Width =1020
                    Height =270
                    TabIndex =10
                    Name ="RCSent"
                    ControlSource ="RCSent"
                    Format ="mm/dd/yy"
                    StatusBarText ="per unit"

                    LayoutCachedLeft =9120
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =270
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
' See "UnitRcallsub.cls"
