Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =2
    ViewsAllowed =1
    TabularCharSet =238
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13200
    DatasheetFontHeight =11
    ItemSuffix =46
    Left =6975
    Top =7290
    Right =23850
    Bottom =14790
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x6c01b418d003e540
    End
    RecordSource ="SELECT Lines.Linebank, Lines.Lineid, Lines.LineDt, Lines.LineType, Lines.LineMin"
        "Loan, [lineamt]+[lineinc1]+[lineinc2]+[lineinc3]+[lineinc4]+[lineinc5] AS lineto"
        "t, IIf(IsNull([SumOfPNAmt]),0,[SumOfPNAmt]) AS Use FROM Lines LEFT JOIN PntSumq "
        "ON Lines.Lineid = PntSumq.Lineid WHERE (((Lines.Linebank)=[forms]![bank_view]![c"
        "lientid]));"
    Caption ="LeaseList"
    DatasheetFontName ="Calibri"
    Moveable =0
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =480
            BackColor =15707392
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =11490
                    Top =60
                    Width =75
                    Height =315
                    Name ="Client"
                    ControlSource ="Linebank"
                    StatusBarText ="link to client file"

                    LayoutCachedLeft =11490
                    LayoutCachedTop =60
                    LayoutCachedWidth =11565
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =2
                    Left =1455
                    Top =60
                    Width =1485
                    Height =315
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label25"
                    Caption ="Line Dt:"
                    LayoutCachedLeft =1455
                    LayoutCachedTop =60
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =163
                    Left =3450
                    Top =60
                    Width =1440
                    Height =315
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label26"
                    Caption ="Line Type:"
                    LayoutCachedLeft =3450
                    LayoutCachedTop =60
                    LayoutCachedWidth =4890
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =3
                    Left =5055
                    Top =60
                    Width =1965
                    Height =315
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label27"
                    Caption ="Line Amt:"
                    LayoutCachedLeft =5055
                    LayoutCachedTop =60
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =3
                    Left =7245
                    Top =60
                    Width =1965
                    Height =315
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label37"
                    Caption ="Use:"
                    LayoutCachedLeft =7245
                    LayoutCachedTop =60
                    LayoutCachedWidth =9210
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =163
                    TextAlign =3
                    Left =9240
                    Top =60
                    Width =1965
                    Height =315
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label38"
                    Caption ="Available:"
                    LayoutCachedLeft =9240
                    LayoutCachedTop =60
                    LayoutCachedWidth =11205
                    LayoutCachedHeight =375
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =11325
                    Top =90
                    Width =900
                    Height =285
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    ForeColor =-2147483615
                    Name ="addlease"
                    Caption ="Add Note"
                    FontName ="Tahoma"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =12
                        Begin
                            Action ="OpenForm"
                            Argument ="LinePromNTf"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="0"
                            Argument ="0"
                        End
                        Begin
                            Action ="SetValue"
                            Argument ="[Forms]![LinePromNTf]![lineid]"
                            Argument ="[Forms]![bank_view]![Linesum].[Form]![Lineid]"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"addlease\" xmlns=\"http://schemas.microsoft.com/office/acces"
                                "sservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument "
                                "Name=\"FormName\">LinePromNTf</Ar"
                        End
                        Begin
                            Comment ="_AXL:gument><Argument Name=\"DataMode\">Add</Argument></Action><Action Name=\"Se"
                                "tValue\"><Argument Name=\"Item\">[Forms]![LinePromNTf]![lineid]</Argument><Argum"
                                "ent Name=\"Expression\">[Forms]![bank_view]![Linesum].[Form]![Lineid]</Argument>"
                                "</Action></Statements"
                        End
                        Begin
                            Comment ="_AXL:></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11325
                    LayoutCachedTop =90
                    LayoutCachedWidth =12225
                    LayoutCachedHeight =375
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12300
                    Top =90
                    Width =900
                    Height =285
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    ForeColor =-2147483615
                    Name ="PronNTSUBREFBTN"
                    Caption ="Refresh"
                    FontName ="Tahoma"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="RunMacro"
                            Argument ="Refresh"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"PronNTSUBREFBTN\" xmlns=\"http://schemas.microsoft.com/offic"
                                "e/accessservices/2009/11/application\"><Statements><Action Name=\"RunMacro\"><Ar"
                                "gument Name=\"MacroName\">Refresh"
                        End
                        Begin
                            Comment ="_AXL:</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12300
                    LayoutCachedTop =90
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =375
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =3750
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =163
                    IMESentenceMode =3
                    Width =1065
                    Height =300
                    FontWeight =700
                    Name ="Lineid"
                    ControlSource ="Lineid"

                    LayoutCachedWidth =1065
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =163
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1380
                    Height =315
                    FontWeight =700
                    TabIndex =1
                    Name ="LineDt"
                    ControlSource ="LineDt"
                    Format ="mm/dd/yy"
                    StatusBarText ="document date"

                    LayoutCachedLeft =1380
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =163
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3450
                    Height =315
                    FontWeight =700
                    TabIndex =2
                    Name ="LineType"
                    ControlSource ="LineType"

                    LayoutCachedLeft =3450
                    LayoutCachedWidth =4890
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =163
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5055
                    Width =1965
                    Height =315
                    ColumnWidth =1620
                    FontWeight =700
                    TabIndex =3
                    Name ="LineAmt"
                    ControlSource ="linetot"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =5055
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =163
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7320
                    Width =1965
                    Height =315
                    FontWeight =700
                    TabIndex =4
                    Name ="Use"
                    ControlSource ="Use"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =7320
                    LayoutCachedWidth =9285
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =163
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9255
                    Width =1965
                    Height =315
                    FontWeight =700
                    TabIndex =5
                    Name ="Text39"
                    ControlSource ="=[linetot]-[use]"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =9255
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =315
                End
                Begin Subform
                    CanGrow = NotDefault
                    OverlapFlags =215
                    SpecialEffect =2
                    Top =540
                    Width =12645
                    Height =3075
                    TabIndex =6
                    Name ="PromNtSubf"
                    SourceObject ="Form.PromNtSubf"
                    LinkChildFields ="LineID"
                    LinkMasterFields ="Lineid"

                    LayoutCachedTop =540
                    LayoutCachedWidth =12645
                    LayoutCachedHeight =3615
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =95
                            Top =300
                            Width =1215
                            Height =315
                            Name ="PromNtSubf Label"
                            Caption ="PromNtSubf"
                            EventProcPrefix ="PromNtSubf_Label"
                            LayoutCachedTop =300
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =615
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
