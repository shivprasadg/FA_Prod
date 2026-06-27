Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14865
    DatasheetFontHeight =11
    ItemSuffix =81
    Left =3570
    Top =6720
    Right =18915
    Bottom =12570
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x881c5c2f2a3de440
    End
    RecordSource ="Lines"
    Caption ="Line Sub Form"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
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
        Begin Tab
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =555
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =10980
                    Width =1776
                    Height =381
                    FontSize =9
                    FontWeight =700
                    ForeColor =-2147483615
                    Name ="Command417"
                    Caption ="Add Interim Line"
                    FontName ="Tahoma"
                    ControlTipText ="Open Form"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =12
                        Begin
                            Action ="OpenForm"
                            Argument ="linesubform"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Action ="SetValue"
                            Argument ="[Forms]![linesubform]![linebank]"
                            Argument ="[Forms]![bank_view]![clientID]"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command417\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argumen"
                                "t Name=\"FormName\">linesubform</"
                        End
                        Begin
                            Comment ="_AXL:Argument></Action><Action Name=\"SetValue\"><Argument Name=\"Item\">[Forms]"
                                "![linesubform]![linebank]</Argument><Argument Name=\"Expression\">[Forms]![bank_"
                                "view]![clientID]</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10980
                    LayoutCachedWidth =12756
                    LayoutCachedHeight =381
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13110
                    Width =870
                    Height =381
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =-2147483615
                    Name ="LineRefBtn"
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
                                "nterfaceMacro For=\"LineRefBtn\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"RunMacro\"><Argumen"
                                "t Name=\"MacroName\">Refresh</Arg"
                        End
                        Begin
                            Comment ="_AXL:ument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =13110
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =381
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =4560
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =14010
                    Top =3975
                    Width =615
                    Height =315
                    FontSize =9
                    TabIndex =27
                    ForeColor =5026082
                    Name ="Lineid"
                    ControlSource ="Lineid"
                    StatusBarText ="line id"

                    LayoutCachedLeft =14010
                    LayoutCachedTop =3975
                    LayoutCachedWidth =14625
                    LayoutCachedHeight =4290
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =930
                    Top =225
                    Height =315
                    Name ="LineDt"
                    ControlSource ="LineDt"
                    Format ="mm/dd/yy"
                    StatusBarText ="document date"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Line dt is required"

                    LayoutCachedLeft =930
                    LayoutCachedTop =225
                    LayoutCachedWidth =2370
                    LayoutCachedHeight =540
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            Top =225
                            Width =780
                            Height =315
                            FontWeight =700
                            Name ="Label2"
                            Caption ="Line Dt:"
                            LayoutCachedTop =225
                            LayoutCachedWidth =780
                            LayoutCachedHeight =540
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7035
                    Top =225
                    Width =1620
                    Height =315
                    ColumnWidth =1620
                    TabIndex =2
                    Name ="LineAmt"
                    ControlSource ="LineAmt"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">0"
                    ValidationText ="Original Amount must be greater than 0"

                    LayoutCachedLeft =7035
                    LayoutCachedTop =225
                    LayoutCachedWidth =8655
                    LayoutCachedHeight =540
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            Left =5310
                            Top =225
                            Width =1665
                            Height =315
                            FontWeight =700
                            Name ="Label4"
                            Caption ="Original Amount:"
                            LayoutCachedLeft =5310
                            LayoutCachedTop =225
                            LayoutCachedWidth =6975
                            LayoutCachedHeight =540
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11625
                    Top =225
                    Height =315
                    TabIndex =3
                    Name ="LineExp"
                    ControlSource ="LineExp"
                    Format ="mm/dd/yy"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Original Expiry Required"

                    LayoutCachedLeft =11625
                    LayoutCachedTop =225
                    LayoutCachedWidth =13065
                    LayoutCachedHeight =540
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            Left =10110
                            Top =225
                            Width =1485
                            Height =315
                            FontWeight =700
                            Name ="Label5"
                            Caption ="Original Expiry:"
                            LayoutCachedLeft =10110
                            LayoutCachedTop =225
                            LayoutCachedWidth =11595
                            LayoutCachedHeight =540
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11655
                    Top =1470
                    Width =1800
                    Height =345
                    TabIndex =6
                    Name ="LineReNeg"
                    ControlSource ="LineReNeg"
                    Format ="mm/d/yy"
                    StatusBarText ="date can renegotiate terms"

                    LayoutCachedLeft =11655
                    LayoutCachedTop =1470
                    LayoutCachedWidth =13455
                    LayoutCachedHeight =1815
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9615
                            Top =1470
                            Width =1980
                            Height =345
                            Name ="Label6"
                            Caption ="Regnegotiate:"
                            LayoutCachedLeft =9615
                            LayoutCachedTop =1470
                            LayoutCachedWidth =11595
                            LayoutCachedHeight =1815
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11655
                    Top =1065
                    Width =1800
                    Height =315
                    TabIndex =5
                    Name ="LineMaxUsed"
                    ControlSource ="LineMaxUsed"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="max amount for used equipment under schedule or not"

                    LayoutCachedLeft =11655
                    LayoutCachedTop =1065
                    LayoutCachedWidth =13455
                    LayoutCachedHeight =1380
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9255
                            Top =1065
                            Width =2340
                            Height =315
                            Name ="Label7"
                            Caption ="Total Used Equip. Notes:"
                            LayoutCachedLeft =9255
                            LayoutCachedTop =1065
                            LayoutCachedWidth =11595
                            LayoutCachedHeight =1380
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11655
                    Top =645
                    Width =1800
                    Height =315
                    TabIndex =4
                    Name ="LineMinLoan"
                    ControlSource ="LineMinLoan"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="minimum amount of loan or drawn"
                    ValidationRule =">=0"
                    ValidationText ="Minimum Note must be greater than or equal to 0"

                    LayoutCachedLeft =11655
                    LayoutCachedTop =645
                    LayoutCachedWidth =13455
                    LayoutCachedHeight =960
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9465
                            Top =645
                            Width =2130
                            Height =315
                            Name ="Label10"
                            Caption ="Minimum Note :"
                            LayoutCachedLeft =9465
                            LayoutCachedTop =645
                            LayoutCachedWidth =11595
                            LayoutCachedHeight =960
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4155
                    Top =2280
                    Width =1620
                    Height =315
                    TabIndex =9
                    Name ="LineInc1"
                    ControlSource ="LineInc1"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Amdt1 amount"
                    ValidationRule =">=0"
                    ValidationText ="Line 5 increase must be >=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =4155
                    LayoutCachedTop =2280
                    LayoutCachedWidth =5775
                    LayoutCachedHeight =2595
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =195
                            Top =2265
                            Width =765
                            Height =315
                            FontWeight =700
                            Name ="lin1lbl"
                            Caption ="1:"
                            LayoutCachedLeft =195
                            LayoutCachedTop =2265
                            LayoutCachedWidth =960
                            LayoutCachedHeight =2580
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1050
                    Top =2265
                    Width =1245
                    Height =315
                    TabIndex =7
                    Name ="Line1IncDT"
                    ControlSource ="Line1IncDT"
                    Format ="mm/dd/yy"
                    StatusBarText ="Amdt1 dt"

                    LayoutCachedLeft =1050
                    LayoutCachedTop =2265
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =2580
                    Begin
                        Begin Label
                            FontUnderline = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =2
                            Left =1050
                            Top =1800
                            Width =1245
                            Height =315
                            FontWeight =700
                            Name ="/rvDtlbl"
                            Caption ="Date:"
                            EventProcPrefix ="Ctl_rvDtlbl"
                            LayoutCachedLeft =1050
                            LayoutCachedTop =1800
                            LayoutCachedWidth =2295
                            LayoutCachedHeight =2115
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2415
                    Top =2250
                    Height =315
                    TabIndex =8
                    Name ="LineExp1"
                    ControlSource ="LineExp1"
                    Format ="mm/dd/yy"
                    StatusBarText ="Amdt1 expiry"

                    LayoutCachedLeft =2415
                    LayoutCachedTop =2250
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =2565
                    Begin
                        Begin Label
                            FontUnderline = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =2
                            Left =2400
                            Top =1785
                            Width =1455
                            Height =315
                            FontWeight =700
                            Name ="revexplbl"
                            Caption ="Expiry:"
                            LayoutCachedLeft =2400
                            LayoutCachedTop =1785
                            LayoutCachedWidth =3855
                            LayoutCachedHeight =2100
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4155
                    Top =2670
                    Width =1620
                    Height =315
                    TabIndex =12
                    Name ="LineInc2"
                    ControlSource ="LineInc2"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Amdt2 amount"
                    ValidationRule =">=0"
                    ValidationText ="Line2 increase must be >=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =4155
                    LayoutCachedTop =2670
                    LayoutCachedWidth =5775
                    LayoutCachedHeight =2985
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =195
                            Top =2685
                            Width =765
                            Height =315
                            FontWeight =700
                            Name ="line2lbl"
                            Caption ="2:"
                            LayoutCachedLeft =195
                            LayoutCachedTop =2685
                            LayoutCachedWidth =960
                            LayoutCachedHeight =3000
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1050
                    Top =2730
                    Width =1245
                    Height =315
                    TabIndex =10
                    Name ="Line2IncDT"
                    ControlSource ="Line2IncDT"
                    Format ="mm/dd/yy"
                    StatusBarText ="Amdt2 dt"

                    LayoutCachedLeft =1050
                    LayoutCachedTop =2730
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =3045
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2415
                    Top =2715
                    Height =315
                    TabIndex =11
                    Name ="LineExp2"
                    ControlSource ="LineExp2"
                    Format ="mm/dd/yy"
                    StatusBarText ="Amdt2 expiry"

                    LayoutCachedLeft =2415
                    LayoutCachedTop =2715
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =3030
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4155
                    Top =3120
                    Width =1620
                    Height =315
                    TabIndex =15
                    Name ="LineInc3"
                    ControlSource ="LineInc3"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Amdt3 dt"
                    ValidationRule =">=0"
                    ValidationText ="Line 3 increase must be >=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =4155
                    LayoutCachedTop =3120
                    LayoutCachedWidth =5775
                    LayoutCachedHeight =3435
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =195
                            Top =3120
                            Width =765
                            Height =315
                            FontWeight =700
                            Name ="liln3lbl"
                            Caption ="3:"
                            LayoutCachedLeft =195
                            LayoutCachedTop =3120
                            LayoutCachedWidth =960
                            LayoutCachedHeight =3435
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2415
                    Top =3150
                    Height =315
                    TabIndex =14
                    Name ="LineExp3"
                    ControlSource ="LineExp3"
                    Format ="mm/dd/yy"
                    StatusBarText ="amdt3 dt"

                    LayoutCachedLeft =2415
                    LayoutCachedTop =3150
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =3465
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1050
                    Top =3165
                    Width =1245
                    Height =315
                    TabIndex =13
                    Name ="Line3IncDT"
                    ControlSource ="Line3IncDT"
                    Format ="mm/dd/yy"
                    StatusBarText ="amdt3 amount"

                    LayoutCachedLeft =1050
                    LayoutCachedTop =3165
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =3480
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =13035
                    Top =3975
                    Width =645
                    Height =315
                    FontSize =9
                    TabIndex =28
                    ForeColor =5026082
                    Name ="Linebank"
                    ControlSource ="Linebank"
                    StatusBarText ="Bank record in C/P table"

                    LayoutCachedLeft =13035
                    LayoutCachedTop =3975
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =4290
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    Left =4155
                    Top =1800
                    Width =1485
                    Height =315
                    FontWeight =700
                    Name ="revinclbl"
                    Caption ="$ Increase:"
                    LayoutCachedLeft =4155
                    LayoutCachedTop =1800
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =2115
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6915
                    Top =2310
                    Width =1755
                    Height =315
                    TabIndex =22
                    Name ="Curr1Amt"
                    ControlSource ="=[lineamt]+[lineinc1]"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =6915
                    LayoutCachedTop =2310
                    LayoutCachedWidth =8670
                    LayoutCachedHeight =2625
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    Left =6915
                    Top =1815
                    Width =1755
                    Height =315
                    FontWeight =700
                    Name ="currAmtlbl"
                    Caption ="Current Line Amt"
                    LayoutCachedLeft =6915
                    LayoutCachedTop =1815
                    LayoutCachedWidth =8670
                    LayoutCachedHeight =2130
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6915
                    Top =2670
                    Width =1755
                    Height =315
                    TabIndex =23
                    Name ="CurrAmt2"
                    ControlSource ="=IIf([lineinc2]=0 And IsNull([line2incdt]),Null,[lineamt]+[lineinc1]+[lineinc2])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =6915
                    LayoutCachedTop =2670
                    LayoutCachedWidth =8670
                    LayoutCachedHeight =2985
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6915
                    Top =3105
                    Width =1755
                    Height =315
                    TabIndex =24
                    Name ="CurrAmt3"
                    ControlSource ="=IIf([lineinc3]=0 And IsNull([line3incdt]),Null,[lineamt]+[lineinc1]+[lineinc2]+"
                        "[lineinc3])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =6915
                    LayoutCachedTop =3105
                    LayoutCachedWidth =8670
                    LayoutCachedHeight =3420
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3735
                    Top =225
                    Height =330
                    TabIndex =1
                    Name ="LineType"
                    ControlSource ="LineType"
                    RowSourceType ="Value List"
                    RowSource ="\"Leased\";\"Rmkt\";\"Both\""
                    ColumnWidths ="7200"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Line Type must be selected"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3735
                    LayoutCachedTop =225
                    LayoutCachedWidth =5175
                    LayoutCachedHeight =555
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            Left =2655
                            Top =225
                            Width =1020
                            Height =315
                            FontWeight =700
                            Name ="Label3"
                            Caption ="Line Type:"
                            LayoutCachedLeft =2655
                            LayoutCachedTop =225
                            LayoutCachedWidth =3675
                            LayoutCachedHeight =540
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4155
                    Top =3615
                    Width =1620
                    Height =315
                    TabIndex =18
                    Name ="LineInc4"
                    ControlSource ="LineInc4"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Amdt3 dt"
                    ValidationRule =">=0"
                    ValidationText ="Line 4 increase must be >=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =4155
                    LayoutCachedTop =3615
                    LayoutCachedWidth =5775
                    LayoutCachedHeight =3930
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =195
                            Top =3615
                            Width =765
                            Height =315
                            FontWeight =700
                            Name ="line4lbl"
                            Caption ="4:"
                            LayoutCachedLeft =195
                            LayoutCachedTop =3615
                            LayoutCachedWidth =960
                            LayoutCachedHeight =3930
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2415
                    Top =3645
                    Height =315
                    TabIndex =17
                    Name ="LineExp4"
                    ControlSource ="LineExp4"
                    Format ="mm/dd/yy"
                    StatusBarText ="amdt3 dt"

                    LayoutCachedLeft =2415
                    LayoutCachedTop =3645
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =3960
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1050
                    Top =3660
                    Width =1245
                    Height =315
                    TabIndex =16
                    Name ="Line4IncDT"
                    ControlSource ="Line4IncDT"
                    Format ="mm/dd/yy"
                    StatusBarText ="amdt3 amount"

                    LayoutCachedLeft =1050
                    LayoutCachedTop =3660
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =3975
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6915
                    Top =3600
                    Width =1755
                    Height =315
                    TabIndex =25
                    Name ="CurrAmt4"
                    ControlSource ="=IIf([lineinc4]=0 And IsNull([line4incdt]),Null,[lineamt]+[lineinc1]+[lineinc2]+"
                        "[lineinc3]+[lineinc4])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =6915
                    LayoutCachedTop =3600
                    LayoutCachedWidth =8670
                    LayoutCachedHeight =3915
                End
                Begin TextBox
                    DecimalPlaces =2
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4155
                    Top =4065
                    Width =1620
                    Height =315
                    TabIndex =21
                    Name ="LineInc5"
                    ControlSource ="LineInc5"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Amdt3 dt"
                    ValidationRule =">=0"
                    ValidationText ="Line 5 increase must be >=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =4155
                    LayoutCachedTop =4065
                    LayoutCachedWidth =5775
                    LayoutCachedHeight =4380
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =195
                            Top =4065
                            Width =765
                            Height =315
                            FontWeight =700
                            Name ="lin5lbl"
                            Caption ="5:"
                            LayoutCachedLeft =195
                            LayoutCachedTop =4065
                            LayoutCachedWidth =960
                            LayoutCachedHeight =4380
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2400
                    Top =4230
                    Height =315
                    TabIndex =20
                    Name ="LineExp5"
                    ControlSource ="LineExp5"
                    Format ="mm/dd/yy"
                    StatusBarText ="amdt3 dt"

                    LayoutCachedLeft =2400
                    LayoutCachedTop =4230
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =4545
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1035
                    Top =4245
                    Width =1245
                    Height =315
                    TabIndex =19
                    Name ="LineIncDT"
                    ControlSource ="Line5IncDT"
                    Format ="mm/dd/yy"
                    StatusBarText ="amdt3 amount"

                    LayoutCachedLeft =1035
                    LayoutCachedTop =4245
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =4560
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6915
                    Top =4035
                    Width =1755
                    Height =315
                    TabIndex =26
                    Name ="CurrAmt5"
                    ControlSource ="=IIf([lineinc5]=0 And IsNull([line5incdt]),Null,[lineamt]+[lineinc1]+[lineinc2]+"
                        "[lineinc3]+[lineinc4]+[lineinc5])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    DefaultValue ="0"

                    LayoutCachedLeft =6915
                    LayoutCachedTop =4035
                    LayoutCachedWidth =8670
                    LayoutCachedHeight =4350
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    Top =1800
                    Width =975
                    Height =315
                    FontWeight =700
                    Name ="Label58"
                    Caption ="Rev/Amd:"
                    LayoutCachedTop =1800
                    LayoutCachedWidth =975
                    LayoutCachedHeight =2115
                End
            End
        End
        Begin FormFooter
            Height =15
            Name ="FormFooter"
        End
    End
End
