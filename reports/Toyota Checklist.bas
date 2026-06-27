Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15105
    DatasheetFontHeight =11
    ItemSuffix =88
    Left =3885
    Top =2775
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x388fc420f90ee440
    End
    RecordSource ="SELECT SchCkListQ2.*\015\012FROM SchCkListQ2\015\012WHERE ([SchID]=493);"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
        End
        Begin ListBox
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="MLNo"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="SchNo"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="SGrpNO"
        End
        Begin BreakLevel
            ControlSource ="UnitUnitNum"
        End
        Begin BreakLevel
            ControlSource ="UnitType"
        End
        Begin BreakLevel
            ControlSource ="VIN4"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="UnitID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =315
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Width =8160
                    Height =315
                    FontSize =12
                    FontWeight =700
                    Name ="Text1"
                    ControlSource ="=UCase([CLIENTSHNM] & \":  \" & [MLNO] & \"  SCHEDULE \" & [SCHNO] & \"  CHECKLI"
                        "ST\")"

                    LayoutCachedWidth =8160
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =8700
                    Height =315
                    TabIndex =1
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =8700
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =315
                End
            End
        End
        Begin PageHeader
            Height =810
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =915
                    Top =120
                    Width =1740
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label11"
                    Caption ="VIN:"
                    LayoutCachedLeft =915
                    LayoutCachedTop =120
                    LayoutCachedWidth =2655
                    LayoutCachedHeight =360
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =5370
                    Top =90
                    Width =1080
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label13"
                    Caption ="MSO Rec'd:"
                    LayoutCachedLeft =5370
                    LayoutCachedTop =90
                    LayoutCachedWidth =6450
                    LayoutCachedHeight =330
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =6450
                    Top =90
                    Width =1455
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label14"
                    Caption ="Invoice #:"
                    LayoutCachedLeft =6450
                    LayoutCachedTop =90
                    LayoutCachedWidth =7905
                    LayoutCachedHeight =330
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =7920
                    Top =90
                    Width =1140
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label15"
                    Caption ="FinalCost:"
                    LayoutCachedLeft =7920
                    LayoutCachedTop =90
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =330
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =9105
                    Top =90
                    Width =1200
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label18"
                    Caption ="Paid:"
                    LayoutCachedLeft =9105
                    LayoutCachedTop =90
                    LayoutCachedWidth =10305
                    LayoutCachedHeight =330
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =10305
                    Top =90
                    Width =1020
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label19"
                    Caption ="Accepted Dt:"
                    LayoutCachedLeft =10305
                    LayoutCachedTop =90
                    LayoutCachedWidth =11325
                    LayoutCachedHeight =330
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =11325
                    Top =90
                    Width =1005
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label20"
                    Caption ="COA Back:"
                    LayoutCachedLeft =11325
                    LayoutCachedTop =90
                    LayoutCachedWidth =12330
                    LayoutCachedHeight =330
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextFontFamily =34
                    Left =12360
                    Top =90
                    Width =1080
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label25"
                    Caption ="Title Rec'd:"
                    LayoutCachedLeft =12360
                    LayoutCachedTop =90
                    LayoutCachedWidth =13440
                    LayoutCachedHeight =330
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextFontFamily =34
                    Left =13800
                    Top =105
                    Width =855
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label35"
                    Caption ="Dom:"
                    LayoutCachedLeft =13800
                    LayoutCachedTop =105
                    LayoutCachedWidth =14655
                    LayoutCachedHeight =345
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =75
                    Width =915
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label85"
                    Caption ="Unit No.:"
                    LayoutCachedLeft =75
                    LayoutCachedWidth =990
                    LayoutCachedHeight =240
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =300
            BackColor =14545371
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Top =240
                    Width =4380
                    Height =60
                    FontSize =9
                    Name ="Text5"
                    ControlSource ="=IIf(IsNull([MLGUARANTOR]),Null,IIf(IsNull([MLGUARBACK]),\"GUARANTY O/S: SENT: \""
                        " & [MLGUARSENT],\"GUARANTY BACK\"))"

                    LayoutCachedTop =240
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =2880
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="Text3"
                    ControlSource ="=IIf(IsNull([MLBACK]),\"MASTER LEASE O/S: SENT: \" & [MLBACK],\"MASTER LEASE BAC"
                        "K\")"

                    LayoutCachedWidth =2880
                    LayoutCachedHeight =240
                End
                Begin CheckBox
                    Left =2880
                    Top =30
                    TabIndex =2
                    Name ="MLINSBLANKET"
                    ControlSource ="MLINSBLANKET"
                    StatusBarText ="IF YES AT LEASE LEVEL IS NO AT SCHEDULE LEVEL"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =30
                    LayoutCachedWidth =3140
                    LayoutCachedHeight =270
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =3105
                            Width =1455
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label7"
                            Caption ="Blanket Insurance"
                            LayoutCachedLeft =3105
                            LayoutCachedWidth =4560
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4620
                    Width =780
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="Text8"
                    ControlSource ="=IIf([MLinsblanket]=No,Null,IIf([mlinsExp]<Date(),\"EXPIRED\",\"Expires\"))"

                    LayoutCachedLeft =4620
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5400
                    Width =900
                    FontSize =9
                    TabIndex =4
                    Name ="MLInsExp"
                    ControlSource ="MLInsExp"
                    Format ="mm/dd/yy"
                    ShowDatePicker =1

                    LayoutCachedLeft =5400
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =240
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =240
            BreakLevel =1
            BackColor =14545371
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =3240
                    FontSize =9
                    FontWeight =700
                    Name ="Text6"
                    ControlSource ="=IIf(IsNull([SCHIN]),\"Schedule O/S: SENT: \" & [SCHOUT],\"Schedule Back\")"
                    Format =">"

                    LayoutCachedWidth =3240
                    LayoutCachedHeight =240
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =525
            BreakLevel =2
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    TextFontCharSet =163
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =900
                    Width =1320
                    FontSize =9
                    FontWeight =700
                    Name ="SGrpNO"
                    ControlSource ="SGrpNO"

                    LayoutCachedLeft =900
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =240
                    Begin
                        Begin Label
                            TextFontCharSet =163
                            TextFontFamily =34
                            Width =855
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label10"
                            Caption ="Grp No:"
                            LayoutCachedWidth =855
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3300
                    Width =480
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="UnitsInGroup\015\012\015\012\015\012\015\012"
                    ControlSource ="UnitsInGroup\015\012\015\012\015\012\015\012"
                    EventProcPrefix ="UnitsInGroup________"

                    LayoutCachedLeft =3300
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =240
                    Begin
                        Begin Label
                            TextFontCharSet =163
                            TextFontFamily =34
                            Left =2280
                            Width =960
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label30"
                            Caption ="No Units:"
                            LayoutCachedLeft =2280
                            LayoutCachedWidth =3240
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7725
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="OrigEquipCost"
                    ControlSource ="OrigEquipCost"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =7725
                    LayoutCachedWidth =9165
                    LayoutCachedHeight =240
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontCharSet =163
                            TextAlign =3
                            TextFontFamily =34
                            Left =5850
                            Width =1485
                            Height =240
                            FontSize =10
                            FontWeight =700
                            Name ="Label31"
                            Caption =" Total Equip Cost:"
                            LayoutCachedLeft =5850
                            LayoutCachedWidth =7335
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7650
                    Top =285
                    FontSize =9
                    TabIndex =3
                    Name ="Text32"
                    ControlSource ="=Sum([unitfinalcost])+IIf(IsNull(Sum([SUMOFchildfinalcost])),0,Sum([SUMOFchildfi"
                        "nalcost]))"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =7650
                    LayoutCachedTop =285
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =525
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =1
                            TextFontCharSet =163
                            TextAlign =3
                            TextFontFamily =34
                            Left =6555
                            Top =285
                            Width =840
                            Height =240
                            FontSize =10
                            Name ="Label33"
                            Caption ="Check:"
                            LayoutCachedLeft =6555
                            LayoutCachedTop =285
                            LayoutCachedWidth =7395
                            LayoutCachedHeight =525
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3285
                    Top =225
                    Width =480
                    FontSize =10
                    TabIndex =4
                    Name ="Text37"
                    ControlSource ="=Count([unitid])/2"

                    LayoutCachedLeft =3285
                    LayoutCachedTop =225
                    LayoutCachedWidth =3765
                    LayoutCachedHeight =465
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextFontFamily =34
                            Left =2475
                            Top =225
                            Width =750
                            Height =240
                            FontSize =10
                            Name ="Label38"
                            Caption ="Check:"
                            LayoutCachedLeft =2475
                            LayoutCachedTop =225
                            LayoutCachedWidth =3225
                            LayoutCachedHeight =465
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =315
            BreakLevel =6
            BackColor =14545371
            Name ="GroupHeader2"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1440
                    Top =45
                    Width =1680
                    FontSize =9
                    Name ="Text66"
                    ControlSource ="UnitVIN"

                    LayoutCachedLeft =1440
                    LayoutCachedTop =45
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5745
                    Top =75
                    Width =1005
                    FontSize =9
                    TabIndex =1
                    Name ="Text67"
                    ControlSource ="UnitMSORecd"
                    ShowDatePicker =1

                    LayoutCachedLeft =5745
                    LayoutCachedTop =75
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6765
                    Top =75
                    Width =1320
                    FontSize =9
                    TabIndex =2
                    Name ="Text68"
                    ControlSource ="UnitInvoiceNum"

                    LayoutCachedLeft =6765
                    LayoutCachedTop =75
                    LayoutCachedWidth =8085
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3135
                    Top =75
                    Width =2610
                    FontSize =9
                    TabIndex =3
                    Name ="Text69"
                    ControlSource ="UDESC"

                    LayoutCachedLeft =3135
                    LayoutCachedTop =75
                    LayoutCachedWidth =5745
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9435
                    Top =75
                    Width =1080
                    FontSize =9
                    TabIndex =4
                    Name ="Text70"
                    ControlSource ="UnitVendorPytDate"
                    Format ="mm/dd/yy"
                    ShowDatePicker =1

                    LayoutCachedLeft =9435
                    LayoutCachedTop =75
                    LayoutCachedWidth =10515
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10515
                    Top =75
                    Width =1080
                    FontSize =9
                    TabIndex =5
                    Name ="Text71"
                    ControlSource ="UnitAcceptDate"
                    Format ="mm/dd/yy"
                    ShowDatePicker =1

                    LayoutCachedLeft =10515
                    LayoutCachedTop =75
                    LayoutCachedWidth =11595
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11595
                    Top =75
                    Width =1080
                    FontSize =9
                    TabIndex =6
                    Name ="Text72"
                    ControlSource ="UnitCOABack"
                    Format ="mm/dd/yy"
                    ShowDatePicker =1

                    LayoutCachedLeft =11595
                    LayoutCachedTop =75
                    LayoutCachedWidth =12675
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12690
                    Top =75
                    Width =1035
                    FontSize =9
                    TabIndex =7
                    Name ="Text73"
                    ControlSource ="UnitTitleRecd"
                    Format ="mm/dd/yy"
                    ShowDatePicker =1

                    LayoutCachedLeft =12690
                    LayoutCachedTop =75
                    LayoutCachedWidth =13725
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13740
                    Top =75
                    Width =330
                    FontSize =9
                    TabIndex =8
                    Name ="Text74"
                    ControlSource ="LState"
                    StatusBarText ="lookup in lookup_state table"

                    LayoutCachedLeft =13740
                    LayoutCachedTop =75
                    LayoutCachedWidth =14070
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =14850
                    Top =75
                    Width =0
                    TabIndex =9
                    Name ="Text75"
                    ControlSource ="SumOfchildFinalCost"

                    LayoutCachedLeft =14850
                    LayoutCachedTop =75
                    LayoutCachedWidth =14850
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =2715
                    Top =75
                    Width =120
                    TabIndex =10
                    Name ="Text78"
                    ControlSource ="UnitID"

                    LayoutCachedLeft =2715
                    LayoutCachedTop =75
                    LayoutCachedWidth =2835
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8175
                    Top =75
                    Width =1260
                    FontSize =9
                    TabIndex =11
                    Name ="Text79"
                    ControlSource ="UnitFinalCost"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8175
                    LayoutCachedTop =75
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14070
                    Top =75
                    Width =300
                    FontSize =9
                    TabIndex =12
                    Name ="Text80"
                    ControlSource ="StTaxType"

                    LayoutCachedLeft =14070
                    LayoutCachedTop =75
                    LayoutCachedWidth =14370
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Top =75
                    Width =870
                    ColumnWidth =2205
                    FontSize =9
                    TabIndex =13
                    Name ="UnitUnitNum"
                    ControlSource ="UnitUnitNum"

                    LayoutCachedTop =75
                    LayoutCachedWidth =870
                    LayoutCachedHeight =315
                End
                Begin Line
                    Width =14625
                    Height =15
                    Name ="Line47"
                    LayoutCachedWidth =14625
                    LayoutCachedHeight =15
                End
                Begin CheckBox
                    OldBorderStyle =0
                    Left =14430
                    Top =75
                    TabIndex =14
                    Name ="Check76"
                    ControlSource ="UnitTaxExempt"

                    LayoutCachedLeft =14430
                    LayoutCachedTop =75
                    LayoutCachedWidth =14690
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            TextAlign =1
                            TextFontFamily =34
                            Left =14700
                            Top =75
                            Width =405
                            Height =240
                            FontSize =9
                            Name ="Label77"
                            Caption =" Expt"
                            LayoutCachedLeft =14700
                            LayoutCachedTop =75
                            LayoutCachedWidth =15105
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =855
                    Top =60
                    Width =585
                    Height =225
                    FontSize =9
                    FontWeight =700
                    TabIndex =15
                    ForeColor =255
                    Name ="Text86"
                    ControlSource ="UnitID"

                    LayoutCachedLeft =855
                    LayoutCachedTop =60
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =285
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1425
            Name ="Detail"
            Begin
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =660
                    Top =45
                    Width =12990
                    Height =1380
                    Name ="SchChecklistSub"
                    SourceObject ="Report.SchChecklistSub"
                    LinkChildFields ="UnitREF"
                    LinkMasterFields ="UnitID"

                    LayoutCachedLeft =660
                    LayoutCachedTop =45
                    LayoutCachedWidth =13650
                    LayoutCachedHeight =1425
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =14100
                    Width =0
                    Height =60
                    TabIndex =1
                    Name ="Text60"
                    ControlSource ="SumOfchildFinalCost"

                    LayoutCachedLeft =14100
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =60
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Width =120
                    Height =120
                    TabIndex =2
                    Name ="Text63"
                    ControlSource ="UnitID"

                    LayoutCachedWidth =120
                    LayoutCachedHeight =120
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =14040
                    Width =60
                    Height =60
                    TabIndex =3
                    Name ="SumOfchildFinalCost"
                    ControlSource ="SumOfchildFinalCost"

                    LayoutCachedLeft =14040
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =60
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Width =120
                    Height =120
                    TabIndex =4
                    Name ="UnitID"
                    ControlSource ="UnitID"

                    LayoutCachedWidth =120
                    LayoutCachedHeight =120
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =300
            BreakLevel =6
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6915
                    Top =30
                    Width =2100
                    FontSize =9
                    FontWeight =700
                    Name ="TotChild"
                    ControlSource ="TotUnit"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =6915
                    LayoutCachedTop =30
                    LayoutCachedWidth =9015
                    LayoutCachedHeight =270
                End
            End
        End
        Begin PageFooter
            Height =240
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =2040
                    FontSize =9
                    Name ="Text44"
                    ControlSource ="=Date()"
                    Format ="mmm dd\", \"yyyy"

                    LayoutCachedWidth =2040
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11475
                    Width =3600
                    FontSize =9
                    TabIndex =1
                    Name ="Text50"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =11475
                    LayoutCachedWidth =15075
                    LayoutCachedHeight =240
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
