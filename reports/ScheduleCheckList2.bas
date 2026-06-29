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
    Width =15240
    DatasheetFontHeight =11
    ItemSuffix =96
    Left =11205
    Top =2850
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x8378e8147384e440
    End
    RecordSource ="SchCkListQ2"
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
            ControlSource ="VIN4"
        End
        Begin BreakLevel
            ControlSource ="UnitType"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="UnitID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =390
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
            Height =0
            Name ="PageHeaderSection"
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
            Height =924
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
                    Left =7530
                    Width =1560
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="OrigEquipCost"
                    ControlSource ="OrigEquipCost"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =7530
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =240
                    Begin
                        Begin Label
                            BackStyle =1
                            TextFontCharSet =163
                            TextAlign =3
                            TextFontFamily =34
                            Left =5250
                            Width =2220
                            Height =240
                            FontSize =10
                            FontWeight =700
                            Name ="Label31"
                            Caption =" Total Equip Cost:"
                            LayoutCachedLeft =5250
                            LayoutCachedWidth =7470
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
                    Left =7530
                    Top =285
                    Width =1560
                    FontSize =9
                    TabIndex =3
                    Name ="Text32"
                    ControlSource ="=Sum([unitfinalcost])+IIf(IsNull(Sum([SUMOFchildfinalcost])),0,Sum([SUMOFchildfi"
                        "nalcost]))"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =7530
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
                            Left =6330
                            Top =285
                            Width =1080
                            Height =240
                            FontSize =10
                            Name ="Label33"
                            Caption ="Check:"
                            LayoutCachedLeft =6330
                            LayoutCachedTop =285
                            LayoutCachedWidth =7410
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
                    ControlSource ="=Count([unitid])"

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
                Begin Line
                    LineSlant = NotDefault
                    Top =510
                    Width =14460
                    Height =15
                    Name ="Line47"
                    LayoutCachedTop =510
                    LayoutCachedWidth =14460
                    LayoutCachedHeight =525
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =744
                    Top =660
                    Width =1752
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label11"
                    Caption ="VIN:"
                    GroupTable =1
                    LayoutCachedLeft =744
                    LayoutCachedTop =660
                    LayoutCachedWidth =2496
                    LayoutCachedHeight =900
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =10068
                    Top =660
                    Width =900
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label13"
                    Caption ="MSO Rec'd:"
                    GroupTable =1
                    LayoutCachedLeft =10068
                    LayoutCachedTop =660
                    LayoutCachedWidth =10968
                    LayoutCachedHeight =900
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =6744
                    Top =660
                    Width =1212
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label14"
                    Caption ="Invoice #:"
                    GroupTable =1
                    LayoutCachedLeft =6744
                    LayoutCachedTop =660
                    LayoutCachedWidth =7956
                    LayoutCachedHeight =900
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =2556
                    Top =660
                    Width =696
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =255
                    Name ="Label86"
                    Caption ="Asset#:"
                    GroupTable =1
                    LayoutCachedLeft =2556
                    LayoutCachedTop =660
                    LayoutCachedWidth =3252
                    LayoutCachedHeight =900
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =3312
                    Top =660
                    Width =3372
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label93"
                    Caption ="Desc"
                    GroupTable =1
                    LayoutCachedLeft =3312
                    LayoutCachedTop =660
                    LayoutCachedWidth =6684
                    LayoutCachedHeight =900
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =8016
                    Top =660
                    Width =1116
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label15"
                    Caption ="Unit Cost:"
                    GroupTable =1
                    LayoutCachedLeft =8016
                    LayoutCachedTop =660
                    LayoutCachedWidth =9132
                    LayoutCachedHeight =900
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =9192
                    Top =660
                    Width =816
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label18"
                    Caption ="Paid:"
                    GroupTable =1
                    LayoutCachedLeft =9192
                    LayoutCachedTop =660
                    LayoutCachedWidth =10008
                    LayoutCachedHeight =900
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =11028
                    Top =660
                    Width =756
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label19"
                    Caption ="Accepted"
                    GroupTable =1
                    LayoutCachedLeft =11028
                    LayoutCachedTop =660
                    LayoutCachedWidth =11784
                    LayoutCachedHeight =900
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =11844
                    Top =660
                    Width =756
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label20"
                    Caption ="COA Back:"
                    GroupTable =1
                    LayoutCachedLeft =11844
                    LayoutCachedTop =660
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =900
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextFontFamily =34
                    Left =13776
                    Top =660
                    Width =300
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label35"
                    Caption ="Dom:"
                    GroupTable =1
                    LayoutCachedLeft =13776
                    LayoutCachedTop =660
                    LayoutCachedWidth =14076
                    LayoutCachedHeight =900
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =14136
                    Top =660
                    Width =540
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label87"
                    Caption ="Type:"
                    GroupTable =1
                    LayoutCachedLeft =14136
                    LayoutCachedTop =660
                    LayoutCachedWidth =14676
                    LayoutCachedHeight =900
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =12660
                    Top =660
                    Width =1056
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label94"
                    Caption ="Title Event"
                    GroupTable =1
                    LayoutCachedLeft =12660
                    LayoutCachedTop =660
                    LayoutCachedWidth =13716
                    LayoutCachedHeight =900
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =14736
                    Top =660
                    Width =360
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label95"
                    Caption ="Expt"
                    GroupTable =1
                    LayoutCachedLeft =14736
                    LayoutCachedTop =660
                    LayoutCachedWidth =15096
                    LayoutCachedHeight =900
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =36
                    Top =660
                    Width =648
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label89"
                    Caption ="Unit #"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =660
                    LayoutCachedWidth =684
                    LayoutCachedHeight =900
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =306
            BreakLevel =5
            BackColor =14545371
            Name ="GroupHeader2"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =744
                    Top =36
                    Width =1752
                    FontSize =9
                    TabIndex =1
                    Name ="Text66"
                    ControlSource ="UnitVIN"
                    GroupTable =1

                    LayoutCachedLeft =744
                    LayoutCachedTop =36
                    LayoutCachedWidth =2496
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10068
                    Top =36
                    Width =900
                    FontSize =9
                    TabIndex =8
                    Name ="Text67"
                    ControlSource ="UnitMSORecd"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    ShowDatePicker =1

                    LayoutCachedLeft =10068
                    LayoutCachedTop =36
                    LayoutCachedWidth =10968
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6744
                    Top =36
                    Width =1212
                    FontSize =9
                    TabIndex =4
                    Name ="Text68"
                    ControlSource ="UnitInvoiceNum"
                    GroupTable =1

                    LayoutCachedLeft =6744
                    LayoutCachedTop =36
                    LayoutCachedWidth =7956
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3312
                    Top =36
                    Width =3372
                    FontSize =9
                    TabIndex =3
                    Name ="Text69"
                    ControlSource ="UDESC"
                    GroupTable =1

                    LayoutCachedLeft =3312
                    LayoutCachedTop =36
                    LayoutCachedWidth =6684
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9192
                    Top =36
                    Width =816
                    FontSize =9
                    TabIndex =6
                    Name ="Text70"
                    ControlSource ="UnitVendorPytDate"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    ShowDatePicker =1

                    LayoutCachedLeft =9192
                    LayoutCachedTop =36
                    LayoutCachedWidth =10008
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11028
                    Top =36
                    Width =756
                    FontSize =9
                    TabIndex =7
                    Name ="Text71"
                    ControlSource ="UnitAcceptDate"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    ShowDatePicker =1

                    LayoutCachedLeft =11028
                    LayoutCachedTop =36
                    LayoutCachedWidth =11784
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11844
                    Top =36
                    Width =756
                    FontSize =9
                    TabIndex =9
                    Name ="Text72"
                    ControlSource ="UnitCOABack"
                    Format ="mm/dd/yy"
                    GroupTable =1
                    ShowDatePicker =1

                    LayoutCachedLeft =11844
                    LayoutCachedTop =36
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13776
                    Top =36
                    Width =300
                    FontSize =9
                    TabIndex =11
                    Name ="Text74"
                    ControlSource ="LState"
                    StatusBarText ="lookup in lookup_state table"
                    GroupTable =1

                    LayoutCachedLeft =13776
                    LayoutCachedTop =36
                    LayoutCachedWidth =14076
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CheckBox
                    OldBorderStyle =0
                    Left =14736
                    Top =36
                    Width =360
                    TabIndex =13
                    Name ="TaxExempt"
                    ControlSource ="UnitTaxExempt"
                    GroupTable =1

                    LayoutCachedLeft =14736
                    LayoutCachedTop =36
                    LayoutCachedWidth =15096
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8016
                    Top =36
                    Width =1116
                    FontSize =9
                    TabIndex =5
                    Name ="Text79"
                    ControlSource ="UnitFinalCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =1

                    LayoutCachedLeft =8016
                    LayoutCachedTop =36
                    LayoutCachedWidth =9132
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =163
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14136
                    Top =36
                    Width =540
                    FontSize =8
                    TabIndex =12
                    Name ="TaxType"
                    ControlSource ="StTaxType"
                    GroupTable =1

                    LayoutCachedLeft =14136
                    LayoutCachedTop =36
                    LayoutCachedWidth =14676
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2556
                    Top =36
                    Width =696
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    ForeColor =255
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    GroupTable =1

                    LayoutCachedLeft =2556
                    LayoutCachedTop =36
                    LayoutCachedWidth =3252
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =36
                    Width =648
                    ColumnWidth =2205
                    FontSize =9
                    Name ="UnitUnitNum"
                    ControlSource ="UnitUnitNum"
                    GroupTable =1

                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =684
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =30
                    Top =240
                    Width =660
                    Height =60
                    FontSize =9
                    TabIndex =14
                    Name ="UnitUnitNum2"
                    ControlSource ="unitunitnum2"

                    LayoutCachedLeft =30
                    LayoutCachedTop =240
                    LayoutCachedWidth =690
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12660
                    Top =36
                    Width =1056
                    FontSize =9
                    TabIndex =10
                    Name ="Text91"
                    ControlSource ="=[Action] & \": \" & Format([lastofunittitlelogdate],\"mm/dd/yy\")"
                    GroupTable =1

                    LayoutCachedLeft =12660
                    LayoutCachedTop =36
                    LayoutCachedWidth =13716
                    LayoutCachedHeight =276
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =14640
                    Top =60
                    Width =240
                    Height =120
                    TabIndex =15
                    Name ="SumOfchildFinalCost"
                    ControlSource ="SumOfchildFinalCost"

                    LayoutCachedLeft =14640
                    LayoutCachedTop =60
                    LayoutCachedWidth =14880
                    LayoutCachedHeight =180
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2505
            Name ="Detail"
            Begin
                Begin Subform
                    Locked = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =780
                    Top =120
                    Width =12360
                    Height =2385
                    Name ="SchChecklistSub"
                    SourceObject ="Report.SchChecklistSub"
                    LinkChildFields ="UnitREF"
                    LinkMasterFields ="UnitID"

                    LayoutCachedLeft =780
                    LayoutCachedTop =120
                    LayoutCachedWidth =13140
                    LayoutCachedHeight =2505
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =240
            BreakLevel =5
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7459
                    Width =1665
                    FontSize =9
                    FontWeight =700
                    Name ="TotChild"
                    ControlSource ="TotUnit"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =7459
                    LayoutCachedWidth =9124
                    LayoutCachedHeight =240
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
                    Left =11070
                    Width =3600
                    FontSize =9
                    TabIndex =1
                    Name ="Text50"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                    LayoutCachedLeft =11070
                    LayoutCachedWidth =14670
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
