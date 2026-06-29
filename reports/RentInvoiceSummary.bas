Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =204
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14561
    DatasheetFontHeight =11
    ItemSuffix =267
    Left =2580
    Top =3345
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0xc8fe3dc9d63ce540
    End
    RecordSource ="RentInvoiceSum"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin Rectangle
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
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
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="LocationName"
        End
        Begin BreakLevel
            ControlSource ="LocationName"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =270
            Name ="ReportHeader"
            Begin
                Begin Line
                    Left =185
                    Top =15
                    Width =0
                    Height =216
                    BorderColor =-2147483615
                    Name ="Line234"
                    LayoutCachedLeft =185
                    LayoutCachedTop =15
                    LayoutCachedWidth =185
                    LayoutCachedHeight =231
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =2925
                    Top =15
                    Width =3975
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label229"
                    Caption ="Location Name:"
                    LayoutCachedLeft =2925
                    LayoutCachedTop =15
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =240
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =7845
                    Top =15
                    Width =1095
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label230"
                    Caption ="Rent"
                    LayoutCachedLeft =7845
                    LayoutCachedTop =15
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =240
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =10620
                    Top =15
                    Width =1065
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label232"
                    Caption ="Tax:"
                    LayoutCachedLeft =10620
                    LayoutCachedTop =15
                    LayoutCachedWidth =11685
                    LayoutCachedHeight =240
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =12825
                    Top =15
                    Width =1395
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label233"
                    Caption ="Total"
                    LayoutCachedLeft =12825
                    LayoutCachedTop =15
                    LayoutCachedWidth =14220
                    LayoutCachedHeight =240
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =7065
                    Top =15
                    Width =600
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label242"
                    Caption ="Units"
                    LayoutCachedLeft =7065
                    LayoutCachedTop =15
                    LayoutCachedWidth =7665
                    LayoutCachedHeight =240
                End
                Begin Label
                    BackStyle =1
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =60
                    Width =2805
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="Label252"
                    Caption ="▼  Summary by Location"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =2865
                    LayoutCachedHeight =270
                    BackThemeColorIndex =2
                End
                Begin Line
                    BorderWidth =3
                    Left =60
                    Top =14
                    Width =14501
                    Name ="Line221"
                    LayoutCachedLeft =60
                    LayoutCachedTop =14
                    LayoutCachedWidth =14561
                    LayoutCachedHeight =14
                    BorderThemeColorIndex =2
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9240
                    Top =15
                    Width =1230
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label260"

                    LayoutCachedLeft =9240
                    LayoutCachedTop =15
                    LayoutCachedWidth =10470
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11745
                    Top =15
                    Width =1005
                    Height =225
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="FeeLabel"
                    ControlSource ="=\"Serv Fee\""

                    LayoutCachedLeft =11745
                    LayoutCachedTop =15
                    LayoutCachedWidth =12750
                    LayoutCachedHeight =240
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
            Height =0
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =288
            Name ="Detail"
            Begin
                Begin Line
                    Left =7860
                    Top =270
                    Width =6365
                    Name ="Line264"
                    LayoutCachedLeft =7860
                    LayoutCachedTop =270
                    LayoutCachedWidth =14225
                    LayoutCachedHeight =270
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                End
                Begin Rectangle
                    OldBorderStyle =0
                    Left =8940
                    Width =1665
                    Height =285
                    Name ="Box266"
                    LayoutCachedLeft =8940
                    LayoutCachedWidth =10605
                    LayoutCachedHeight =285
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7860
                    Width =1080
                    Height =282
                    FontSize =9
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="UnitPerdiem"
                    ControlSource ="BilledRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =7860
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =282
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10620
                    Width =1065
                    Height =282
                    FontSize =9
                    TabIndex =1
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="unitTaxes"
                    ControlSource ="UnitTaxes"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =10620
                    LayoutCachedWidth =11685
                    LayoutCachedHeight =282
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12840
                    Width =1365
                    Height =282
                    FontSize =9
                    TabIndex =2
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="UnitLineTot"
                    ControlSource ="UnitTotalRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =12840
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =282
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9240
                    Top =60
                    Width =1155
                    Height =210
                    FontSize =7
                    FontWeight =700
                    TabIndex =3
                    BorderColor =-2147483615
                    ForeColor =8355711
                    Name ="Text227"
                    ControlSource ="=IIf([UnitTaxes]=0,\"--\",\" Tax Type: \" & [sttaxtype])"
                    TextFormat =1

                    LayoutCachedLeft =9240
                    LayoutCachedTop =60
                    LayoutCachedWidth =10395
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2925
                    Top =27
                    Width =3975
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    BorderColor =-2147483615
                    ForeColor =10512000
                    Name ="Text228"
                    ControlSource ="=[LocationName] & \", \" & [LState] & \" - Sub Total ►\""

                    LayoutCachedLeft =2925
                    LayoutCachedTop =27
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =282
                End
                Begin Line
                    Left =185
                    Width =0
                    Height =288
                    BorderColor =-2147483615
                    Name ="Line240"
                    LayoutCachedLeft =185
                    LayoutCachedWidth =185
                    LayoutCachedHeight =288
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7065
                    Width =600
                    Height =267
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =10512000
                    Name ="Text241"
                    ControlSource ="=[UnitCount]"

                    LayoutCachedLeft =7065
                    LayoutCachedWidth =7665
                    LayoutCachedHeight =267
                    BorderThemeColorIndex =0
                    BorderTint =65.0
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11745
                    Width =1005
                    Height =282
                    FontSize =9
                    TabIndex =6
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="FSLFeeX1"
                    ControlSource ="=IIf([BilledFees]=0,\"--   \",CCur([BilledFees]))"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =11745
                    LayoutCachedWidth =12750
                    LayoutCachedHeight =282
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =330
            Name ="ReportFooter"
            Begin
                Begin Line
                    Left =185
                    Top =216
                    Width =13680
                    BorderColor =-2147483615
                    Name ="Line250"
                    LayoutCachedLeft =185
                    LayoutCachedTop =216
                    LayoutCachedWidth =13865
                    LayoutCachedHeight =216
                End
                Begin Line
                    Left =185
                    Width =0
                    Height =216
                    BorderColor =-2147483615
                    Name ="Line251"
                    LayoutCachedLeft =185
                    LayoutCachedWidth =185
                    LayoutCachedHeight =216
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7860
                    Top =30
                    Width =1080
                    Height =255
                    FontSize =9
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="SubtotalUnits"
                    ControlSource ="=Sum([BilledRent])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =7860
                    LayoutCachedTop =30
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10620
                    Top =30
                    Width =1065
                    Height =270
                    FontSize =9
                    TabIndex =1
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="SubTotalTaxes"
                    ControlSource ="=Sum([UnitTaxes])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =10620
                    LayoutCachedTop =30
                    LayoutCachedWidth =11685
                    LayoutCachedHeight =300
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =12840
                    Top =30
                    Width =1365
                    Height =255
                    FontSize =9
                    TabIndex =2
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="SubTotal"
                    ControlSource ="=Sum([UnitTotalRent])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =12840
                    LayoutCachedTop =30
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2925
                    Top =30
                    Width =3975
                    Height =255
                    FontSize =9
                    TabIndex =3
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Text257"
                    ControlSource ="=\"Schedule: \" & [Schedule] & \"   - Invoice Total ►\""

                    LayoutCachedLeft =2925
                    LayoutCachedTop =30
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7065
                    Top =30
                    Width =600
                    Height =255
                    FontSize =9
                    TabIndex =4
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Text258"
                    ControlSource ="=Sum([UnitCount])"

                    LayoutCachedLeft =7065
                    LayoutCachedTop =30
                    LayoutCachedWidth =7665
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11745
                    Top =30
                    Width =1005
                    Height =255
                    FontSize =9
                    TabIndex =5
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="FSLFeeX2"
                    ControlSource ="=Sum([BilledFees])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =11745
                    LayoutCachedTop =30
                    LayoutCachedWidth =12750
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
            End
        End
    End
End
CodeBehindForm
' See "RentInvoiceSummary.cls"
