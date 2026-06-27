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
    Width =14565
    DatasheetFontHeight =11
    ItemSuffix =260
    Left =10170
    Top =10695
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x41bd6841173ce540
    End
    RecordSource ="SELECT PD.LocationName, PD.LState, PD.SchNo, PD.sttaxtype, PD.UnitTaxExempt, Sum"
        "(PD.Countme) AS UnitCount, Sum(PD.TotalPerDiem) AS TotalPerDiem, Sum(PD.Taxes) A"
        "S TotalTaxes, Sum(PD.Refund) AS Refund, Sum(PD.RefundTax) AS RefundTax, Sum(PD.L"
        "ineTotal) AS LineTotal, Sum(PD.PDDAYS) AS TotPDDAYS, PD.BillID FROM vw_PerDiemIn"
        "voice AS PD GROUP BY PD.LocationName, PD.LState, PD.SchNo, PD.sttaxtype, PD.Unit"
        "TaxExempt, PD.BillID HAVING (((PD.BillID)=[Forms]![Billing]![BillID]));"
    Caption ="Perdiem Bill Summary"
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
            Height =300
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
                    TextAlign =1
                    TextFontFamily =34
                    Left =6510
                    Top =45
                    Width =3960
                    Height =255
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label229"
                    Caption ="Location:"
                    LayoutCachedLeft =6510
                    LayoutCachedTop =45
                    LayoutCachedWidth =10470
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =10635
                    Top =45
                    Width =1035
                    Height =255
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label230"
                    Caption ="Rent:"
                    LayoutCachedLeft =10635
                    LayoutCachedTop =45
                    LayoutCachedWidth =11670
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =12360
                    Top =45
                    Width =870
                    Height =255
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label232"
                    Caption ="Tax:"
                    LayoutCachedLeft =12360
                    LayoutCachedTop =45
                    LayoutCachedWidth =13230
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =13320
                    Top =45
                    Width =1200
                    Height =255
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label233"
                    Caption ="Total:"
                    LayoutCachedLeft =13320
                    LayoutCachedTop =45
                    LayoutCachedWidth =14520
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =4740
                    Top =45
                    Width =600
                    Height =255
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label242"
                    Caption ="Units:"
                    LayoutCachedLeft =4740
                    LayoutCachedTop =45
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =300
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =60
                    Width =2805
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BackColor =9605778
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Label252"
                    Caption ="▼  Summary by Location"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =2865
                    LayoutCachedHeight =270
                End
                Begin Line
                    BorderWidth =3
                    Left =60
                    Top =14
                    Width =14501
                    BorderColor =9605778
                    Name ="Line221"
                    LayoutCachedLeft =60
                    LayoutCachedTop =14
                    LayoutCachedWidth =14561
                    LayoutCachedHeight =14
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
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10575
                    Top =27
                    Width =1095
                    FontSize =9
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="UnitPerdiem"
                    ControlSource ="TotalPerDiem"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =10575
                    LayoutCachedTop =27
                    LayoutCachedWidth =11670
                    LayoutCachedHeight =267
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12255
                    Top =27
                    Width =975
                    FontSize =9
                    TabIndex =1
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="unitTaxes"
                    ControlSource ="TotalTaxes"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =12255
                    LayoutCachedTop =27
                    LayoutCachedWidth =13230
                    LayoutCachedHeight =267
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13335
                    Top =27
                    Width =1200
                    FontSize =9
                    TabIndex =2
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="UnitLineTot"
                    ControlSource ="LineTotal"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =13335
                    LayoutCachedTop =27
                    LayoutCachedWidth =14535
                    LayoutCachedHeight =267
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6420
                    Top =27
                    Width =4050
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    BorderColor =-2147483615
                    ForeColor =10512000
                    Name ="Text228"
                    ControlSource ="=[LocationName] & \", \" & [LState] & \" - Sub Total ►\""

                    LayoutCachedLeft =6420
                    LayoutCachedTop =27
                    LayoutCachedWidth =10470
                    LayoutCachedHeight =267
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
                    Left =4740
                    Top =27
                    Width =600
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    BorderColor =-2147483615
                    ForeColor =10512000
                    Name ="Text241"
                    ControlSource ="UnitCount"

                    LayoutCachedLeft =4740
                    LayoutCachedTop =27
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =267
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
            Height =975
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
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10575
                    Top =675
                    Width =1095
                    Height =255
                    FontSize =9
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="SubTotalWRefund"
                    ControlSource ="=Sum([TotalPerDiem])-Sum([Refund])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =10575
                    LayoutCachedTop =675
                    LayoutCachedWidth =11670
                    LayoutCachedHeight =930
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =12255
                    Top =690
                    Width =975
                    Height =255
                    FontSize =9
                    TabIndex =1
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="SumTotalTaxes"
                    ControlSource ="=Sum([TotalTaxes])-Sum([RefundTax])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =12255
                    LayoutCachedTop =690
                    LayoutCachedWidth =13230
                    LayoutCachedHeight =945
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =13335
                    Top =690
                    Width =1200
                    Height =255
                    FontSize =9
                    TabIndex =2
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="GTotal"
                    ControlSource ="=Sum([LineTotal])-(Sum([Refund])+Sum([RefundTax]))"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =13335
                    LayoutCachedTop =690
                    LayoutCachedWidth =14535
                    LayoutCachedHeight =945
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6420
                    Top =675
                    Width =4050
                    Height =255
                    FontSize =9
                    TabIndex =3
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Text247"
                    ControlSource ="=\"Schedule: \" & [Schno] & \"   - Invoice Total ►\""

                    LayoutCachedLeft =6420
                    LayoutCachedTop =675
                    LayoutCachedWidth =10470
                    LayoutCachedHeight =930
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4740
                    Top =690
                    Width =600
                    Height =255
                    FontSize =9
                    TabIndex =4
                    BorderColor =-2147483615
                    ForeColor =16777215
                    Name ="Text249"
                    ControlSource ="=Sum([UnitCount])"

                    LayoutCachedLeft =4740
                    LayoutCachedTop =690
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =945
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =13335
                    Top =330
                    Width =1200
                    Height =255
                    FontSize =9
                    TabIndex =5
                    BackColor =3368550
                    BorderColor =3368550
                    ForeColor =16777215
                    Name ="TotalRefund"
                    ControlSource ="=IIf([Refund]=0,Null,(Sum([Refund])+Sum([RefundTax]))*-1)"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Tot Total"

                    LayoutCachedLeft =13335
                    LayoutCachedTop =330
                    LayoutCachedWidth =14535
                    LayoutCachedHeight =585
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10575
                    Top =345
                    Width =1095
                    Height =255
                    FontSize =9
                    TabIndex =6
                    BackColor =3368550
                    BorderColor =3368550
                    ForeColor =16777215
                    Name ="Taxes"
                    ControlSource ="=IIf([Refund]=0,Null,Sum([Refund])*-1)"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =10575
                    LayoutCachedTop =345
                    LayoutCachedWidth =11670
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6420
                    Top =345
                    Width =4050
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =7
                    BorderColor =3368550
                    ForeColor =3368550
                    Name ="Text138"
                    ControlSource ="=IIf([Refund]=0,Null,\" Ancillary Services Fee Credit - Sub Total ►\")"

                    LayoutCachedLeft =6420
                    LayoutCachedTop =345
                    LayoutCachedWidth =10470
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextFontCharSet =163
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =12240
                    Top =345
                    Width =990
                    Height =255
                    FontSize =9
                    TabIndex =8
                    BackColor =3368550
                    BorderColor =3368550
                    ForeColor =16777215
                    Name ="SubTotalTaxesRefund"
                    ControlSource ="=IIf([Refund]=0,Null,Sum([RefundTax])*-1)"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =12240
                    LayoutCachedTop =345
                    LayoutCachedWidth =13230
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10575
                    Top =30
                    Width =1095
                    Height =255
                    FontSize =9
                    TabIndex =9
                    BackColor =8355711
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="SubtotalUnits"
                    ControlSource ="=Sum([TotalPerDiem])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =10575
                    LayoutCachedTop =30
                    LayoutCachedWidth =11670
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =12255
                    Top =30
                    Width =975
                    Height =255
                    FontSize =9
                    TabIndex =10
                    BackColor =8355711
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="SubTotalTaxes"
                    ControlSource ="=Sum([TotalTaxes])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =12255
                    LayoutCachedTop =30
                    LayoutCachedWidth =13230
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =13335
                    Top =30
                    Width =1200
                    Height =255
                    FontSize =9
                    TabIndex =11
                    BackColor =8355711
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="SubTotal"
                    ControlSource ="=Sum([LineTotal])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =13335
                    LayoutCachedTop =30
                    LayoutCachedWidth =14535
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6420
                    Top =30
                    Width =4050
                    Height =255
                    FontSize =9
                    TabIndex =12
                    BackColor =8355711
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Text257"
                    ControlSource ="=\"Schedule: \" & [Schno] & \"   - Invoice Total ►\""

                    LayoutCachedLeft =6420
                    LayoutCachedTop =30
                    LayoutCachedWidth =10470
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4740
                    Top =30
                    Width =600
                    Height =255
                    FontSize =9
                    TabIndex =13
                    BackColor =8355711
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Text258"
                    ControlSource ="=Sum([UnitCount])"

                    LayoutCachedLeft =4740
                    LayoutCachedTop =30
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =186
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2205
                    Top =360
                    Width =1065
                    ColumnWidth =2760
                    FontSize =9
                    TabIndex =14
                    Name ="TotPDDAYS"
                    ControlSource ="TotPDDAYS"

                    LayoutCachedLeft =2205
                    LayoutCachedTop =360
                    LayoutCachedWidth =3270
                    LayoutCachedHeight =600
                End
            End
        End
    End
End
