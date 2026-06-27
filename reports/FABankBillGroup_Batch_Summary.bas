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
    ItemSuffix =264
    Left =6000
    Top =4785
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0x8874c89c2c46e640
    End
    RecordSource ="FABankBillGroupBatchSummary"
    Caption ="BankBillSummary"
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
            Height =285
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
                    Left =4920
                    Top =15
                    Width =3480
                    Height =270
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label229"
                    Caption ="Location Name:"
                    LayoutCachedLeft =4920
                    LayoutCachedTop =15
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =285
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =9300
                    Top =15
                    Width =1200
                    Height =270
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label230"
                    Caption ="Rent:"
                    LayoutCachedLeft =9300
                    LayoutCachedTop =15
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =285
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =10620
                    Top =15
                    Width =1065
                    Height =270
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
                    LayoutCachedHeight =285
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =12840
                    Top =15
                    Width =1365
                    Height =270
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label233"
                    Caption ="Total:"
                    LayoutCachedLeft =12840
                    LayoutCachedTop =15
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =285
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =8565
                    Top =15
                    Width =600
                    Height =270
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label242"
                    Caption ="Units"
                    LayoutCachedLeft =8565
                    LayoutCachedTop =15
                    LayoutCachedWidth =9165
                    LayoutCachedHeight =285
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
                    Left =11745
                    Top =15
                    Width =1005
                    Height =270
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Text261"
                    ControlSource ="=\"Serv. Fee:\""

                    LayoutCachedLeft =11745
                    LayoutCachedTop =15
                    LayoutCachedWidth =12750
                    LayoutCachedHeight =285
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
            Height =288
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9295
                    Width =1200
                    Height =282
                    FontSize =9
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="UnitPerdiem"
                    ControlSource ="RentNoTax"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =9295
                    LayoutCachedWidth =10495
                    LayoutCachedHeight =282
                End
                Begin TextBox
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10615
                    Width =1065
                    Height =282
                    FontSize =9
                    TabIndex =1
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="unitTaxes"
                    ControlSource ="=[TaxAmt]+[FSLTax]"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =10615
                    LayoutCachedWidth =11680
                    LayoutCachedHeight =282
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12835
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

                    LayoutCachedLeft =12835
                    LayoutCachedWidth =14200
                    LayoutCachedHeight =282
                End
                Begin TextBox
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4915
                    Width =3480
                    Height =282
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    BorderColor =-2147483615
                    ForeColor =10512000
                    Name ="Text228"
                    ControlSource ="=[LocationName] & \", \" & [LState] & \" - Sub Total ►\""

                    LayoutCachedLeft =4915
                    LayoutCachedWidth =8395
                    LayoutCachedHeight =282
                End
                Begin Line
                    Left =180
                    Width =0
                    Height =288
                    BorderColor =-2147483615
                    Name ="Line240"
                    LayoutCachedLeft =180
                    LayoutCachedWidth =180
                    LayoutCachedHeight =288
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8560
                    Width =600
                    Height =282
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    BorderColor =-2147483615
                    ForeColor =10512000
                    Name ="Text241"
                    ControlSource ="=[UnitCount]"

                    LayoutCachedLeft =8560
                    LayoutCachedWidth =9160
                    LayoutCachedHeight =282
                End
                Begin TextBox
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11740
                    Width =1005
                    Height =282
                    FontSize =9
                    TabIndex =5
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Text262"
                    ControlSource ="FSLFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =11740
                    LayoutCachedWidth =12745
                    LayoutCachedHeight =282
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="Detail"
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
            Height =420
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
                    Left =9300
                    Top =30
                    Width =1200
                    Height =255
                    FontSize =9
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="SubtotalUnits"
                    ControlSource ="=Sum([RentNoTax])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =9300
                    LayoutCachedTop =30
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10620
                    Top =30
                    Width =1065
                    Height =255
                    FontSize =9
                    TabIndex =1
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="SubTotalTaxes"
                    ControlSource ="=Sum([TaxAmt]+[FSLTax])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =10620
                    LayoutCachedTop =30
                    LayoutCachedWidth =11685
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =1
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
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4920
                    Top =30
                    Width =3480
                    Height =255
                    FontSize =9
                    TabIndex =3
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Text257"
                    ControlSource ="=\"   - Invoice Total ►\""

                    LayoutCachedLeft =4920
                    LayoutCachedTop =30
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8565
                    Top =30
                    Width =600
                    Height =255
                    FontSize =9
                    TabIndex =4
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Text258"
                    ControlSource ="=Sum([UnitCount])"

                    LayoutCachedLeft =8565
                    LayoutCachedTop =30
                    LayoutCachedWidth =9165
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =1
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
                    Name ="Text263"
                    ControlSource ="=Sum([FSLFee])"
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
