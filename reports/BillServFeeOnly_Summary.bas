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
    Left =4380
    Top =9315
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0xc98c19653180e440
    End
    RecordSource ="FABankBillRegSummary"
    Caption ="PerDiem Bill"
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
            Height =303
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
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Left =3570
                    Top =29
                    Width =3480
                    Height =274
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label229"
                    Caption ="Location Name:"
                    LayoutCachedLeft =3570
                    LayoutCachedTop =29
                    LayoutCachedWidth =7050
                    LayoutCachedHeight =303
                End
                Begin Label
                    Visible = NotDefault
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =12840
                    Top =29
                    Width =1365
                    Height =274
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label233"
                    Caption ="Total:"
                    LayoutCachedLeft =12840
                    LayoutCachedTop =29
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =303
                End
                Begin Label
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    Left =7215
                    Top =29
                    Width =600
                    Height =274
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label242"
                    Caption ="Units"
                    LayoutCachedLeft =7215
                    LayoutCachedTop =29
                    LayoutCachedWidth =7815
                    LayoutCachedHeight =303
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
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9240
                    Top =29
                    Width =1230
                    Height =274
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label260"

                    LayoutCachedLeft =9240
                    LayoutCachedTop =29
                    LayoutCachedWidth =10470
                    LayoutCachedHeight =303
                End
                Begin TextBox
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11745
                    Top =29
                    Width =1005
                    Height =274
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Text261"
                    ControlSource ="=\"Serv. Fee:\""

                    LayoutCachedLeft =11745
                    LayoutCachedTop =29
                    LayoutCachedWidth =12750
                    LayoutCachedHeight =303
                End
                Begin TextBox
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10605
                    Top =29
                    Width =1080
                    Height =274
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label232"

                    LayoutCachedLeft =10605
                    LayoutCachedTop =29
                    LayoutCachedWidth =11685
                    LayoutCachedHeight =303
                End
                Begin TextBox
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7980
                    Top =29
                    Width =1110
                    Height =274
                    FontSize =9
                    FontWeight =700
                    BackColor =14211288
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Label230"

                    LayoutCachedLeft =7980
                    LayoutCachedTop =29
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =303
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
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7980
                    Width =1110
                    Height =286
                    FontSize =9
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="UnitPerdiem"
                    ControlSource ="RentNoTax"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =7980
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =286
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10620
                    Width =1065
                    Height =286
                    FontSize =9
                    TabIndex =1
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="unitTaxes"
                    ControlSource ="=[TaxAmt]+[FSLTax]"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =10620
                    LayoutCachedWidth =11685
                    LayoutCachedHeight =286
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12840
                    Width =1365
                    Height =286
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
                    LayoutCachedHeight =286
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =186
                    TextAlign =1
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
                    ControlSource ="=\" Tax Type: \" & IIf([sttaxtype]=\"Upfront\",[sttaxtype],IIf([UnitTaxExempt]=\""
                        "Yes\",\"Exempt\",[sttaxtype]))"
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
                    Left =3570
                    Top =27
                    Width =3480
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    BorderColor =-2147483615
                    ForeColor =10512000
                    Name ="Text228"
                    ControlSource ="=[LocationName] & \", \" & [LState] & \" - Sub Total ►\""

                    LayoutCachedLeft =3570
                    LayoutCachedTop =27
                    LayoutCachedWidth =7050
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
                    Left =7215
                    Width =600
                    Height =259
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    BorderColor =-2147483615
                    ForeColor =10512000
                    Name ="Text241"
                    ControlSource ="=[UnitCount]"

                    LayoutCachedLeft =7215
                    LayoutCachedWidth =7815
                    LayoutCachedHeight =259
                End
                Begin TextBox
                    SpecialEffect =5
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11745
                    Width =1005
                    Height =286
                    FontSize =9
                    TabIndex =6
                    BorderColor =-2147483615
                    ForeColor =-2147483617
                    Name ="Text262"
                    ControlSource ="FSLFee"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="TaxAmt Total"

                    LayoutCachedLeft =11745
                    LayoutCachedWidth =12750
                    LayoutCachedHeight =286
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
            Height =660
            Name ="ReportFooter"
            Begin
                Begin Line
                    Left =185
                    Top =216
                    Width =12096
                    BorderColor =-2147483615
                    Name ="Line250"
                    LayoutCachedLeft =185
                    LayoutCachedTop =216
                    LayoutCachedWidth =12281
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
                    Visible = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7980
                    Top =30
                    Width =1110
                    Height =255
                    FontSize =9
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="SubtotalUnits"
                    ControlSource ="=Sum([RentNoTax])"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="LineTot Total"

                    LayoutCachedLeft =7980
                    LayoutCachedTop =30
                    LayoutCachedWidth =9090
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    Visible = NotDefault
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
                    Visible = NotDefault
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
                    Left =3570
                    Top =30
                    Width =3480
                    Height =255
                    FontSize =9
                    TabIndex =3
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Text257"
                    ControlSource ="=\"Schedule: \" & [Schno] & \"   - Invoice Total ►\""

                    LayoutCachedLeft =3570
                    LayoutCachedTop =30
                    LayoutCachedWidth =7050
                    LayoutCachedHeight =285
                    BackThemeColorIndex =2
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextFontCharSet =186
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7215
                    Top =30
                    Width =600
                    Height =255
                    FontSize =9
                    TabIndex =4
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Text258"
                    ControlSource ="=Sum([UnitCount])"

                    LayoutCachedLeft =7215
                    LayoutCachedTop =30
                    LayoutCachedWidth =7815
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
