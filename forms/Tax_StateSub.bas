Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8064
    DatasheetFontHeight =11
    ItemSuffix =208
    Left =570
    Top =1785
    Right =8955
    Bottom =4050
    RecSrcDt = Begin
        0x3ae97d2bb0f8e440
    End
    RecordSource ="Select * From TaxStates Order by PostalAbbreviation"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BoundObjectFrame
            AddColon = NotDefault
            SizeMode =3
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =-1800
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =480
            BackColor =14211288
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =30
                    Width =465
                    Height =435
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="StateLabel"
                    Caption ="Abbr"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =60
                    LayoutCachedTop =30
                    LayoutCachedWidth =525
                    LayoutCachedHeight =465
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5700
                    Top =30
                    Width =1395
                    Height =435
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label3"
                    Caption ="Tax Cert\015\012Renewal"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =5700
                    LayoutCachedTop =30
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =465
                    ColumnStart =5
                    ColumnEnd =5
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =540
                    Top =30
                    Width =690
                    Height =435
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label23"
                    Caption =" Accepts\015\012FA Trust"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =540
                    LayoutCachedTop =30
                    LayoutCachedWidth =1230
                    LayoutCachedHeight =465
                    ColumnStart =1
                    ColumnEnd =2
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5535
                    Top =60
                    Width =120
                    Height =210
                    ColumnOrder =0
                    FontSize =8
                    BorderColor =12632256
                    ForeColor =0
                    Name ="txtCurrent"
                    GridlineColor =0

                    LayoutCachedLeft =5535
                    LayoutCachedTop =60
                    LayoutCachedWidth =5655
                    LayoutCachedHeight =270
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =1260
                    Top =30
                    Width =1665
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="StateNameLabel"
                    Caption ="State"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =1260
                    LayoutCachedTop =30
                    LayoutCachedWidth =2925
                    LayoutCachedHeight =255
                    ColumnStart =3
                    ColumnEnd =3
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =2955
                    Top =30
                    Width =2550
                    Height =216
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="TaxNumberLabel"
                    Caption ="State"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2955
                    LayoutCachedTop =30
                    LayoutCachedWidth =5505
                    LayoutCachedHeight =246
                    ColumnStart =4
                    ColumnEnd =4
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7110
                    Top =30
                    Width =945
                    Height =435
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label169"
                    Caption ="Rent Rceipt\015\012Avail"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =7110
                    LayoutCachedTop =30
                    LayoutCachedWidth =8055
                    LayoutCachedHeight =465
                    ColumnStart =7
                    ColumnEnd =9
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =2955
                    Top =246
                    Width =2550
                    Height =216
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label206"
                    Caption ="Tax ID Number"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2955
                    LayoutCachedTop =246
                    LayoutCachedWidth =5505
                    LayoutCachedHeight =462
                    ColumnStart =4
                    ColumnEnd =4
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =1260
                    Top =255
                    Width =1665
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label207"
                    Caption ="Name"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =1260
                    LayoutCachedTop =255
                    LayoutCachedWidth =2925
                    LayoutCachedHeight =480
                    ColumnStart =3
                    ColumnEnd =3
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =283
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5700
                    Width =1395
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    BorderColor =10319446
                    ForeColor =0
                    Name ="SalesTaxOrUseTaxRenewDate"
                    ControlSource ="SalesTaxOrUseTaxRenewDate"
                    Format ="Short Date"
                    ControlTipText ="FA Tax Cert Renewal Date for this State"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5700
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Width =401
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="State"
                    ControlSource ="PostalAbbreviation"
                    ControlTipText ="State"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =60
                    LayoutCachedWidth =461
                    LayoutCachedHeight =255
                    RowStart =1
                    RowEnd =2
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1260
                    Width =1665
                    Height =255
                    FontSize =10
                    TabIndex =3
                    ForeColor =0
                    Name ="StateName"
                    ControlSource ="StateName"
                    ControlTipText ="StateName"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1260
                    LayoutCachedWidth =2925
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =50
                    IMESentenceMode =3
                    Left =2955
                    Width =2550
                    Height =255
                    FontSize =10
                    TabIndex =4
                    BorderColor =10319446
                    ForeColor =0
                    Name ="SalesTaxOrUseTaxNumber"
                    ControlSource ="SalesTaxOrUseTaxNumber"
                    FontName ="OCR A Extended"
                    ControlTipText ="Tax Cert Number provided to FA from this state"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2955
                    LayoutCachedWidth =5505
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7110
                    Width =273
                    Height =231
                    FontSize =8
                    Name ="Id"
                    ControlSource ="Id"
                    TopPadding =0

                    LayoutCachedLeft =7110
                    LayoutCachedWidth =7383
                    LayoutCachedHeight =231
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Top =240
                    Width =8064
                    Height =43
                    TabIndex =7
                    BorderColor =12632256
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x0100000086000000010000000100000000000000000000001200000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00490044005d003d005b00740078007400430075007200720065006e007400 ,
                        0x5d0000000000
                    End
                    GridlineColor =0

                    LayoutCachedTop =240
                    LayoutCachedWidth =8064
                    LayoutCachedHeight =283
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00110000005b00 ,
                        0x490044005d003d005b00740078007400430075007200720065006e0074005d00 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =780
                    Top =50
                    Width =450
                    Height =195
                    TabIndex =2
                    Name ="AcceptsFATitleTrust"
                    ControlSource ="AcceptsFATitleTrust"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =50
                    LayoutCachedWidth =1230
                    LayoutCachedHeight =245
                    RowStart =1
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =7515
                    Top =50
                    Width =390
                    Height =195
                    TabIndex =6
                    Name ="RentReceiptAvailable"
                    ControlSource ="RentReceiptAvailable"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7515
                    LayoutCachedTop =50
                    LayoutCachedWidth =7905
                    LayoutCachedHeight =245
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =9
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =14211288
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
' See "Tax_StateSub.cls"
