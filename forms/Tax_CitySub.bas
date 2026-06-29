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
    Width =8154
    DatasheetFontHeight =11
    ItemSuffix =210
    Left =2070
    Top =4755
    Right =13335
    Bottom =10815
    RecSrcDt = Begin
        0xa10567abb3f8e440
    End
    RecordSource ="SELECT TaxCities.*, TaxStatesID, [CityName] & ' (' & [CountyName] & ')' AS City2"
        " FROM (TaxCities INNER JOIN TaxCounties ON TaxCities.TaxCountyID = TaxCounties.I"
        "D) INNER JOIN TaxStates ON TaxCounties.TaxStatesID = TaxStates.ID WHERE TaxState"
        "sID= 10 ORDER BY [CityName] & ' (' & [CountyName] & ')';"
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
                    TextAlign =2
                    Left =5685
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
                    LayoutCachedLeft =5685
                    LayoutCachedTop =30
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =465
                    ColumnStart =5
                    ColumnEnd =5
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5520
                    Top =60
                    Width =120
                    Height =210
                    ColumnOrder =0
                    FontSize =8
                    BorderColor =12632256
                    ForeColor =0
                    Name ="txtCurrent"
                    GridlineColor =0

                    LayoutCachedLeft =5520
                    LayoutCachedTop =60
                    LayoutCachedWidth =5640
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
                    Left =60
                    Top =30
                    Width =2895
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="StateNameLabel"
                    Caption ="City (county)"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =60
                    LayoutCachedTop =30
                    LayoutCachedWidth =2955
                    LayoutCachedHeight =255
                    ColumnStart =3
                    ColumnEnd =3
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =1
                    Left =2940
                    Top =30
                    Width =2475
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="TaxNumberLabel"
                    Caption ="City"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2940
                    LayoutCachedTop =30
                    LayoutCachedWidth =5415
                    LayoutCachedHeight =255
                    ColumnStart =4
                    ColumnEnd =4
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7095
                    Top =30
                    Width =945
                    Height =435
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label169"
                    Caption ="Rent Receipt"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =7095
                    LayoutCachedTop =30
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =465
                    ColumnStart =7
                    ColumnEnd =9
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =1
                    Left =2940
                    Top =246
                    Width =2475
                    Height =216
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label206"
                    Caption ="Tax ID Number"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2940
                    LayoutCachedTop =246
                    LayoutCachedWidth =5415
                    LayoutCachedHeight =462
                    ColumnStart =4
                    ColumnEnd =4
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =60
                    Top =255
                    Width =2895
                    Height =225
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="Label207"
                    Caption ="Name"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =60
                    LayoutCachedTop =255
                    LayoutCachedWidth =2955
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
                    Left =5685
                    Width =1395
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    BorderColor =10319446
                    ForeColor =0
                    Name ="SalesTaxOrUseTaxRenewDate"
                    ControlSource ="SalesTaxOrUseTaxRenewDate"
                    Format ="Short Date"
                    ControlTipText ="FA Tax Cert Renewal Date for this City"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5685
                    LayoutCachedWidth =7080
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
                    Width =2895
                    Height =255
                    FontSize =10
                    TabIndex =1
                    ForeColor =0
                    Name ="CityName"
                    ControlSource ="City2"
                    ControlTipText ="Double-Click to Edit Event"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =60
                    LayoutCachedWidth =2955
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
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2940
                    Width =2475
                    Height =255
                    FontSize =10
                    TabIndex =2
                    BorderColor =10319446
                    ForeColor =0
                    Name ="SalesTaxOrUseTaxNumber"
                    ControlSource ="SalesTaxOrUseTaxNumber"
                    ControlTipText ="Tax Cert Number provided to FA from this City"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2940
                    LayoutCachedWidth =5415
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
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
                    Left =7095
                    Width =273
                    Height =231
                    FontSize =8
                    Name ="ID"
                    ControlSource ="ID"
                    TopPadding =0

                    LayoutCachedLeft =7095
                    LayoutCachedWidth =7368
                    LayoutCachedHeight =231
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Top =240
                    Width =8154
                    Height =43
                    TabIndex =5
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
                    LayoutCachedWidth =8154
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
                    OverlapFlags =93
                    Left =7500
                    Top =50
                    Width =390
                    Height =195
                    TabIndex =4
                    Name ="RentReceiptAvailable"
                    ControlSource ="RentReceiptAvailable"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7500
                    LayoutCachedTop =50
                    LayoutCachedWidth =7890
                    LayoutCachedHeight =245
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =9
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7725
                    Width =273
                    Height =231
                    FontSize =8
                    TabIndex =6
                    Name ="TaxCountyID"
                    ControlSource ="TaxCountyID"
                    TopPadding =0

                    LayoutCachedLeft =7725
                    LayoutCachedWidth =7998
                    LayoutCachedHeight =231
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
' See "Tax_CitySub.cls"
