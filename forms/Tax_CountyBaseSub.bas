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
    Width =10800
    DatasheetFontHeight =11
    ItemSuffix =217
    Left =885
    Top =12705
    Right =17700
    Bottom =15090
    RecSrcDt = Begin
        0x2d182052cef8e440
    End
    RecordSource ="SELECT TaxCountiesBase.id, TaxCountiesBase.TaxCountyId, TaxCountiesBase.TaxRule,"
        " TaxCountiesBase.TaxTypeId, TaxCountiesBase.BaseTaxRate, TaxCountiesBase.StartsO"
        "n, TaxCountiesBase.EndsOn, IIf(Not IsNull([StartsOn]) And IsNull([EndsOn]),-1,0)"
        " AS IsActive, TaxCounties.CountyName, TaxType.TaxFullName, TaxType.TaxAbbreviati"
        "on, TaxType.TaxDescription FROM (TaxCountiesBase INNER JOIN TaxType ON TaxCounti"
        "esBase.TaxTypeId = TaxType.Id) INNER JOIN TaxCounties ON TaxCountiesBase.TaxCoun"
        "tyId = TaxCounties.Id;"
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
                    OverlapFlags =93
                    TextAlign =2
                    Left =4935
                    Top =30
                    Width =1080
                    Height =435
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label3"
                    Caption ="Starts \015\012ON"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =4935
                    LayoutCachedTop =30
                    LayoutCachedWidth =6015
                    LayoutCachedHeight =465
                    ColumnStart =5
                    ColumnEnd =5
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5535
                    Top =60
                    Width =120
                    Height =210
                    ColumnOrder =0
                    FontSize =8
                    BorderColor =12632256
                    Name ="txtCurrent"

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
                    Left =4080
                    Top =30
                    Width =835
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="TaxNumberLabel"
                    Caption ="Tax"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =4080
                    LayoutCachedTop =30
                    LayoutCachedWidth =4915
                    LayoutCachedHeight =255
                    ColumnStart =4
                    ColumnEnd =4
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =7095
                    Top =30
                    Width =945
                    Height =435
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label169"
                    Caption ="\015\012Active"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
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
                    Left =4080
                    Top =246
                    Width =835
                    Height =216
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label206"
                    Caption ="Rate %"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =4080
                    LayoutCachedTop =246
                    LayoutCachedWidth =4915
                    LayoutCachedHeight =462
                    ColumnStart =4
                    ColumnEnd =4
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =6015
                    Top =30
                    Width =1080
                    Height =435
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label210"
                    Caption ="Ends\015\012ON"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =6015
                    LayoutCachedTop =30
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =465
                    ColumnStart =5
                    ColumnEnd =5
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =2355
                    Top =30
                    Width =1725
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label212"
                    Caption ="Tax"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2355
                    LayoutCachedTop =30
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =255
                    ColumnStart =3
                    ColumnEnd =3
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =2355
                    Top =255
                    Width =1725
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label213"
                    Caption ="Description"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2355
                    LayoutCachedTop =255
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =480
                    ColumnStart =3
                    ColumnEnd =3
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =86
                    Width =2295
                    Height =314
                    FontWeight =700
                    TabIndex =1
                    Name ="StateNameLabel"
                    ControlSource ="CountyName"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =86
                    LayoutCachedWidth =2355
                    LayoutCachedHeight =400
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
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
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4935
                    Width =1080
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    BorderColor =10921638
                    Name ="EventType"
                    ControlSource ="StartsOn"
                    ControlTipText ="Double-Click to Edit Event"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =4935
                    LayoutCachedWidth =6015
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
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
                    Width =2220
                    Height =255
                    FontSize =10
                    TabIndex =1
                    BorderColor =10921638
                    Name ="StateName"
                    ControlSource ="TaxFullName"
                    ControlTipText ="Double-Click to Edit Event"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedWidth =2280
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
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1920
                    Width =273
                    Height =231
                    FontSize =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Id"
                    ControlSource ="Id"
                    TopPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedWidth =2193
                    LayoutCachedHeight =231
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =7500
                    Top =50
                    Width =216
                    Height =195
                    TabIndex =4
                    BorderColor =10921638
                    Name ="IsActive"
                    ControlSource ="IsActive"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =7500
                    LayoutCachedTop =50
                    LayoutCachedWidth =7716
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
                    Left =1560
                    Width =273
                    Height =231
                    FontSize =8
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TaxCountyId"
                    ControlSource ="TaxCountyId"
                    TopPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1560
                    LayoutCachedWidth =1833
                    LayoutCachedHeight =231
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6015
                    Width =1080
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =7
                    BorderColor =10921638
                    Name ="Text211"
                    ControlSource ="EndsOn"
                    ControlTipText ="Double-Click to Edit Event"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =6015
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2340
                    Width =1725
                    Height =255
                    FontSize =10
                    TabIndex =8
                    BorderColor =10921638
                    Name ="Text214"
                    ControlSource ="TaxAbbreviation"
                    ControlTipText ="Double-Click to Edit Event"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =2340
                    LayoutCachedWidth =4065
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
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Top =240
                    Width =10800
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

                    LayoutCachedTop =240
                    LayoutCachedWidth =10800
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
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4080
                    Width =835
                    Height =255
                    FontSize =10
                    TabIndex =2
                    BackColor =11525325
                    BorderColor =9211020
                    Name ="TaxNumber"
                    ControlSource ="BaseTaxRate"
                    Format ="Percent"
                    ValidationRule ="<=1 And >=-1"
                    ValidationText ="Tax Rates cannot be > 100%"
                    OnGotFocus ="[Event Procedure]"
                    ControlTipText ="The percentage of the Tax Rate"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =4080
                    LayoutCachedWidth =4915
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =8294
                    Width =465
                    Height =256
                    FontSize =8
                    TabIndex =9
                    Name ="EditMe"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Edit This Exception"
                    TopPadding =0
                    BottomPadding =0
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4af9b17d4a78b17d4a18 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a81b17d4affb17d4af3 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a27b17d4af6b17d4a03 ,
                        0xb17d4ab7b17d4a6c000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a0cb17d4ab7 ,
                        0xb17d4affb17d4affb17d4a7b0000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a87 ,
                        0xb17d4affb17d4affb17d4affb17d4a8700000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xb17d4a8db17d4affb17d4affb17d4affb17d4a93000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000b17d4a90b17d4affb17d4affb17d4aabb17d4a12b17d4a0300000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a96b17d4aabb17d4a15b17d4acfb17d4aa500000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a12b17d4acfb17d4affb17d4af000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a03b17d4aa2b17d4afcb17d4a2a00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =8294
                    LayoutCachedWidth =8759
                    LayoutCachedHeight =256
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =8965045
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =9211020
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    Overlaps =1
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
' See "Tax_CountyBaseSub.cls"
