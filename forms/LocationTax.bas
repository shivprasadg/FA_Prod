Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5160
    DatasheetFontHeight =11
    ItemSuffix =348
    Left =5760
    Top =9285
    Right =12060
    Bottom =11565
    RecSrcDt = Begin
        0x9480fa5bf5f7e440
    End
    RecordSource ="SELECT * From vw_LocationTaxes"
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =2160
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =93
                    Left =555
                    Top =165
                    Width =3930
                    Height =375
                    BorderColor =8872517
                    Name ="Box297"
                    LayoutCachedLeft =555
                    LayoutCachedTop =165
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =540
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =223
                    Left =480
                    Top =240
                    Width =1455
                    Height =345
                    Name ="Box298"
                    LayoutCachedLeft =480
                    LayoutCachedTop =240
                    LayoutCachedWidth =1935
                    LayoutCachedHeight =585
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =990
                    Top =900
                    Width =825
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    BackColor =8872517
                    ForeColor =16777215
                    Name ="SalesTax"
                    ControlSource ="=[StateSales]+[CountySales]+[CitySales]"
                    Format ="Percent"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =990
                    LayoutCachedTop =900
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =1185
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =990
                    Top =660
                    Width =825
                    Height =240
                    FontSize =8
                    BorderColor =0
                    Name ="Label69"
                    Caption ="Tax Rate"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =990
                    LayoutCachedTop =660
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =900
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1875
                    Top =900
                    Width =765
                    Height =285
                    FontSize =10
                    TabIndex =10
                    ForeColor =0
                    Name ="StateSalesC"
                    ControlSource ="[StateSales]"
                    Format ="Percent"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1875
                    LayoutCachedTop =900
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1185
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =1875
                    Top =660
                    Width =765
                    Height =240
                    FontSize =8
                    BorderColor =0
                    Name ="Label5"
                    Caption ="State"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1875
                    LayoutCachedTop =660
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =900
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Top =900
                    Width =765
                    Height =285
                    FontSize =10
                    TabIndex =11
                    ForeColor =0
                    Name ="CountySales"
                    ControlSource ="CountySales"
                    Format ="Percent"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2700
                    LayoutCachedTop =900
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =1185
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =2700
                    Top =660
                    Width =765
                    Height =240
                    FontSize =8
                    BorderColor =0
                    Name ="Label7"
                    Caption ="County"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2700
                    LayoutCachedTop =660
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =900
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3525
                    Top =900
                    Width =765
                    Height =285
                    FontSize =10
                    TabIndex =12
                    ForeColor =0
                    Name ="CitySales"
                    ControlSource ="CitySales"
                    Format ="Percent"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3525
                    LayoutCachedTop =900
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1185
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =3525
                    Top =660
                    Width =765
                    Height =240
                    FontSize =8
                    BorderColor =0
                    Name ="Label9"
                    Caption ="City"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3525
                    LayoutCachedTop =660
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =900
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =990
                    Top =1185
                    Width =825
                    Height =0
                    FontSize =10
                    FontWeight =700
                    TabIndex =15
                    BackColor =4754549
                    ForeColor =16777215
                    Name ="AdValorem"
                    ControlSource ="=[StateAdValorem]+[CountyAdValorem]+[CityAdValorem]"
                    Format ="Percent"
                    Tag ="AdValorem"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =990
                    LayoutCachedTop =1185
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =1185
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1875
                    Top =1185
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =16
                    ForeColor =0
                    Name ="StateAdValorem"
                    ControlSource ="[StateAdValorem]"
                    Format ="Percent"
                    Tag ="AdValorem"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1875
                    LayoutCachedTop =1185
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1185
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Top =1185
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =17
                    ForeColor =0
                    Name ="CountyAdValorem"
                    ControlSource ="[CountyAdValorem]"
                    Format ="Percent"
                    Tag ="AdValorem"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2700
                    LayoutCachedTop =1185
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =1185
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3525
                    Top =1185
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =18
                    ForeColor =0
                    Name ="CityAdValorem"
                    ControlSource ="[CityAdValorem]"
                    Format ="Percent"
                    Tag ="AdValorem"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3525
                    LayoutCachedTop =1185
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1185
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =990
                    Top =1185
                    Width =825
                    Height =0
                    FontSize =10
                    FontWeight =700
                    TabIndex =20
                    BackColor =7633277
                    ForeColor =16777215
                    Name ="FHUTTax"
                    ControlSource ="=[StateFHUT]+[CountyFHUT]+[CityFHUT]"
                    Format ="Percent"
                    Tag ="FHUT"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =990
                    LayoutCachedTop =1185
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =1185
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1875
                    Top =1185
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =21
                    ForeColor =0
                    Name ="StateFHUT"
                    ControlSource ="StateFHUT"
                    Format ="Percent"
                    Tag ="FHUT"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1875
                    LayoutCachedTop =1185
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1185
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Top =1185
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =22
                    ForeColor =0
                    Name ="CountyFHUT"
                    ControlSource ="CountyFHUT"
                    Format ="Percent"
                    Tag ="FHUT"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2700
                    LayoutCachedTop =1185
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =1185
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3525
                    Top =1185
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =23
                    ForeColor =0
                    Name ="CityFHUT"
                    ControlSource ="CityFHUT"
                    Format ="Percent"
                    Tag ="FHUT"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3525
                    LayoutCachedTop =1185
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1185
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =990
                    Top =1215
                    Width =825
                    Height =0
                    FontSize =10
                    FontWeight =700
                    TabIndex =25
                    BackColor =9211020
                    ForeColor =16777215
                    Name ="HUTTax"
                    ControlSource ="=[StateHUT]+[CountyHUT]+[CityHUT]"
                    Format ="Percent"
                    Tag ="HUT"
                    GroupTable =37

                    LayoutCachedLeft =990
                    LayoutCachedTop =1215
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =1215
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1875
                    Top =1215
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =26
                    ForeColor =0
                    Name ="StateHUT"
                    ControlSource ="StateHUT"
                    Format ="Percent"
                    Tag ="HUT"
                    GroupTable =37

                    LayoutCachedLeft =1875
                    LayoutCachedTop =1215
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1215
                    RowStart =5
                    RowEnd =5
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Top =1215
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =27
                    ForeColor =0
                    Name ="CountyHUT"
                    ControlSource ="CountyHUT"
                    Format ="Percent"
                    Tag ="HUT"
                    GroupTable =37

                    LayoutCachedLeft =2700
                    LayoutCachedTop =1215
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =1215
                    RowStart =5
                    RowEnd =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3525
                    Top =1215
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =28
                    ForeColor =0
                    Name ="CityHUT"
                    ControlSource ="CityHUT"
                    Format ="Percent"
                    Tag ="HUT"
                    GroupTable =37

                    LayoutCachedLeft =3525
                    LayoutCachedTop =1215
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1215
                    RowStart =5
                    RowEnd =5
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =1875
                    Top =420
                    Width =765
                    Height =240
                    FontSize =8
                    TabIndex =5
                    Name ="EditState"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit this Units ExchangeIT Options"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1875
                    LayoutCachedTop =420
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =660
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =8872517
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =16777215
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    GroupTable =37
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =2700
                    Top =420
                    Width =765
                    Height =240
                    FontSize =8
                    TabIndex =6
                    Name ="EditCounty"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit this Units ExchangeIT Options"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2700
                    LayoutCachedTop =420
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =660
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =8872517
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =16777215
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    GroupTable =37
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3525
                    Top =420
                    Width =765
                    Height =240
                    FontSize =8
                    TabIndex =7
                    Name ="EditCity"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit this Units ExchangeIT Options"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3525
                    LayoutCachedTop =420
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =660
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =8872517
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =16777215
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    GroupTable =37
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1875
                    Top =1275
                    Width =765
                    Height =225
                    FontSize =10
                    TabIndex =29
                    ForeColor =0
                    Name ="StateRReceipt"
                    ControlSource ="StateRReceipt"
                    GroupTable =37

                    LayoutCachedLeft =1875
                    LayoutCachedTop =1275
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1500
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =1275
                            Width =1785
                            Height =225
                            FontSize =8
                            BorderColor =0
                            Name ="Label32"
                            Caption ="Rent Receipt Available? "
                            GroupTable =37
                            LayoutCachedLeft =30
                            LayoutCachedTop =1275
                            LayoutCachedWidth =1815
                            LayoutCachedHeight =1500
                            RowStart =6
                            RowEnd =6
                            ColumnEnd =1
                            LayoutGroup =1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =37
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Top =1275
                    Width =765
                    Height =225
                    FontSize =10
                    TabIndex =30
                    ForeColor =0
                    Name ="CountyRReceipt"
                    ControlSource ="CountyRReceipt"
                    GroupTable =37

                    LayoutCachedLeft =2700
                    LayoutCachedTop =1275
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =1500
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3525
                    Top =1275
                    Width =765
                    Height =225
                    FontSize =10
                    TabIndex =31
                    ForeColor =0
                    Name ="CityRReceipt"
                    ControlSource ="CityRReceipt"
                    GroupTable =37

                    LayoutCachedLeft =3525
                    LayoutCachedTop =1275
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1500
                    RowStart =6
                    RowEnd =6
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1875
                    Top =1560
                    Width =765
                    Height =225
                    FontSize =10
                    TabIndex =32
                    ForeColor =0
                    Name ="StateException"
                    ControlSource ="StateException"
                    Format ="Percent"
                    GroupTable =37

                    LayoutCachedLeft =1875
                    LayoutCachedTop =1560
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1785
                    RowStart =7
                    RowEnd =7
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =1560
                            Width =1785
                            Height =225
                            FontSize =8
                            BorderColor =0
                            Name ="Label36"
                            Caption ="Exception Appliable? "
                            GroupTable =37
                            LayoutCachedLeft =30
                            LayoutCachedTop =1560
                            LayoutCachedWidth =1815
                            LayoutCachedHeight =1785
                            RowStart =7
                            RowEnd =7
                            ColumnEnd =1
                            LayoutGroup =1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =37
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Top =1560
                    Width =765
                    Height =225
                    FontSize =10
                    TabIndex =33
                    ForeColor =0
                    Name ="CountyException"
                    ControlSource ="CountyException"
                    Format ="Percent"
                    GroupTable =37

                    LayoutCachedLeft =2700
                    LayoutCachedTop =1560
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =1785
                    RowStart =7
                    RowEnd =7
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3525
                    Top =1560
                    Width =765
                    Height =225
                    FontSize =10
                    TabIndex =34
                    ForeColor =0
                    Name ="CityException"
                    ControlSource ="CityException"
                    Format ="Percent"
                    GroupTable =37

                    LayoutCachedLeft =3525
                    LayoutCachedTop =1560
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1785
                    RowStart =7
                    RowEnd =7
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =30
                    Top =900
                    Width =900
                    Height =285
                    FontSize =8
                    TabIndex =8
                    BorderColor =0
                    Name ="Label16"
                    ControlSource ="=\"Sales\""
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =30
                    LayoutCachedTop =900
                    LayoutCachedWidth =930
                    LayoutCachedHeight =1185
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =30
                    Top =1185
                    Width =900
                    Height =0
                    FontSize =8
                    TabIndex =14
                    BorderColor =0
                    Name ="Label17"
                    ControlSource ="=\"Ad Valorem\""
                    Tag ="AdValorem"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =30
                    LayoutCachedTop =1185
                    LayoutCachedWidth =930
                    LayoutCachedHeight =1185
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =30
                    Top =1185
                    Width =900
                    Height =0
                    FontSize =8
                    TabIndex =19
                    BorderColor =0
                    Name ="Label22"
                    ControlSource ="=\"FHUT\""
                    Tag ="FHUT"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =30
                    LayoutCachedTop =1185
                    LayoutCachedWidth =930
                    LayoutCachedHeight =1185
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =30
                    Top =1215
                    Width =900
                    Height =0
                    FontSize =8
                    TabIndex =24
                    BorderColor =0
                    Name ="Label27"
                    ControlSource ="=\"HUT\""
                    Tag ="HUT"
                    GroupTable =37

                    LayoutCachedLeft =30
                    LayoutCachedTop =1215
                    LayoutCachedWidth =930
                    LayoutCachedHeight =1215
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =30
                    Top =660
                    Width =900
                    Name ="EmptyCell275"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =30
                    LayoutCachedTop =660
                    LayoutCachedWidth =930
                    LayoutCachedHeight =900
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =30
                    Top =420
                    Width =900
                    Name ="EmptyCell288"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =30
                    LayoutCachedTop =420
                    LayoutCachedWidth =930
                    LayoutCachedHeight =660
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =990
                    Top =420
                    Width =825
                    Name ="EmptyCell289"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =990
                    LayoutCachedTop =420
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =660
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =30
                    Top =1845
                    Width =900
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =35
                    Name ="SaveUpdate"
                    Caption ="Update"
                    OnClick ="[Event Procedure]"
                    GroupTable =37
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000090a468ff90a468e490a46878 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff000000000000000090a468ff90a468f3 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff0000000090a468240000000090a468ff ,
                        0x90a4680990a4682482c2eaff82c2eac682c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff0000000090a468ff90a4683390a468ff ,
                        0x90a4683390a468ff82c2eaff82c2ea5a82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff0000000090a468ff90a468ff90a468ff ,
                        0x90a468ff90a468ff82c2eaff82c2ea0682c2eae782c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff0000000090a4683990a468ff90a468ff ,
                        0x90a468ff90a4683982c2eaff0000000082c2ea7e82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eafc0000000090a4683990a468ff ,
                        0x90a468390000000082c2eaff0000000082c2ea1882c2eaf982c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eafc0000000000000000 ,
                        0x0000000082c2eae782c2eaff00000000727272ff82c2ea6682c2eaf982c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eafc82c2eaae ,
                        0x82c2eaff82c2eab482c2eaff00000000727272ff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000082c2eaff00000000727272ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0xb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff ,
                        0x727272ff0000000082c2ea6300000000727272ffffffffffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffffffffffffb3b3b3ffffffffffffffffffffffffff ,
                        0x727272ff000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffffffffffffb3b3b3ffffffffffffffffffffffffff ,
                        0x727272ff000000000000000000000000b17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4aff000000000000000000000000b17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4aff000000000000000000000000b17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4aff00000000
                    End

                    LayoutCachedLeft =30
                    LayoutCachedTop =1845
                    LayoutCachedWidth =930
                    LayoutCachedHeight =2100
                    PictureCaptionArrangement =5
                    Alignment =1
                    RowStart =8
                    RowEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =10319446
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =8872517
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    GroupTable =37
                    QuickStyle =30
                    QuickStyleMask =-117
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =2685
                    Width =2160
                    Height =105
                    FontSize =8
                    BorderColor =0
                    ForeColor =0
                    Name ="Label119"
                    Caption ="LOC         State   Cnty     City    Act"
                    GridlineColor =0
                    LayoutCachedLeft =2685
                    LayoutCachedWidth =4845
                    LayoutCachedHeight =105
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2685
                    Top =112
                    Width =558
                    Height =156
                    ColumnOrder =0
                    FontSize =8
                    TabIndex =1
                    Name ="LocationID"
                    ControlSource ="LocationID"
                    TopPadding =0

                    LayoutCachedLeft =2685
                    LayoutCachedTop =112
                    LayoutCachedWidth =3243
                    LayoutCachedHeight =268
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3657
                    Top =112
                    Width =432
                    Height =156
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =2
                    Name ="CountyID"
                    ControlSource ="LCountyID"
                    TopPadding =0

                    LayoutCachedLeft =3657
                    LayoutCachedTop =112
                    LayoutCachedWidth =4089
                    LayoutCachedHeight =268
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4521
                    Top =112
                    Width =327
                    Height =156
                    ColumnOrder =3
                    FontSize =8
                    TabIndex =4
                    Name ="TaxTypeID"
                    TopPadding =0

                    LayoutCachedLeft =4521
                    LayoutCachedTop =112
                    LayoutCachedWidth =4848
                    LayoutCachedHeight =268
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4089
                    Top =112
                    Width =432
                    Height =156
                    ColumnOrder =4
                    FontSize =8
                    TabIndex =3
                    Name ="CityID"
                    ControlSource ="LCityID"
                    TopPadding =0

                    LayoutCachedLeft =4089
                    LayoutCachedTop =112
                    LayoutCachedWidth =4521
                    LayoutCachedHeight =268
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    Width =645
                    Height =315
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label31"
                    Caption ="Taxes"
                    FontName ="Segoe UI"
                    LayoutCachedWidth =645
                    LayoutCachedHeight =315
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1980
                    Top =60
                    Width =357
                    Height =201
                    ColumnOrder =1
                    FontSize =8
                    Name ="StateID"
                    ControlSource ="LStateID"
                    TopPadding =0

                    LayoutCachedLeft =1980
                    LayoutCachedTop =60
                    LayoutCachedWidth =2337
                    LayoutCachedHeight =261
                End
                Begin EmptyCell
                    Left =1875
                    Top =1845
                    Width =765
                    Height =255
                    Name ="EmptyCell303"
                    GroupTable =37
                    LayoutCachedLeft =1875
                    LayoutCachedTop =1845
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =2100
                    RowStart =8
                    RowEnd =8
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =2700
                    Top =1845
                    Width =765
                    Height =255
                    Name ="EmptyCell304"
                    GroupTable =37
                    LayoutCachedLeft =2700
                    LayoutCachedTop =1845
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =2100
                    RowStart =8
                    RowEnd =8
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =3525
                    Top =1845
                    Width =765
                    Height =255
                    Name ="EmptyCell305"
                    GroupTable =37
                    LayoutCachedLeft =3525
                    LayoutCachedTop =1845
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =2100
                    RowStart =8
                    RowEnd =8
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =990
                    Top =1845
                    Width =825
                    Height =255
                    Name ="EmptyCell306"
                    GroupTable =37
                    LayoutCachedLeft =990
                    LayoutCachedTop =1845
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =2100
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4350
                    Top =900
                    Width =765
                    Height =285
                    FontSize =10
                    TabIndex =13
                    BackColor =15986395
                    BorderColor =4138256
                    ForeColor =0
                    Name ="OtherTax"
                    Format ="Percent"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4350
                    LayoutCachedTop =900
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =1185
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BackTint =20.0
                    BorderThemeColorIndex =-1
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GroupTable =37
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4350
                    Top =660
                    Width =765
                    Height =240
                    FontSize =8
                    BackColor =9732411
                    BorderColor =0
                    Name ="OtherTaxLabel"
                    Caption ="Other"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4350
                    LayoutCachedTop =660
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =900
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =4350
                    Top =420
                    Width =765
                    Name ="EmptyCell337"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4350
                    LayoutCachedTop =420
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =660
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =4350
                    Top =1185
                    Width =765
                    Height =0
                    Name ="EmptyCell340"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4350
                    LayoutCachedTop =1185
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =1185
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =4350
                    Top =1185
                    Width =765
                    Height =0
                    Name ="EmptyCell341"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4350
                    LayoutCachedTop =1185
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =1185
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =4350
                    Top =1215
                    Width =765
                    Height =0
                    Name ="EmptyCell342"
                    GroupTable =37
                    LayoutCachedLeft =4350
                    LayoutCachedTop =1215
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =1215
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =4350
                    Top =1275
                    Width =765
                    Height =225
                    Name ="EmptyCell343"
                    GroupTable =37
                    LayoutCachedLeft =4350
                    LayoutCachedTop =1275
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =1500
                    RowStart =6
                    RowEnd =6
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =4350
                    Top =1560
                    Width =765
                    Height =225
                    Name ="EmptyCell344"
                    GroupTable =37
                    LayoutCachedLeft =4350
                    LayoutCachedTop =1560
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =1785
                    RowStart =7
                    RowEnd =7
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =4350
                    Top =1845
                    Width =765
                    Height =255
                    Name ="EmptyCell345"
                    GroupTable =37
                    LayoutCachedLeft =4350
                    LayoutCachedTop =1845
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =2100
                    RowStart =8
                    RowEnd =8
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =37
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "LocationTax.cls"
