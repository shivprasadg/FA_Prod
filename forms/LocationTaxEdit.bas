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
    Width =5055
    DatasheetFontHeight =11
    ItemSuffix =309
    Left =3615
    Top =3045
    Right =8685
    Bottom =5880
    RecSrcDt = Begin
        0x024d5d6f7307e540
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
            Height =3015
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
                    GridlineColor =10921638
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
                    BorderColor =10921638
                    Name ="Box298"
                    GridlineColor =10921638
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
                    Top =915
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
                    GridlineColor =10921638

                    LayoutCachedLeft =990
                    LayoutCachedTop =915
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =1200
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
                    Top =675
                    Width =825
                    Height =240
                    FontSize =8
                    Name ="Label69"
                    Caption ="Tax Rate"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =990
                    LayoutCachedTop =675
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =915
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
                    Top =915
                    Width =765
                    Height =285
                    FontSize =10
                    TabIndex =10
                    Name ="StateSalesC"
                    ControlSource ="[StateSales]"
                    Format ="Percent"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1875
                    LayoutCachedTop =915
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1200
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
                    OverlapFlags =85
                    TextAlign =2
                    Left =1875
                    Top =675
                    Width =765
                    Height =240
                    FontSize =8
                    Name ="Label5"
                    Caption ="State"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =1875
                    LayoutCachedTop =675
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =915
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
                    Top =915
                    Width =765
                    Height =285
                    FontSize =10
                    TabIndex =11
                    Name ="CountySales"
                    ControlSource ="CountySales"
                    Format ="Percent"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =2700
                    LayoutCachedTop =915
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =1200
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
                    OverlapFlags =85
                    TextAlign =2
                    Left =2700
                    Top =675
                    Width =765
                    Height =240
                    FontSize =8
                    Name ="Label7"
                    Caption ="County"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2700
                    LayoutCachedTop =675
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =915
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
                    Top =915
                    Width =765
                    Height =285
                    FontSize =10
                    TabIndex =12
                    Name ="CitySales"
                    ControlSource ="CitySales"
                    Format ="Percent"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =3525
                    LayoutCachedTop =915
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1200
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
                    OverlapFlags =85
                    TextAlign =2
                    Left =3525
                    Top =675
                    Width =765
                    Height =240
                    FontSize =8
                    Name ="Label9"
                    Caption ="City"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =3525
                    LayoutCachedTop =675
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =915
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
                    Top =1200
                    Width =825
                    Height =0
                    FontSize =10
                    FontWeight =700
                    TabIndex =14
                    BackColor =4754549
                    ForeColor =16777215
                    Name ="AdValorem"
                    ControlSource ="=[StateAdValorem]+[CountyAdValorem]+[CityAdValorem]"
                    Format ="Percent"
                    Tag ="AdValorem"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =990
                    LayoutCachedTop =1200
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =1200
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
                    Top =1200
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =15
                    Name ="StateAdValorem"
                    ControlSource ="[StateAdValorem]"
                    Format ="Percent"
                    Tag ="AdValorem"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1875
                    LayoutCachedTop =1200
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1200
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
                    Top =1200
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =16
                    Name ="CountyAdValorem"
                    ControlSource ="[CountyAdValorem]"
                    Format ="Percent"
                    Tag ="AdValorem"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =2700
                    LayoutCachedTop =1200
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =1200
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
                    Top =1200
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =17
                    Name ="CityAdValorem"
                    ControlSource ="[CityAdValorem]"
                    Format ="Percent"
                    Tag ="AdValorem"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =3525
                    LayoutCachedTop =1200
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1200
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
                    Top =1200
                    Width =825
                    Height =0
                    FontSize =10
                    FontWeight =700
                    TabIndex =19
                    BackColor =7633277
                    ForeColor =16777215
                    Name ="FHUTTax"
                    ControlSource ="=[StateFHUT]+[CountyFHUT]+[CityFHUT]"
                    Format ="Percent"
                    Tag ="FHUT"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =990
                    LayoutCachedTop =1200
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =1200
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
                    Top =1200
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =20
                    Name ="StateFHUT"
                    ControlSource ="StateFHUT"
                    Format ="Percent"
                    Tag ="FHUT"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1875
                    LayoutCachedTop =1200
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1200
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
                    Top =1200
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =21
                    Name ="CountyFHUT"
                    ControlSource ="CountyFHUT"
                    Format ="Percent"
                    Tag ="FHUT"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =2700
                    LayoutCachedTop =1200
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =1200
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
                    Top =1200
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =22
                    Name ="CityFHUT"
                    ControlSource ="CityFHUT"
                    Format ="Percent"
                    Tag ="FHUT"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =3525
                    LayoutCachedTop =1200
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1200
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
                    Top =1230
                    Width =825
                    Height =0
                    FontSize =10
                    FontWeight =700
                    TabIndex =24
                    BackColor =9211020
                    ForeColor =16777215
                    Name ="HUTTax"
                    ControlSource ="=[StateHUT]+[CountyHUT]+[CityHUT]"
                    Format ="Percent"
                    Tag ="HUT"
                    GroupTable =37
                    GridlineColor =10921638

                    LayoutCachedLeft =990
                    LayoutCachedTop =1230
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =1230
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
                    Top =1230
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =25
                    Name ="StateHUT"
                    ControlSource ="StateHUT"
                    Format ="Percent"
                    Tag ="HUT"
                    GroupTable =37
                    GridlineColor =10921638

                    LayoutCachedLeft =1875
                    LayoutCachedTop =1230
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1230
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
                    Top =1230
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =26
                    Name ="CountyHUT"
                    ControlSource ="CountyHUT"
                    Format ="Percent"
                    Tag ="HUT"
                    GroupTable =37
                    GridlineColor =10921638

                    LayoutCachedLeft =2700
                    LayoutCachedTop =1230
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =1230
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
                    Top =1230
                    Width =765
                    Height =0
                    FontSize =10
                    TabIndex =27
                    Name ="CityHUT"
                    ControlSource ="CityHUT"
                    Format ="Percent"
                    Tag ="HUT"
                    GroupTable =37
                    GridlineColor =10921638

                    LayoutCachedLeft =3525
                    LayoutCachedTop =1230
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1230
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
                    ForeColor =8210719
                    Name ="EditState"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit this State's taxes"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =15
                    GridlineColor =10921638
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

                    LayoutCachedLeft =1875
                    LayoutCachedTop =420
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =660
                    PictureCaptionArrangement =5
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =8210719
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =8210719
                    PressedForeThemeColorIndex =-1
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
                    ForeColor =8210719
                    Name ="EditCounty"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit this County's taxes"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =15
                    GridlineColor =10921638
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

                    LayoutCachedLeft =2700
                    LayoutCachedTop =420
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =660
                    PictureCaptionArrangement =5
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =8210719
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =8210719
                    PressedForeThemeColorIndex =-1
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
                    ForeColor =8210719
                    Name ="EditCity"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit this City's taxes"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =15
                    GridlineColor =10921638
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

                    LayoutCachedLeft =3525
                    LayoutCachedTop =420
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =660
                    PictureCaptionArrangement =5
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =8210719
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =8210719
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    GroupTable =37
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =30
                    Top =1290
                    Width =1785
                    Height =225
                    FontSize =8
                    Name ="Label32"
                    Caption ="Rent Receipt Available? "
                    GroupTable =37
                    GridlineColor =10921638
                    LayoutCachedLeft =30
                    LayoutCachedTop =1290
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =1515
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
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1875
                    Top =1290
                    Width =765
                    Height =225
                    FontSize =10
                    TabIndex =28
                    Name ="StateRReceipt"
                    ControlSource ="StateRReceipt"
                    GroupTable =37
                    GridlineColor =10921638

                    LayoutCachedLeft =1875
                    LayoutCachedTop =1290
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1515
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
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Top =1290
                    Width =765
                    Height =225
                    FontSize =10
                    TabIndex =29
                    Name ="CountyRReceipt"
                    ControlSource ="CountyRReceipt"
                    GroupTable =37
                    GridlineColor =10921638

                    LayoutCachedLeft =2700
                    LayoutCachedTop =1290
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =1515
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
                    Top =1290
                    Width =765
                    Height =225
                    FontSize =10
                    TabIndex =30
                    Name ="CityRReceipt"
                    ControlSource ="CityRReceipt"
                    GroupTable =37
                    GridlineColor =10921638

                    LayoutCachedLeft =3525
                    LayoutCachedTop =1290
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1515
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
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =30
                    Top =1575
                    Width =1785
                    Height =225
                    FontSize =8
                    Name ="Label36"
                    Caption ="Exception Appliable? "
                    GroupTable =37
                    GridlineColor =10921638
                    LayoutCachedLeft =30
                    LayoutCachedTop =1575
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =1800
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
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1875
                    Top =1575
                    Width =765
                    Height =225
                    FontSize =10
                    TabIndex =31
                    Name ="StateException"
                    ControlSource ="StateException"
                    Format ="Percent"
                    GroupTable =37
                    GridlineColor =10921638

                    LayoutCachedLeft =1875
                    LayoutCachedTop =1575
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1800
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
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Top =1575
                    Width =765
                    Height =225
                    FontSize =10
                    TabIndex =32
                    Name ="CountyException"
                    ControlSource ="CountyException"
                    Format ="Percent"
                    GroupTable =37
                    GridlineColor =10921638

                    LayoutCachedLeft =2700
                    LayoutCachedTop =1575
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =1800
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
                    Top =1575
                    Width =765
                    Height =225
                    FontSize =10
                    TabIndex =33
                    Name ="CityException"
                    ControlSource ="CityException"
                    Format ="Percent"
                    GroupTable =37
                    GridlineColor =10921638

                    LayoutCachedLeft =3525
                    LayoutCachedTop =1575
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1800
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
                    Top =915
                    Width =900
                    Height =285
                    FontSize =8
                    TabIndex =8
                    Name ="Label16"
                    ControlSource ="=\"Sales\""
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =30
                    LayoutCachedTop =915
                    LayoutCachedWidth =930
                    LayoutCachedHeight =1200
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
                    Top =1200
                    Width =900
                    Height =0
                    FontSize =8
                    TabIndex =13
                    Name ="Label17"
                    ControlSource ="=\"Ad Valorem\""
                    Tag ="AdValorem"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =30
                    LayoutCachedTop =1200
                    LayoutCachedWidth =930
                    LayoutCachedHeight =1200
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
                    Top =1200
                    Width =900
                    Height =0
                    FontSize =8
                    TabIndex =18
                    Name ="Label22"
                    ControlSource ="=\"FHUT\""
                    Tag ="FHUT"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =30
                    LayoutCachedTop =1200
                    LayoutCachedWidth =930
                    LayoutCachedHeight =1200
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
                    Top =1230
                    Width =900
                    Height =0
                    FontSize =8
                    TabIndex =23
                    Name ="Label27"
                    ControlSource ="=\"HUT\""
                    Tag ="HUT"
                    GroupTable =37
                    GridlineColor =10921638

                    LayoutCachedLeft =30
                    LayoutCachedTop =1230
                    LayoutCachedWidth =930
                    LayoutCachedHeight =1230
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
                    Top =675
                    Width =900
                    Name ="EmptyCell275"
                    GroupTable =37
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =30
                    LayoutCachedTop =675
                    LayoutCachedWidth =930
                    LayoutCachedHeight =915
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
                    BottomPadding =15
                    GridlineColor =10921638
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
                    BottomPadding =15
                    GridlineColor =10921638
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
                    Top =1860
                    Width =900
                    Height =255
                    FontSize =8
                    TabIndex =34
                    ForeColor =8210719
                    Name ="SaveUpdate"
                    Caption ="Update"
                    OnClick ="[Event Procedure]"
                    GroupTable =37
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x000000000000000090a4682190a468ae90a468e790a468e490a468a890a4681b ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff00000000 ,
                        0x90a468ff90a4682a90a468f390a468ff90a468d590a468d890a468ff90a468ed ,
                        0x90a4682100000000727272ffffffffff727272ffffffffffffffffff00000000 ,
                        0x90a468ff90a468cc90a468f390a4683f000000000000000090a4684890a468f6 ,
                        0x90a468bd00000000727272ffffffffff727272ffffffffffffffffff00000000 ,
                        0x90a468ff90a468ff90a468720000000000000000000000000000000090a46878 ,
                        0x90a468ff90a4682a727272ffffffffff727272ffffffffffffffffff00000000 ,
                        0x90a468ff90a468ff90a468ff90a468ff00000000000000000000000000000000 ,
                        0x0000000000000000727272ffffffffff727272ffffffffffffffffffffffff18 ,
                        0x00000000000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff90a468ff727272ffffffffff727272ffffffffffffffffffffffff4b ,
                        0x90a4683690a468ff90a468960000000000000000000000000000000090a46890 ,
                        0x90a468ff90a468ff727272ffffffffff727272ffffffffffffffffffffffffa5 ,
                        0x0000000090a468ae90a468ff90a4689390a4682a90a4682a90a4689390a468ff ,
                        0x90a468b490a468ff727272ffffffffff727272fffffffffffffffffffffffffc ,
                        0xffffff4290a4681290a468c990a468ff90a468ff90a468ff90a468ff90a468c9 ,
                        0x90a4681590a468ff727272ffffffffff727272ffffffffffffffffffffffffff ,
                        0xffffffedffffff5190a4680390a4685a90a4689c90a4689f90a4685d90a46803 ,
                        0x0000000000000000727272ffffffffff727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffab727272ff727272ff727272ff727272ff72727203 ,
                        0x0000000000000000727272ffffffffff727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffff727272ffdcdcdcff777777f67272723f00000000 ,
                        0x0000000000000000727272ffffffffff727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffff727272ff787878f47272723c0000000000000000 ,
                        0x0000000000000000727272ffffffffff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff7272723c000000000000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =30
                    LayoutCachedTop =1860
                    LayoutCachedWidth =930
                    LayoutCachedHeight =2115
                    PictureCaptionArrangement =5
                    RowStart =8
                    RowEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =8210719
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13017476
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10319446
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =8210719
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    GroupTable =37
                    Overlaps =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =2685
                    Width =2160
                    Height =105
                    FontSize =8
                    Name ="Label119"
                    Caption ="LOC         State   Cnty     City    Act"
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
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="LocationID"
                    ControlSource ="LocationID"
                    TopPadding =0
                    GridlineColor =10921638

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
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="CountyID"
                    ControlSource ="LCountyID"
                    TopPadding =0
                    GridlineColor =10921638

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
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TaxTypeID"
                    TopPadding =0
                    GridlineColor =10921638

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
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="CityID"
                    ControlSource ="LCityID"
                    TopPadding =0
                    GridlineColor =10921638

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
                    BorderColor =8355711
                    ForeColor =8210719
                    Name ="Label31"
                    Caption ="Taxes"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
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
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="StateID"
                    ControlSource ="LStateID"
                    TopPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =1980
                    LayoutCachedTop =60
                    LayoutCachedWidth =2337
                    LayoutCachedHeight =261
                End
                Begin EmptyCell
                    Left =1875
                    Top =1860
                    Width =765
                    Height =255
                    Name ="EmptyCell303"
                    GroupTable =37
                    GridlineColor =10921638
                    LayoutCachedLeft =1875
                    LayoutCachedTop =1860
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =2115
                    RowStart =8
                    RowEnd =8
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =2700
                    Top =1860
                    Width =765
                    Height =255
                    Name ="EmptyCell304"
                    GroupTable =37
                    GridlineColor =10921638
                    LayoutCachedLeft =2700
                    LayoutCachedTop =1860
                    LayoutCachedWidth =3465
                    LayoutCachedHeight =2115
                    RowStart =8
                    RowEnd =8
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =3525
                    Top =1860
                    Width =765
                    Height =255
                    Name ="EmptyCell305"
                    GroupTable =37
                    GridlineColor =10921638
                    LayoutCachedLeft =3525
                    LayoutCachedTop =1860
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =2115
                    RowStart =8
                    RowEnd =8
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =37
                End
                Begin EmptyCell
                    Left =990
                    Top =1860
                    Width =825
                    Height =255
                    Name ="EmptyCell306"
                    GroupTable =37
                    GridlineColor =10921638
                    LayoutCachedLeft =990
                    LayoutCachedTop =1860
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =2115
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =1
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
' See "LocationTaxEdit.cls"
