Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5076
    DatasheetFontHeight =11
    ItemSuffix =14
    Left =2595
    Top =1110
    RecSrcDt = Begin
        0x5991f997ecfbe440
    End
    RecordSource ="SELECT * From vw_LocationTaxes"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
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
            ShowDatePicker =0
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
            KeepTogether = NotDefault
            Height =360
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    Visible = NotDefault
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
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2685
                    Top =112
                    Width =558
                    Height =156
                    ColumnOrder =0
                    FontSize =8
                    Name ="Text2"
                    ControlSource ="LocationID"
                    TopPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2685
                    LayoutCachedTop =112
                    LayoutCachedWidth =3243
                    LayoutCachedHeight =268
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3657
                    Top =112
                    Width =432
                    Height =156
                    ColumnOrder =1
                    FontSize =8
                    TabIndex =1
                    Name ="CountyID"
                    ControlSource ="LCountyID"
                    TopPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =3657
                    LayoutCachedTop =112
                    LayoutCachedWidth =4089
                    LayoutCachedHeight =268
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4521
                    Top =112
                    Width =327
                    Height =156
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =2
                    Name ="TaxTypeID"
                    TopPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =4521
                    LayoutCachedTop =112
                    LayoutCachedWidth =4848
                    LayoutCachedHeight =268
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4089
                    Top =112
                    Width =432
                    Height =156
                    ColumnOrder =3
                    FontSize =8
                    TabIndex =3
                    Name ="CityID"
                    ControlSource ="LCityID"
                    TopPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =4089
                    LayoutCachedTop =112
                    LayoutCachedWidth =4521
                    LayoutCachedHeight =268
                End
                Begin Label
                    Width =4530
                    Height =315
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label31"
                    Caption ="Taxes ──────────────────────●"
                    FontName ="Segoe UI"
                    LayoutCachedWidth =4530
                    LayoutCachedHeight =315
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1980
                    Top =60
                    Width =357
                    Height =201
                    ColumnOrder =4
                    FontSize =8
                    TabIndex =4
                    Name ="StateID"
                    ControlSource ="LStateID"
                    TopPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1980
                    LayoutCachedTop =60
                    LayoutCachedWidth =2337
                    LayoutCachedHeight =261
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =1092
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =996
                    Top =312
                    Width =828
                    Height =288
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BackColor =8872517
                    ForeColor =16777215
                    Name ="SalesTax"
                    ControlSource ="=[StateSales]+[CountySales]+[CitySales]"
                    Format ="Percent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =996
                    LayoutCachedTop =312
                    LayoutCachedWidth =1824
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    Left =996
                    Top =36
                    Width =828
                    Height =240
                    FontSize =8
                    BorderColor =0
                    Name ="Label69"
                    Caption ="Tax Rate"
                    GroupTable =1
                    LayoutCachedLeft =996
                    LayoutCachedTop =36
                    LayoutCachedWidth =1824
                    LayoutCachedHeight =276
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1884
                    Top =312
                    Width =768
                    Height =288
                    FontSize =10
                    TabIndex =3
                    ForeColor =0
                    Name ="StateSalesC"
                    ControlSource ="StateSales"
                    Format ="Percent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1884
                    LayoutCachedTop =312
                    LayoutCachedWidth =2652
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    Left =1884
                    Top =36
                    Width =768
                    Height =240
                    FontSize =8
                    BorderColor =0
                    Name ="Label5"
                    Caption ="State"
                    GroupTable =1
                    LayoutCachedLeft =1884
                    LayoutCachedTop =36
                    LayoutCachedWidth =2652
                    LayoutCachedHeight =276
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2712
                    Top =312
                    Width =768
                    Height =288
                    FontSize =10
                    TabIndex =4
                    ForeColor =0
                    Name ="CountySales"
                    ControlSource ="CountySales"
                    Format ="Percent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2712
                    LayoutCachedTop =312
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    Left =2712
                    Top =36
                    Width =768
                    Height =240
                    FontSize =8
                    BorderColor =0
                    Name ="Label7"
                    Caption ="County"
                    GroupTable =1
                    LayoutCachedLeft =2712
                    LayoutCachedTop =36
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =276
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3516
                    Top =312
                    Width =768
                    Height =288
                    FontSize =10
                    TabIndex =5
                    ForeColor =0
                    Name ="CitySales"
                    ControlSource ="CitySales"
                    Format ="Percent"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =3516
                    LayoutCachedTop =312
                    LayoutCachedWidth =4284
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderShade =50.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    Left =3516
                    Top =36
                    Width =768
                    Height =240
                    FontSize =8
                    BorderColor =0
                    Name ="Label9"
                    Caption ="City"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =3516
                    LayoutCachedTop =36
                    LayoutCachedWidth =4284
                    LayoutCachedHeight =276
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =996
                    Top =600
                    Width =828
                    Height =12
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BackColor =4754549
                    ForeColor =16777215
                    Name ="AdValorem"
                    ControlSource ="=IIf(Nz([StateAdValoremT],0)+Nz([CountyAdValoremT],0)+Nz([CityAdValoremT],0)=0,N"
                        "ull,(Nz([StateAdValoremT],0)+Nz([CountyAdValoremT],0)+Nz([CityAdValoremT],0)))"
                    Format ="Percent"
                    Tag ="AdValorem"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =996
                    LayoutCachedTop =600
                    LayoutCachedWidth =1824
                    LayoutCachedHeight =612
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1884
                    Top =600
                    Width =768
                    Height =12
                    FontSize =10
                    TabIndex =8
                    ForeColor =0
                    Name ="StateAdValoremT"
                    ControlSource ="=IIf([StateAdValorem]=0,Null,[StateAdValorem])"
                    Format ="Percent"
                    Tag ="AdValorem"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1884
                    LayoutCachedTop =600
                    LayoutCachedWidth =2652
                    LayoutCachedHeight =612
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2712
                    Top =600
                    Width =768
                    Height =12
                    FontSize =10
                    TabIndex =9
                    ForeColor =0
                    Name ="CountyAdValoremT"
                    ControlSource ="=IIf([CountyAdValorem]=0,Null,[CountyAdValorem])"
                    Format ="Percent"
                    Tag ="AdValorem"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2712
                    LayoutCachedTop =600
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =612
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3516
                    Top =600
                    Width =768
                    Height =12
                    FontSize =10
                    TabIndex =10
                    ForeColor =0
                    Name ="CityAdValoremT"
                    ControlSource ="=IIf([CityAdValorem]=0,Null,[CityAdValorem])"
                    Format ="Percent"
                    Tag ="AdValorem"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =3516
                    LayoutCachedTop =600
                    LayoutCachedWidth =4284
                    LayoutCachedHeight =612
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =996
                    Top =612
                    Width =828
                    Height =12
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    BackColor =7633277
                    ForeColor =16777215
                    Name ="FHUTTax"
                    ControlSource ="=IIf(Nz([StateFHUT],0)+Nz([CountyFHUT],0)+Nz([CityFHUT],0)=0,Null,Nz([StateFHUT]"
                        ",0)+Nz([CountyFHUT],0)+Nz([CityFHUT],0))"
                    Format ="Percent"
                    Tag ="FHUT"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =996
                    LayoutCachedTop =612
                    LayoutCachedWidth =1824
                    LayoutCachedHeight =624
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1884
                    Top =612
                    Width =768
                    Height =12
                    FontSize =10
                    TabIndex =13
                    ForeColor =0
                    Name ="StateFHUTt"
                    ControlSource ="=IIf(Nz([FHUTTax],\"\")=\"\",Null,[StateFHUT])"
                    Format ="Percent"
                    Tag ="FHUT"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1884
                    LayoutCachedTop =612
                    LayoutCachedWidth =2652
                    LayoutCachedHeight =624
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2712
                    Top =612
                    Width =768
                    Height =12
                    FontSize =10
                    TabIndex =14
                    ForeColor =0
                    Name ="CountyFHUTt"
                    ControlSource ="=IIf(Nz([FHUTTax],\"\")=\"\",Null,[CountyFHUT])"
                    Format ="Percent"
                    Tag ="FHUT"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2712
                    LayoutCachedTop =612
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =624
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3516
                    Top =612
                    Width =768
                    Height =12
                    FontSize =10
                    TabIndex =15
                    ForeColor =0
                    Name ="CityFHUTt"
                    ControlSource ="=IIf(Nz([FHUTTax],\"\")=\"\",Null,[CityFHUT])"
                    Format ="Percent"
                    Tag ="FHUT"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =3516
                    LayoutCachedTop =612
                    LayoutCachedWidth =4284
                    LayoutCachedHeight =624
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =996
                    Top =624
                    Width =828
                    Height =12
                    FontSize =10
                    FontWeight =700
                    TabIndex =17
                    BackColor =9211020
                    ForeColor =16777215
                    Name ="HUTTax"
                    ControlSource ="=IIf(Nz([StateHUT],0)+Nz([CountyHUT],0)+Nz([CityHUT],0)=0,Null,Nz([StateHUT],0)+"
                        "Nz([CountyHUT],0)+Nz([CityHUT],0))"
                    Format ="Percent"
                    Tag ="HUT"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =996
                    LayoutCachedTop =624
                    LayoutCachedWidth =1824
                    LayoutCachedHeight =636
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1884
                    Top =624
                    Width =768
                    Height =12
                    FontSize =10
                    TabIndex =18
                    ForeColor =0
                    Name ="StateHUTt"
                    ControlSource ="=IIf(Nz([HUTTax],\"\")=\"\",Null,[StateHUT])"
                    Format ="Percent"
                    Tag ="HUT"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1884
                    LayoutCachedTop =624
                    LayoutCachedWidth =2652
                    LayoutCachedHeight =636
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2712
                    Top =624
                    Width =768
                    Height =12
                    FontSize =10
                    TabIndex =19
                    ForeColor =0
                    Name ="CountyHUTt"
                    ControlSource ="=IIf(Nz([HUTTax],\"\")=\"\",Null,[CountyHUT])"
                    Format ="Percent"
                    Tag ="HUT"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2712
                    LayoutCachedTop =624
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =636
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =3
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3516
                    Top =624
                    Width =768
                    Height =12
                    FontSize =10
                    TabIndex =20
                    ForeColor =0
                    Name ="CityHUTt"
                    ControlSource ="=IIf(Nz([HUTTax],\"\")=\"\",Null,[CityHUT])"
                    Format ="Percent"
                    Tag ="HUT"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =3516
                    LayoutCachedTop =624
                    LayoutCachedWidth =4284
                    LayoutCachedHeight =636
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
                    GroupTable =1
                End
                Begin CommandButton
                    DisplayWhen =2
                    Left =4296
                    Top =36
                    Width =768
                    Height =240
                    FontSize =8
                    Name ="EditCity"
                    Caption ="Edit"
                    ControlTipText ="Add/Edit this Units ExchangeIT Options"
                    GroupTable =1
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =4296
                    LayoutCachedTop =36
                    LayoutCachedWidth =5064
                    LayoutCachedHeight =276
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =8872517
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderWidth =1
                    BorderColor =0
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
                    GroupTable =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1884
                    Top =636
                    Width =768
                    Height =228
                    FontSize =10
                    TabIndex =21
                    ForeColor =0
                    Name ="StateRReceipt"
                    ControlSource ="StateRReceipt"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1884
                    LayoutCachedTop =636
                    LayoutCachedWidth =2652
                    LayoutCachedHeight =864
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
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =4
                            TextAlign =3
                            Left =36
                            Top =636
                            Width =1788
                            Height =228
                            FontSize =8
                            BorderColor =0
                            Name ="Label32"
                            Caption ="Rent Receipt Available? "
                            GroupTable =1
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =36
                            LayoutCachedTop =636
                            LayoutCachedWidth =1824
                            LayoutCachedHeight =864
                            RowStart =5
                            RowEnd =5
                            ColumnEnd =1
                            LayoutGroup =1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2712
                    Top =636
                    Width =768
                    Height =228
                    FontSize =10
                    TabIndex =22
                    ForeColor =0
                    Name ="CountyRReceipt"
                    ControlSource ="CountyRReceipt"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2712
                    LayoutCachedTop =636
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =864
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3516
                    Top =636
                    Width =768
                    Height =228
                    FontSize =10
                    TabIndex =23
                    ForeColor =0
                    Name ="CityRReceipt"
                    ControlSource ="CityRReceipt"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =3516
                    LayoutCachedTop =636
                    LayoutCachedWidth =4284
                    LayoutCachedHeight =864
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1884
                    Top =864
                    Width =768
                    Height =228
                    FontSize =10
                    TabIndex =24
                    ForeColor =0
                    Name ="StateException"
                    ControlSource ="StateException"
                    Format ="Percent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1884
                    LayoutCachedTop =864
                    LayoutCachedWidth =2652
                    LayoutCachedHeight =1092
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
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =4
                            TextAlign =3
                            Left =36
                            Top =864
                            Width =1788
                            Height =228
                            FontSize =8
                            BorderColor =0
                            Name ="Label36"
                            Caption ="Exception Appliable? "
                            GroupTable =1
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =36
                            LayoutCachedTop =864
                            LayoutCachedWidth =1824
                            LayoutCachedHeight =1092
                            RowStart =6
                            RowEnd =6
                            ColumnEnd =1
                            LayoutGroup =1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2712
                    Top =864
                    Width =768
                    Height =228
                    FontSize =10
                    TabIndex =25
                    ForeColor =0
                    Name ="CountyException"
                    ControlSource ="CountyException"
                    Format ="Percent"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2712
                    LayoutCachedTop =864
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =1092
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
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3516
                    Top =864
                    Width =768
                    Height =228
                    FontSize =10
                    TabIndex =26
                    ForeColor =0
                    Name ="CityException"
                    ControlSource ="CityException"
                    Format ="Percent"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =3516
                    LayoutCachedTop =864
                    LayoutCachedWidth =4284
                    LayoutCachedHeight =1092
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
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =312
                    Width =900
                    Height =288
                    FontSize =8
                    TabIndex =1
                    BorderColor =0
                    Name ="Label16"
                    ControlSource ="=\"Sales\""
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =36
                    LayoutCachedTop =312
                    LayoutCachedWidth =936
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =600
                    Width =900
                    Height =12
                    FontSize =8
                    TabIndex =6
                    BorderColor =0
                    Name ="Label17"
                    ControlSource ="=IIf(Nz([AdValorem],\"\")=\"\",Null,\"Ad Valorem\")"
                    Tag ="AdValorem"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =36
                    LayoutCachedTop =600
                    LayoutCachedWidth =936
                    LayoutCachedHeight =612
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =612
                    Width =900
                    Height =12
                    FontSize =8
                    TabIndex =11
                    BorderColor =0
                    Name ="Label22"
                    ControlSource ="=IIf(Nz([FHUTTax],\"\")=\"\",Null,\"FHUT\")"
                    Tag ="FHUT"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =36
                    LayoutCachedTop =612
                    LayoutCachedWidth =936
                    LayoutCachedHeight =624
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =4
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =36
                    Top =624
                    Width =900
                    Height =12
                    FontSize =8
                    TabIndex =16
                    BorderColor =0
                    Name ="Label27"
                    ControlSource ="=IIf(Nz([HUTTax],\"\")=\"\",Null,\"HUT\")"
                    Tag ="HUT"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =36
                    LayoutCachedTop =624
                    LayoutCachedWidth =936
                    LayoutCachedHeight =636
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =36
                    Top =36
                    Width =900
                    Name ="EmptyCell275"
                    GroupTable =1
                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =936
                    LayoutCachedHeight =276
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4296
                    Top =312
                    Width =768
                    Height =288
                    Name ="EmptyCell6"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4296
                    LayoutCachedTop =312
                    LayoutCachedWidth =5064
                    LayoutCachedHeight =600
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4296
                    Top =600
                    Width =768
                    Height =12
                    Name ="EmptyCell7"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4296
                    LayoutCachedTop =600
                    LayoutCachedWidth =5064
                    LayoutCachedHeight =612
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4296
                    Top =612
                    Width =768
                    Height =12
                    Name ="EmptyCell8"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4296
                    LayoutCachedTop =612
                    LayoutCachedWidth =5064
                    LayoutCachedHeight =624
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4296
                    Top =624
                    Width =768
                    Height =12
                    Name ="EmptyCell9"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4296
                    LayoutCachedTop =624
                    LayoutCachedWidth =5064
                    LayoutCachedHeight =636
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4296
                    Top =636
                    Width =768
                    Height =228
                    Name ="EmptyCell10"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4296
                    LayoutCachedTop =636
                    LayoutCachedWidth =5064
                    LayoutCachedHeight =864
                    RowStart =5
                    RowEnd =5
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4296
                    Top =864
                    Width =768
                    Height =228
                    Name ="EmptyCell11"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4296
                    LayoutCachedTop =864
                    LayoutCachedWidth =5064
                    LayoutCachedHeight =1092
                    RowStart =6
                    RowEnd =6
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
