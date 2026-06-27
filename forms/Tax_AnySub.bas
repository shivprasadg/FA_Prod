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
    Width =13536
    DatasheetFontHeight =11
    ItemSuffix =217
    Left =2970
    Top =6585
    Right =16245
    Bottom =8655
    RecSrcDt = Begin
        0xff0d3debac01e540
    End
    RecordSource ="Select * From vw_Tax_CountyBaseQ WHERE TaxCountyID=424 Order By StartsOn Desc;"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
            Height =270
            BackColor =15527148
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =10095
                    Width =1080
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label3"
                    Caption ="Starts ON"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =10095
                    LayoutCachedWidth =11175
                    LayoutCachedHeight =225
                    ColumnStart =5
                    ColumnEnd =5
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =10965
                    Top =60
                    Width =120
                    Height =210
                    ColumnOrder =0
                    FontSize =8
                    BorderColor =12632256
                    Name ="txtCurrent"

                    LayoutCachedLeft =10965
                    LayoutCachedTop =60
                    LayoutCachedWidth =11085
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
                    OverlapFlags =85
                    TextAlign =2
                    Left =12525
                    Width =945
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label169"
                    Caption ="\015\012Active"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =12525
                    LayoutCachedWidth =13470
                    LayoutCachedHeight =225
                    ColumnStart =7
                    ColumnEnd =9
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =8655
                    Width =1035
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label206"
                    Caption ="Tax Rate %"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =8655
                    LayoutCachedWidth =9690
                    LayoutCachedHeight =225
                    ColumnStart =4
                    ColumnEnd =4
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =11370
                    Width =1080
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label210"
                    Caption ="Ends ON"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =11370
                    LayoutCachedWidth =12450
                    LayoutCachedHeight =225
                    ColumnStart =5
                    ColumnEnd =5
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =1785
                    Width =6825
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label213"
                    Caption ="Tax Rule / Description"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =1785
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =225
                    ColumnStart =3
                    ColumnEnd =3
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =60
                    Width =1725
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label216"
                    Caption ="Tax Type"
                    LeftPadding =0
                    RightPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =225
                    ColumnStart =3
                    ColumnEnd =3
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =493
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10095
                    Width =1080
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    BorderColor =10921638
                    Name ="EventType"
                    ControlSource ="StartsOn"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =10095
                    LayoutCachedWidth =11175
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Width =1725
                    Height =255
                    FontSize =10
                    TabIndex =1
                    BorderColor =10921638
                    Name ="TaxType"
                    ControlSource ="TaxType"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12525
                    Width =273
                    Height =231
                    FontSize =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Id"
                    ControlSource ="Id"
                    TopPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =12525
                    LayoutCachedWidth =12798
                    LayoutCachedHeight =231
                    BackThemeColorIndex =-1
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =12930
                    Top =50
                    Width =390
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

                    LayoutCachedLeft =12930
                    LayoutCachedTop =50
                    LayoutCachedWidth =13320
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
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13260
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

                    LayoutCachedLeft =13260
                    LayoutCachedWidth =13533
                    LayoutCachedHeight =231
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1785
                    Width =6825
                    Height =465
                    FontSize =10
                    TabIndex =7
                    BorderColor =10921638
                    Name ="Description"
                    ControlSource ="=IIf(Nz([TaxScript],\"\")=\"\",[TaxDescription],[TaxScript])"
                    OnGotFocus ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedLeft =1785
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =465
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11370
                    Width =1080
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =8
                    BorderColor =10921638
                    Name ="Text211"
                    ControlSource ="EndsOn"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =11370
                    LayoutCachedWidth =12450
                    LayoutCachedHeight =255
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Top =450
                    Width =13536
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
                    HorizontalAnchor =2

                    LayoutCachedTop =450
                    LayoutCachedWidth =13536
                    LayoutCachedHeight =493
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
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8655
                    Width =1030
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
                    ControlTipText ="The percentage of the Tax Rate"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =8655
                    LayoutCachedWidth =9685
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
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =0
            BackColor =14211288
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
' See "Tax_AnySub.cls"
