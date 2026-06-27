Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18360
    DatasheetFontHeight =11
    ItemSuffix =496
    Left =4935
    Top =2715
    Right =21615
    Bottom =14760
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xb002067d8f00e540
    End
    RecordSource ="TaxExceptionsTemp"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
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
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin ToggleButton
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
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
            PressedForeThemeColorIndex =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =600
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureType =2
                    Top =465
                    Width =18360
                    Height =135
                    Name ="Image142"
                    Picture ="2_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =465
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =600
                    TabIndex =13
                End
                Begin Label
                    OverlapFlags =255
                    Left =120
                    Width =2295
                    Height =495
                    FontSize =20
                    FontWeight =700
                    Name ="Label105"
                    Caption ="Exceptions:"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =2415
                    LayoutCachedHeight =495
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8688
                    Width =408
                    Height =231
                    ColumnOrder =11
                    FontSize =8
                    Name ="TaxLevelLookupID"
                    TopPadding =0

                    LayoutCachedLeft =8688
                    LayoutCachedWidth =9096
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8220
                    Width =408
                    Height =231
                    ColumnOrder =9
                    FontSize =8
                    TabIndex =1
                    Name ="TaxTypeID2"
                    ControlSource ="TaxTypeID"
                    TopPadding =0

                    LayoutCachedLeft =8220
                    LayoutCachedWidth =8628
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2415
                    Width =5775
                    Height =495
                    ColumnOrder =10
                    FontSize =20
                    FontWeight =700
                    TabIndex =2
                    BorderColor =0
                    ForeColor =12349952
                    Name ="TaxLevel"
                    FontName ="Century Gothic"

                    LayoutCachedLeft =2415
                    LayoutCachedWidth =8190
                    LayoutCachedHeight =495
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9180
                    Width =408
                    Height =231
                    ColumnOrder =7
                    FontSize =8
                    TabIndex =3
                    Name ="TaxExceptionTypeID"
                    ControlSource ="TaxExceptionTypeID"
                    TopPadding =0

                    LayoutCachedLeft =9180
                    LayoutCachedWidth =9588
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9585
                    Width =408
                    Height =231
                    ColumnOrder =8
                    FontSize =8
                    TabIndex =4
                    Name ="ExceptionGroup"
                    TopPadding =0

                    LayoutCachedLeft =9585
                    LayoutCachedWidth =9993
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14844
                    Width =408
                    Height =231
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =5
                    Name ="ID"
                    ControlSource ="TaxExceptionID"
                    TopPadding =0

                    LayoutCachedLeft =14844
                    LayoutCachedWidth =15252
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13620
                    Width =408
                    Height =231
                    ColumnOrder =3
                    FontSize =8
                    TabIndex =6
                    Name ="TaxStatesID"
                    ControlSource ="TaxStatesID"
                    TopPadding =0

                    LayoutCachedLeft =13620
                    LayoutCachedWidth =14028
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14028
                    Width =408
                    Height =231
                    ColumnOrder =4
                    FontSize =8
                    TabIndex =7
                    Name ="TaxCountiesID"
                    ControlSource ="TaxCountiesID"
                    TopPadding =0

                    LayoutCachedLeft =14028
                    LayoutCachedWidth =14436
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14436
                    Width =408
                    Height =231
                    ColumnOrder =5
                    FontSize =8
                    TabIndex =8
                    Name ="TaxCitiesID"
                    ControlSource ="TaxCitiesID"
                    TopPadding =0

                    LayoutCachedLeft =14436
                    LayoutCachedWidth =14844
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12585
                    Width =408
                    Height =231
                    ColumnOrder =6
                    FontSize =8
                    TabIndex =9
                    Name ="Mode"
                    ControlSource ="Mode"
                    TopPadding =0

                    LayoutCachedLeft =12585
                    LayoutCachedWidth =12993
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12993
                    Width =408
                    Height =231
                    ColumnOrder =0
                    FontSize =8
                    TabIndex =10
                    Name ="TaxTypeID"
                    ControlSource ="TaxTypeID"
                    ValidationRule ="Is Not Null"

                    LayoutCachedLeft =12993
                    LayoutCachedWidth =13401
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =85.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =16920
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =11
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="FAMainMenu"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"GoHome\" xmlns=\"http://schemas.microsoft.com/office/accesss"
                                "ervices/2009/11/application\"><Statements><Action Name=\"OpenForm\"><Argument Na"
                                "me=\"FormName\">FAMainMenu</Argum"
                        End
                        Begin
                            Comment ="_AXL:ent></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =16920
                    LayoutCachedTop =120
                    LayoutCachedWidth =17475
                    LayoutCachedHeight =348
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    PictureType =1
                    Left =17487
                    Top =120
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =12
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1

                    LayoutCachedLeft =17487
                    LayoutCachedTop =120
                    LayoutCachedWidth =18042
                    LayoutCachedHeight =348
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =9
                    HoverTint =60.0
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =11460
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =13005
                    Top =3915
                    Width =3225
                    Height =285
                    BackColor =14211288
                    ForeColor =7500402
                    Name ="TermExceedsMonthsLabel"
                    Caption =" Lease Term more than             Mths"
                    LayoutCachedLeft =13005
                    LayoutCachedTop =3915
                    LayoutCachedWidth =16230
                    LayoutCachedHeight =4200
                    ColumnStart =4
                    ColumnEnd =5
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10965
                    Top =1980
                    Width =2985
                    Height =522
                    FontSize =12
                    TabIndex =15
                    Name ="SaveTax"
                    Caption ="  Save This Tax"
                    OnClick ="[Event Procedure]"
                    FontName ="Century Gothic"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000ab598980ab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989800000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffd9b4caffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffd9b4caffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ff0000000000000000ab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =10965
                    LayoutCachedTop =1980
                    LayoutCachedWidth =13950
                    LayoutCachedHeight =2502
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Width =9288
                    Height =441
                    FontSize =14
                    TabIndex =25
                    BackColor =15527148
                    BorderColor =0
                    ForeColor =8210719
                    Name ="EventDescription"
                    FontName ="Segoe UI"
                    TopPadding =0
                    TextFormat =1

                    LayoutCachedLeft =180
                    LayoutCachedWidth =9468
                    LayoutCachedHeight =441
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9585
                    Top =915
                    Width =1560
                    Height =345
                    TabIndex =1
                    BorderColor =0
                    Name ="StartsOnDate"
                    ControlSource ="StartsOn"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Segoe UI"
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =9585
                    LayoutCachedTop =915
                    LayoutCachedWidth =11145
                    LayoutCachedHeight =1260
                End
                Begin TextBox
                    DecimalPlaces =3
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15780
                    Top =885
                    Width =975
                    Height =345
                    TabIndex =4
                    BorderColor =0
                    Name ="TaxRate"
                    ControlSource ="TaxRate"
                    Format ="Percent"
                    ValidationRule ="Is Not Null"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =15780
                    LayoutCachedTop =885
                    LayoutCachedWidth =16755
                    LayoutCachedHeight =1230
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11265
                    Top =915
                    Width =1560
                    Height =345
                    TabIndex =2
                    BorderColor =0
                    Name ="EndsOnDate"
                    ControlSource ="EndsOn"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Segoe UI"
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =11265
                    LayoutCachedTop =915
                    LayoutCachedWidth =12825
                    LayoutCachedHeight =1260
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =14520
                    Top =990
                    Width =1245
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BackColor =9211020
                    ForeColor =12349952
                    Name ="Label116"
                    Caption ="Exception Rate"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =14520
                    LayoutCachedTop =990
                    LayoutCachedWidth =15765
                    LayoutCachedHeight =1230
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Subform
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =180
                    Top =8595
                    Width =18138
                    Height =2865
                    TabIndex =27
                    Name ="Tax_Any_ExceptionSub"
                    SourceObject ="Form.Tax_Add_ExceptionSub"

                    LayoutCachedLeft =180
                    LayoutCachedTop =8595
                    LayoutCachedWidth =18318
                    LayoutCachedHeight =11460
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =180
                    Top =8280
                    Width =2280
                    Height =315
                    FontSize =10
                    FontWeight =700
                    Name ="Label119"
                    Caption ="▼ EXISTING Exceptions"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =180
                    LayoutCachedTop =8280
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =8595
                    ColumnStart =4
                    ColumnEnd =4
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15210
                    Top =3946
                    Width =435
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    BorderColor =0
                    Name ="TermExceedsMonths"
                    ControlSource ="TermExceedsMonths"
                    Format ="General Number"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Selected"

                    LayoutCachedLeft =15210
                    LayoutCachedTop =3946
                    LayoutCachedWidth =15645
                    LayoutCachedHeight =4186
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12945
                    Top =915
                    Width =1335
                    Height =345
                    TabIndex =3
                    BorderColor =0
                    Name ="TaxMaxValue"
                    ControlSource ="TaxMaxValue"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12945
                    LayoutCachedTop =915
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =1260
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15840
                    Top =585
                    Width =975
                    Height =345
                    TabIndex =16
                    BorderColor =0
                    ForeColor =8210719
                    Name ="CurrTaxRate"
                    Format ="Percent"
                    ValidationRule ="Is Not Null"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =15840
                    LayoutCachedTop =585
                    LayoutCachedWidth =16815
                    LayoutCachedHeight =930
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =14550
                    Top =660
                    Width =1215
                    Height =210
                    FontSize =9
                    FontWeight =700
                    BackColor =9211020
                    ForeColor =4144959
                    Name ="Label273"
                    Caption ="Base Tax Rate"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =14550
                    LayoutCachedTop =660
                    LayoutCachedWidth =15765
                    LayoutCachedHeight =870
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =3
                    SpecialEffect =5
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15795
                    Top =1440
                    Width =1110
                    Height =345
                    FontWeight =700
                    TabIndex =17
                    BorderColor =0
                    ForeColor =8210719
                    Name ="EffectiveTax"
                    Format ="Percent"
                    ValidationRule ="Is Not Null"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =15795
                    LayoutCachedTop =1440
                    LayoutCachedWidth =16905
                    LayoutCachedHeight =1785
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =14325
                    Top =1500
                    Width =1440
                    Height =240
                    FontSize =9
                    FontWeight =700
                    BackColor =9211020
                    ForeColor =4144959
                    Name ="Label279"
                    Caption ="Effective Tax Rate"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =14325
                    LayoutCachedTop =1500
                    LayoutCachedWidth =15765
                    LayoutCachedHeight =1740
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    SpecialEffect =5
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =180
                    Top =2030
                    Width =8205
                    Height =285
                    FontSize =10
                    TabIndex =12
                    BorderColor =0
                    Name ="FormURL1"
                    ControlSource ="FormURL1"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =180
                    LayoutCachedTop =2030
                    LayoutCachedWidth =8385
                    LayoutCachedHeight =2315
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =1
                    ForeThemeColorIndex =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =180
                    Top =1800
                    Width =8205
                    Height =230
                    FontSize =9
                    Name ="Label281"
                    Caption ="Tax Authority Website - Forms"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =180
                    LayoutCachedTop =1800
                    LayoutCachedWidth =8385
                    LayoutCachedHeight =2030
                    ColumnStart =7
                    ColumnEnd =9
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    SpecialEffect =5
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =180
                    Top =2630
                    Width =8205
                    Height =285
                    FontSize =10
                    TabIndex =13
                    BorderColor =0
                    Name ="FormURL2"
                    ControlSource ="FormURL2"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =180
                    LayoutCachedTop =2630
                    LayoutCachedWidth =8385
                    LayoutCachedHeight =2915
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =1
                    ForeThemeColorIndex =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =180
                    Top =2400
                    Width =8205
                    Height =230
                    FontSize =9
                    Name ="Label283"
                    Caption ="Tax Authority Documentation"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =180
                    LayoutCachedTop =2400
                    LayoutCachedWidth =8385
                    LayoutCachedHeight =2630
                    ColumnStart =7
                    ColumnEnd =9
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =180
                    Top =3840
                    Width =1125
                    Height =210
                    FontSize =9
                    FontWeight =700
                    BackColor =9732411
                    Name ="Label434"
                    Caption ="Select NAICS "
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =180
                    LayoutCachedTop =3840
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =4050
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =7200
                    Left =180
                    Top =4065
                    Width =1125
                    Height =270
                    FontSize =9
                    TabIndex =22
                    BorderColor =0
                    Name ="NAICSCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT NAICScodes.NAICS, NAICScodes.NAICSDesc, NAICScodes.INDX FROM NAICScodes O"
                        "RDER BY NAICScodes.INDX;"
                    ColumnWidths ="1152;6048"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =180
                    LayoutCachedTop =4065
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =4335
                End
                Begin TextBox
                    SpecialEffect =5
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1395
                    Top =3870
                    Width =3105
                    Height =465
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    ForeColor =5855577
                    Name ="NAICSDesc"
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =1395
                    LayoutCachedTop =3870
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =4335
                    ColumnStart =7
                    ColumnEnd =9
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =180
                    Top =4789
                    Width =4290
                    Height =1320
                    FontSize =9
                    TabIndex =24
                    BackColor =15527148
                    BorderColor =10855845
                    Name ="NAICSList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TaxExceptionsNAICSTemp.indx, TaxExceptionsNAICSTemp.NAICS, TaxExceptionsN"
                        "AICSTemp.NAICSDesc, TaxExceptionsNAICSTemp.TaxLevel FROM TaxExceptionsNAICSTemp;"
                    ColumnWidths ="0;720;3600"

                    LayoutCachedLeft =180
                    LayoutCachedTop =4789
                    LayoutCachedWidth =4470
                    LayoutCachedHeight =6109
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Top =4455
                    Width =2250
                    Height =274
                    FontSize =10
                    TabIndex =23
                    ForeColor =0
                    Name ="AddNAICS"
                    Caption ="  Add This NAICS"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000072727296727272f0727272a2 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272cc727272ff727272e1 ,
                        0x72727203000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff00000000000000000000000000000000000000007272722a00000000 ,
                        0x00000000000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff0000000000000000000000000000000072727218727272337272721b ,
                        0x00000000000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff000000000000000000000000000000007272727e727272ff7272728d ,
                        0x00000000000000000000000090a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff0000000072727287727272ff72727296 ,
                        0x00000000000000000000000090a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff0000000072727293727272ff727272a5 ,
                        0x00000000000000000000000090a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff0000000072727299727272ff727272b1 ,
                        0x00000000000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff00000000000000000000000000000000727272a5727272ff727272ba ,
                        0x00000000000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff00000000000000000000000000000000727272b1727272ff727272c3 ,
                        0x00000000000000000000000000000000000000000000000090a468ff90a468ff ,
                        0x90a468ff00000000000000000000000000000000727272b7727272ff727272d2 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272c3727272ff727272db ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =180
                    LayoutCachedTop =4455
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =4729
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8965045
                    BorderWidth =1
                    BorderColor =16777215
                    ThemeFontIndex =1
                    HoverColor =4754549
                    PressedColor =12566463
                    PressedForeColor =0
                    QuickStyle =30
                    QuickStyleMask =-1013
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Top =6180
                    Width =2250
                    Height =274
                    FontSize =10
                    TabIndex =26
                    ForeColor =0
                    Name ="DeleteNAICS"
                    Caption =" Delete Selected Item"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000696969126969698a69696951000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000069696906 ,
                        0x696969a5696969276969696f696969ff696969fc6969694e0000000000000000 ,
                        0x00000000000000000000000000000000000000000000000069696906696969b7 ,
                        0x696969b40000000069696921696969f9696969ff696969f66969693000000000 ,
                        0x000000000000000000000000000000000000000069696906696969b4696969f9 ,
                        0x6969692400000000000000006969697e696969ff696969ff696969e46969691b ,
                        0x0000000000000000000000000000000069696906696969ae696969ff69696984 ,
                        0x00000000000000000000000000000000696969a5696969ff696969ff696969d2 ,
                        0x6969690f000000000000000069696906696969ae696969ff696969d269696906 ,
                        0x0000000000000000000000000000000069696906696969b4696969ff696969ff ,
                        0x696969c66969691269696906696969ab696969ff696969f96969693000000000 ,
                        0x00000000000000000000000000000000000000006969690369696993696969ff ,
                        0x696969ff696969d8696969c0696969ff696969ff696969630000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000006969695d ,
                        0x696969f3696969ff696969ff696969ff6969699c000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000006969694e ,
                        0x696969f3696969ff696969ff696969ff69696951000000000000000000000000 ,
                        0x00000000000000000000000000000000000000006969690c69696999696969ff ,
                        0x696969ff696969ff696969cc696969e1696969ff6969698a6969690600000000 ,
                        0x0000000000000000000000000000000069696948696969e1696969ff696969ff ,
                        0x696969ff696969ab696969096969690969696984696969f6696969cc69696921 ,
                        0x0000000000000000000000006969698d696969ff696969ff696969ff696969ff ,
                        0x69696987000000000000000000000000000000006969691e69696993696969ea ,
                        0x696969786969690669696903696969f6696969ff696969ff696969d569696942 ,
                        0x0000000000000000000000000000000000000000000000000000000069696912 ,
                        0x6969696369696927000000006969694e696969b76969696f6969690600000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =180
                    LayoutCachedTop =6180
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =6454
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =12566463
                    BorderColor =2366701
                    ThemeFontIndex =1
                    HoverColor =1643706
                    PressedColor =12566463
                    PressedForeColor =0
                    QuickStyle =30
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =180
                    Top =915
                    Width =9270
                    Height =360
                    Name ="ExceptionDescription"
                    ControlSource ="ExceptionDescription"
                    AfterUpdate ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =180
                    LayoutCachedTop =915
                    LayoutCachedWidth =9450
                    LayoutCachedHeight =1275
                    DisplayAsHyperlink =1
                    RowStart =1
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =5160
                    Top =3465
                    Width =2745
                    Height =330
                    FontSize =12
                    FontWeight =700
                    ForeColor =12349952
                    Name ="Text464"
                    Caption ="Unit Types Rules"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =5160
                    LayoutCachedTop =3465
                    LayoutCachedWidth =7905
                    LayoutCachedHeight =3795
                End
                Begin Label
                    OverlapFlags =85
                    Left =13005
                    Top =3465
                    Width =1875
                    Height =330
                    FontSize =12
                    FontWeight =700
                    ForeColor =12349952
                    Name ="Label466"
                    Caption ="Misc Rules"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =13005
                    LayoutCachedTop =3465
                    LayoutCachedWidth =14880
                    LayoutCachedHeight =3795
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =9015
                    Top =4395
                    Width =3240
                    Height =285
                    BackColor =14211288
                    ForeColor =7500402
                    Name ="WeightLimitMinLbsLabel"
                    Caption ="More than or Equal:                        lbs"
                    LayoutCachedLeft =9015
                    LayoutCachedTop =4395
                    LayoutCachedWidth =12255
                    LayoutCachedHeight =4680
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    AutoTab = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10872
                    Top =4421
                    Width =1020
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    BorderColor =0
                    Name ="WeightLimitMinLbs"
                    ControlSource ="WeightLimitMinLbs"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Selected"

                    LayoutCachedLeft =10872
                    LayoutCachedTop =4421
                    LayoutCachedWidth =11892
                    LayoutCachedHeight =4661
                    DisplayAsHyperlink =1
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =5175
                    Top =3900
                    Width =315
                    Height =300
                    BackColor =12349952
                    Name ="Box470"
                    LayoutCachedLeft =5175
                    LayoutCachedTop =3900
                    LayoutCachedWidth =5490
                    LayoutCachedHeight =4200
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =5256
                    Top =3960
                    Width =215
                    Height =210
                    TabIndex =19
                    BorderColor =16777215
                    Name ="AppliesToTractors2"
                    ControlSource ="AppliesToTractors"
                    OnClick ="[Event Procedure]"
                    GridlineColor =0

                    LayoutCachedLeft =5256
                    LayoutCachedTop =3960
                    LayoutCachedWidth =5471
                    LayoutCachedHeight =4170
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =5490
                    Top =3900
                    Width =2880
                    Height =300
                    FontSize =11
                    FontWeight =400
                    TabIndex =5
                    ForeColor =7500402
                    Name ="AppliesToTractors"
                    ControlSource ="AppliesToTractors"
                    Caption ="Units are Tractors"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="If the Fleet is a 'For Hire' fleet"

                    LayoutCachedLeft =5490
                    LayoutCachedTop =3900
                    LayoutCachedWidth =8370
                    LayoutCachedHeight =4200
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =4754549
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =6108695
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    QuickStyle =9
                    QuickStyleMask =-1009
                    Overlaps =1
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =5175
                    Top =4380
                    Width =315
                    Height =300
                    BackColor =12349952
                    Name ="Box473"
                    LayoutCachedLeft =5175
                    LayoutCachedTop =4380
                    LayoutCachedWidth =5490
                    LayoutCachedHeight =4680
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =5256
                    Top =4440
                    Width =215
                    Height =210
                    TabIndex =18
                    BorderColor =16777215
                    Name ="AppliesToStraightTrucks2"
                    ControlSource ="AppliesToStraightTrucks"
                    OnClick ="[Event Procedure]"
                    GridlineColor =0

                    LayoutCachedLeft =5256
                    LayoutCachedTop =4440
                    LayoutCachedWidth =5471
                    LayoutCachedHeight =4650
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =5490
                    Top =4380
                    Width =2880
                    Height =300
                    FontSize =11
                    FontWeight =400
                    TabIndex =6
                    ForeColor =7500402
                    Name ="AppliesToStraightTrucks"
                    ControlSource ="AppliesToStraightTrucks"
                    Caption ="Units are Straight Trucks"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="If the Fleet is a 'For Hire' fleet"

                    LayoutCachedLeft =5490
                    LayoutCachedTop =4380
                    LayoutCachedWidth =8370
                    LayoutCachedHeight =4680
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =4754549
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =6108695
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    QuickStyle =9
                    QuickStyleMask =-753
                    Overlaps =1
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =5175
                    Top =4860
                    Width =315
                    Height =300
                    BackColor =12349952
                    Name ="Box476"
                    LayoutCachedLeft =5175
                    LayoutCachedTop =4860
                    LayoutCachedWidth =5490
                    LayoutCachedHeight =5160
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =5256
                    Top =4920
                    Width =215
                    Height =210
                    TabIndex =20
                    BorderColor =16777215
                    Name ="AppliesToTrailers2"
                    ControlSource ="AppliesToTrailers"
                    OnClick ="[Event Procedure]"
                    GridlineColor =0

                    LayoutCachedLeft =5256
                    LayoutCachedTop =4920
                    LayoutCachedWidth =5471
                    LayoutCachedHeight =5130
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =5490
                    Top =4860
                    Width =2880
                    Height =300
                    FontSize =11
                    FontWeight =400
                    TabIndex =7
                    ForeColor =7500402
                    Name ="AppliesToTrailers"
                    ControlSource ="AppliesToTrailers"
                    Caption ="Units are Trailers"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="If the Fleet is a 'For Hire' fleet"

                    LayoutCachedLeft =5490
                    LayoutCachedTop =4860
                    LayoutCachedWidth =8370
                    LayoutCachedHeight =5160
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =4754549
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =6108695
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    QuickStyle =9
                    QuickStyleMask =-753
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =9015
                    Top =3465
                    Width =3135
                    Height =330
                    FontSize =12
                    FontWeight =700
                    ForeColor =12349952
                    Name ="Label479"
                    Caption ="Unit Weights Rules"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =9015
                    LayoutCachedTop =3465
                    LayoutCachedWidth =12150
                    LayoutCachedHeight =3795
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =9015
                    Top =3915
                    Width =3240
                    Height =285
                    BackColor =14211288
                    ForeColor =7500402
                    Name ="WeightLimitMaxLbsLabel"
                    Caption ="Less than or Equal:                          lbs"
                    LayoutCachedLeft =9015
                    LayoutCachedTop =3915
                    LayoutCachedWidth =12255
                    LayoutCachedHeight =4200
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    AutoTab = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10872
                    Top =3946
                    Width =1020
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    BorderColor =0
                    Name ="WeightLimitMaxLbs"
                    ControlSource ="WeightLimitMaxLbs"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                    Tag ="Selected"
                    ControlTipText ="If the Unit weights this value or less"

                    LayoutCachedLeft =10872
                    LayoutCachedTop =3946
                    LayoutCachedWidth =11892
                    LayoutCachedHeight =4186
                    DisplayAsHyperlink =1
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =3465
                    Width =3930
                    Height =330
                    FontSize =12
                    FontWeight =700
                    ForeColor =12349952
                    Name ="Label482"
                    Caption ="Industry Rules"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =180
                    LayoutCachedTop =3465
                    LayoutCachedWidth =4110
                    LayoutCachedHeight =3795
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =1440
                    Width =4260
                    Height =330
                    FontSize =12
                    FontWeight =700
                    ForeColor =12349952
                    Name ="Label483"
                    Caption ="Exception Documentation and Form"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1440
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =1770
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =16980
                    Top =840
                    Width =300
                    Height =510
                    BackColor =0
                    Name ="Box484"
                    LayoutCachedLeft =16980
                    LayoutCachedTop =840
                    LayoutCachedWidth =17280
                    LayoutCachedHeight =1350
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =17040
                    Top =1035
                    Width =230
                    Height =150
                    TabIndex =21
                    BorderColor =16777215
                    Name ="TaxRule2"
                    ControlSource ="TaxRule"
                    OnClick ="[Event Procedure]"
                    GridlineColor =0

                    LayoutCachedLeft =17040
                    LayoutCachedTop =1035
                    LayoutCachedWidth =17270
                    LayoutCachedHeight =1185
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =17280
                    Top =840
                    Width =945
                    Height =510
                    FontSize =9
                    FontWeight =400
                    TabIndex =14
                    ForeColor =7500402
                    Name ="TaxRule"
                    ControlSource ="TaxRule"
                    Caption ="Replaces \015\012Base Tax"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="If the Fleet is a 'For Hire' fleet"

                    LayoutCachedLeft =17280
                    LayoutCachedTop =840
                    LayoutCachedWidth =18225
                    LayoutCachedHeight =1350
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =4754549
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =15709952
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =-1
                    QuickStyle =9
                    QuickStyleMask =-1009
                    Overlaps =1
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =85
                    Left =15840
                    Top =1320
                    Width =975
                    Name ="Line487"
                    LayoutCachedLeft =15840
                    LayoutCachedTop =1320
                    LayoutCachedWidth =16815
                    LayoutCachedHeight =1320
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =6540
                    Width =3405
                    Height =225
                    FontSize =9
                    TabIndex =28
                    BorderColor =0
                    Name ="Label488"
                    ControlSource ="=\"*NAICS<b> 4842</b> = <i>'For Hire'</i> fleet\""
                    TextFormat =1

                    LayoutCachedLeft =180
                    LayoutCachedTop =6540
                    LayoutCachedWidth =3585
                    LayoutCachedHeight =6765
                    ThemeFontIndex =1
                End
                Begin Label
                    OverlapFlags =93
                    Left =5175
                    Top =5715
                    Width =2580
                    Height =330
                    FontSize =12
                    FontWeight =700
                    ForeColor =8355711
                    Name ="Label489"
                    Caption ="Exception Rules Detail"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =5175
                    LayoutCachedTop =5715
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =6045
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5130
                    Top =6000
                    Width =10320
                    Height =2445
                    FontSize =10
                    TabIndex =29
                    BackColor =15527148
                    Name ="TaxScript"
                    ControlSource ="TaxScript"
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedLeft =5130
                    LayoutCachedTop =6000
                    LayoutCachedWidth =15450
                    LayoutCachedHeight =8445
                End
                Begin TextBox
                    Visible = NotDefault
                    AutoTab = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2580
                    Top =4485
                    Width =1755
                    FontSize =10
                    FontWeight =700
                    TabIndex =30
                    BorderColor =0
                    Name ="IncludedNAICS"
                    ControlSource ="IncludedNAICS"
                    Format ="Standard"
                    ControlTipText ="If the Unit weights this value or less"

                    LayoutCachedLeft =2580
                    LayoutCachedTop =4485
                    LayoutCachedWidth =4335
                    LayoutCachedHeight =4725
                    DisplayAsHyperlink =1
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    Left =180
                    Top =600
                    Width =14220
                    Height =330
                    FontSize =12
                    FontWeight =700
                    ForeColor =12349952
                    Name ="Label493"
                    Caption ="Exception Description                                                           "
                        "                                                            Start Date          "
                        " End Date       Max Amt $   "
                    FontName ="Century Gothic"
                    LayoutCachedLeft =180
                    LayoutCachedTop =600
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =930
                End
                Begin Label
                    OverlapFlags =85
                    Left =15705
                    Top =195
                    Width =1275
                    Height =330
                    FontSize =12
                    FontWeight =700
                    ForeColor =12349952
                    Name ="Label494"
                    Caption ="Tax  Rates"
                    FontName ="Century Gothic"
                    LayoutCachedLeft =15705
                    LayoutCachedTop =195
                    LayoutCachedWidth =16980
                    LayoutCachedHeight =525
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =16920
                    Top =4650
                    Width =1305
                    Height =522
                    FontSize =10
                    TabIndex =31
                    Name ="DeleteMe"
                    Caption ="Delete \015\012Exception"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete This Base Tax"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000008080803980808096868686d6 ,
                        0x828282f7808080ff828282f8868686d68585859c808080390000000000000000 ,
                        0x0000000000000000000000000000000000000000808080ff808080ff808080ff ,
                        0x808080ff808080ff808080ff808080ff808080ff808080ff0000000000000000 ,
                        0x0000000000000000000000000000000080808006818181fbcbcbcbffe6e6e6ff ,
                        0xf9f9f9fffffffffff9f9f9ffe6e6e6ffcbcbcbff808080ff0000000000000000 ,
                        0x000000000000000000000000000000008080801e898989edffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff868686fc8080801500000000 ,
                        0x0000000000000000000000000000000080808036939393e6ffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff929292fc8080802d00000000 ,
                        0x00000000000000000000000000000000808080519f9f9fe3ffffffffe0e5d4ff ,
                        0x90a468ff90a468ff90a468ffdce2cfffffffffffa0a0a0fe8383834400000000 ,
                        0x0000000000000000000000000000000080808069afafafe5ffffffffadbc8fff ,
                        0xcfd7bdffffffffffd8dfcaffa6b686fffefefdffadadadff9393936a00000000 ,
                        0x0000000000000000000000000000000080808081bebebeecfffffffff6f8f3ff ,
                        0xa0b17dffc2cdacff9eb07cfff3f5efffffffffffb9b9b9ff9999998e00000000 ,
                        0x000000000000000000000000000000008080809ccececef9ffffffffffffffff ,
                        0xeff2eaffbbc7a3ffebefe4ffffffffffffffffffc5c5c5ff979797ad00000000 ,
                        0x00000000000000000000000000000000838383b8ddddddffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffcfcfcfff969696c700000000 ,
                        0x00000000000000000000000000000000848484d3e9e9e9ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffddddddff929292dc00000000 ,
                        0x00000000000000000000000000000000838383edccccccff9b9b9bff808080ff ,
                        0x808080ff808080ff808080ff808080ff9b9b9bffc2c2c2ff8d8d8ded00000000 ,
                        0x00000000000000000000000080808003808080ffa4a4a4ffdededeffffffffff ,
                        0xffffffffffffffffffffffffffffffffdededeffa4a4a4ff808080ff00000000 ,
                        0x00000000000000000000000000000000808080ff9c9c9ccddbdbdbe7ffffffff ,
                        0xffffffffffffffffffffffffffffffffddddddf4a4a4a4dd808080ff00000000 ,
                        0x00000000000000000000000000000000000000008080804e808080bd808080ff ,
                        0x808080ff808080ff808080ff808080ff808080bd8080804e0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =16920
                    LayoutCachedTop =4650
                    LayoutCachedWidth =18225
                    LayoutCachedHeight =5172
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =255
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =1
                    ThemeFontIndex =1
                    HoverColor =10855845
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "Tax_Add_Exception.cls"
