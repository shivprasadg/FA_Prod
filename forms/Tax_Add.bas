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
    Width =13680
    DatasheetFontHeight =11
    ItemSuffix =130
    Left =6375
    Top =3360
    Right =20310
    Bottom =10650
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xdb68277df7fbe440
    End
    RecordSource ="TaxBaseTemp"
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
                    Width =13680
                    Height =135
                    Name ="Image142"
                    Picture ="2_WhiteShadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =465
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =600
                    TabIndex =8
                End
                Begin Label
                    OverlapFlags =255
                    Left =120
                    Width =2400
                    Height =495
                    FontSize =20
                    FontWeight =700
                    Name ="Label105"
                    Caption ="Base Taxes- "
                    FontName ="Century Gothic"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =495
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8661
                    Width =543
                    Height =231
                    ColumnOrder =6
                    FontSize =8
                    Name ="TaxLevelLookupID"
                    ControlSource ="TaxLevelLookupID"
                    TopPadding =0

                    LayoutCachedLeft =8661
                    LayoutCachedWidth =9204
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8118
                    Width =543
                    Height =231
                    ColumnOrder =4
                    FontSize =8
                    TabIndex =1
                    Name ="TaxTypeID2"
                    ControlSource ="TaxTypeId"
                    TopPadding =0

                    LayoutCachedLeft =8118
                    LayoutCachedWidth =8661
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2520
                    Width =4965
                    Height =495
                    ColumnOrder =5
                    FontSize =20
                    FontWeight =700
                    TabIndex =2
                    BorderColor =0
                    ForeColor =12349952
                    Name ="TaxLevel"
                    ControlSource ="TaxLevel"
                    FontName ="Century Gothic"

                    LayoutCachedLeft =2520
                    LayoutCachedWidth =7485
                    LayoutCachedHeight =495
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7575
                    Width =543
                    Height =231
                    ColumnOrder =3
                    FontSize =8
                    TabIndex =3
                    Name ="Mode"
                    TopPadding =0

                    LayoutCachedLeft =7575
                    LayoutCachedWidth =8118
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9204
                    Width =543
                    Height =231
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =4
                    Name ="EditedTaxID"
                    ControlSource ="EditedTaxID"
                    TopPadding =0

                    LayoutCachedLeft =9204
                    LayoutCachedWidth =9747
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10110
                    Width =543
                    Height =231
                    ColumnOrder =0
                    FontSize =8
                    TabIndex =5
                    Name ="PreviousTaxID"
                    TopPadding =0

                    LayoutCachedLeft =10110
                    LayoutCachedWidth =10653
                    LayoutCachedHeight =231
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =12300
                    Top =60
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =6
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

                    LayoutCachedLeft =12300
                    LayoutCachedTop =60
                    LayoutCachedWidth =12855
                    LayoutCachedHeight =288
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
                    Left =12867
                    Top =60
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =7
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    HorizontalAnchor =1

                    LayoutCachedLeft =12867
                    LayoutCachedTop =60
                    LayoutCachedWidth =13422
                    LayoutCachedHeight =288
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
            Height =6705
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =120
                    Top =900
                    Width =2970
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =0
                    ForeColor =16777215
                    Name ="Label115"
                    Caption ="Tax Type"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =900
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =1125
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =8820
                    Top =915
                    Width =1920
                    Height =522
                    Name ="SaveTax"
                    Caption ="  Save This Tax"
                    OnClick ="[Event Procedure]"
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

                    LayoutCachedLeft =8820
                    LayoutCachedTop =915
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =1437
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
                    ThemeFontIndex =1
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
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4980
                    Top =1125
                    Width =1560
                    Height =345
                    TabIndex =2
                    BorderColor =0
                    Name ="StartsOnDate"
                    ControlSource ="StartsOn"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Segoe UI"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =4980
                    LayoutCachedTop =1125
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1470
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =120
                    Top =15
                    Width =4068
                    Height =270
                    FontSize =10
                    FontWeight =700
                    Name ="TaxNumberLabel"
                    Caption ="▼ NEW Tax Description"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =15
                    LayoutCachedWidth =4188
                    LayoutCachedHeight =285
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
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =300
                    Width =10023
                    Height =441
                    FontSize =14
                    TabIndex =4
                    BackColor =15527148
                    BorderColor =0
                    ForeColor =8210719
                    Name ="EventDescription"
                    FontName ="Segoe UI"
                    TopPadding =0
                    TextFormat =1

                    LayoutCachedLeft =120
                    LayoutCachedTop =300
                    LayoutCachedWidth =10143
                    LayoutCachedHeight =741
                End
                Begin TextBox
                    DecimalPlaces =4
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3405
                    Top =1125
                    Width =1305
                    Height =345
                    TabIndex =1
                    BorderColor =0
                    Name ="BaseTaxRate"
                    ControlSource ="BaseTaxRate"
                    Format ="Percent"
                    ValidationRule ="Is Not Null"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =3405
                    LayoutCachedTop =1125
                    LayoutCachedWidth =4710
                    LayoutCachedHeight =1470
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =4980
                    Top =900
                    Width =1560
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14151142
                    Name ="Label112"
                    Caption ="Starts On"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4980
                    LayoutCachedTop =900
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1125
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =1
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6774
                    Top =1125
                    Width =1560
                    Height =345
                    TabIndex =3
                    BackColor =11860720
                    BorderColor =0
                    Name ="EndsOnDate"
                    ControlSource ="EndsOn"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Segoe UI"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =6774
                    LayoutCachedTop =1125
                    LayoutCachedWidth =8334
                    LayoutCachedHeight =1470
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =6774
                    Top =900
                    Width =1560
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =14151142
                    Name ="Label114"
                    Caption ="Ends On"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =6774
                    LayoutCachedTop =900
                    LayoutCachedWidth =8334
                    LayoutCachedHeight =1125
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =1
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =12960
                    Left =120
                    Top =1125
                    Width =2970
                    Height =345
                    TabIndex =5
                    BorderColor =0
                    Name ="TaxTypeID"
                    ControlSource ="TaxTypeId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TaxType.Id, TaxType.TaxFullName, TaxType.TaxAbbreviation, TaxType.TaxDesc"
                        "ription FROM TaxType;"
                    ColumnWidths ="0;2160;0;3600"
                    ValidationRule ="Is Not Null"
                    FontName ="Segoe UI"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =120
                    LayoutCachedTop =1125
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =1470
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =3405
                    Top =900
                    Width =1305
                    Height =225
                    FontSize =9
                    FontWeight =700
                    BackColor =0
                    ForeColor =16777215
                    Name ="Label116"
                    Caption ="Tax Rate"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =3405
                    LayoutCachedTop =900
                    LayoutCachedWidth =4710
                    LayoutCachedHeight =1125
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Subform
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =120
                    Top =4380
                    Width =13536
                    Height =2325
                    TabIndex =6
                    Name ="Tax_AnySub"
                    SourceObject ="Form.Tax_AnySub"

                    LayoutCachedLeft =120
                    LayoutCachedTop =4380
                    LayoutCachedWidth =13656
                    LayoutCachedHeight =6705
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =120
                    Top =4095
                    Width =4473
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label119"
                    Caption ="▼ EXISTING Taxes"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =4095
                    LayoutCachedWidth =4593
                    LayoutCachedHeight =4380
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
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =120
                    Top =2625
                    Width =9810
                    Height =1470
                    FontSize =10
                    TabIndex =7
                    BackColor =15527148
                    Name ="ValidateMsg"
                    ControlSource ="TaxScript"
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedLeft =120
                    LayoutCachedTop =2625
                    LayoutCachedWidth =9930
                    LayoutCachedHeight =4095
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =12240
                    Top =915
                    Width =1305
                    Height =522
                    FontSize =10
                    TabIndex =8
                    Name ="DeleteMe"
                    Caption ="Delete \015\012This Tax"
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

                    LayoutCachedLeft =12240
                    LayoutCachedTop =915
                    LayoutCachedWidth =13545
                    LayoutCachedHeight =1437
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
                Begin TextBox
                    SpecialEffect =5
                    BorderWidth =3
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =1800
                    Width =9810
                    Height =750
                    FontSize =10
                    TabIndex =9
                    BackColor =15527148
                    Name ="TaxNote"
                    ControlSource ="TaxNote"
                    FontName ="Segoe UI"
                    TextFormat =1

                    LayoutCachedLeft =120
                    LayoutCachedTop =1800
                    LayoutCachedWidth =9930
                    LayoutCachedHeight =2550
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =105
                    Top =1530
                    Width =4473
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label129"
                    Caption ="▼ Tax  Break-Out / Notes"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =105
                    LayoutCachedTop =1530
                    LayoutCachedWidth =4578
                    LayoutCachedHeight =1815
                    ColumnStart =4
                    ColumnEnd =4
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
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
' See "Tax_Add.cls"
