Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =27936
    DatasheetFontHeight =11
    ItemSuffix =92
    Left =6540
    Top =3915
    Right =-30811
    Bottom =18510
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xf80cbbca5255e440
    End
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
        Begin Tab
            TextFontFamily =18
            FontSize =11
            FontName ="Cambria"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =765
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4440
                    Top =300
                    Width =8460
                    Height =360
                    ColumnOrder =4
                    FontSize =10
                    FontWeight =700
                    Name ="BuyNo"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =4440
                    LayoutCachedTop =300
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4440
                            Top =120
                            Width =1800
                            Height =180
                            FontSize =8
                            Name ="Label13"
                            Caption ="You're Working with:"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =4440
                            LayoutCachedTop =120
                            LayoutCachedWidth =6240
                            LayoutCachedHeight =300
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =24090
                    Top =60
                    Width =480
                    Height =210
                    ColumnOrder =2
                    FontSize =8
                    TabIndex =1
                    ForeColor =-2147483615
                    Name ="Text47"
                    ControlSource ="=[Units_List].[ListCount]-1"

                    LayoutCachedLeft =24090
                    LayoutCachedTop =60
                    LayoutCachedWidth =24570
                    LayoutCachedHeight =270
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =23595
                    Top =75
                    Width =450
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =13056
                    Name ="Label49"
                    Caption ="Units"
                    LayoutCachedLeft =23595
                    LayoutCachedTop =75
                    LayoutCachedWidth =24045
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =24075
                    Top =285
                    Width =480
                    Height =210
                    ColumnOrder =3
                    FontSize =8
                    TabIndex =2
                    ForeColor =-2147483615
                    Name ="Text63"
                    ControlSource ="=[Units_List].[ItemsSelected].[Count]"

                    LayoutCachedLeft =24075
                    LayoutCachedTop =285
                    LayoutCachedWidth =24555
                    LayoutCachedHeight =495
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =23580
                    Top =300
                    Width =450
                    Height =195
                    FontSize =8
                    FontWeight =700
                    ForeColor =13056
                    Name ="Label64"
                    Caption ="Units"
                    LayoutCachedLeft =23580
                    LayoutCachedTop =300
                    LayoutCachedWidth =24030
                    LayoutCachedHeight =495
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =540
                    Width =27936
                    Height =195
                    BackColor =-2147483613
                    BorderColor =14211288
                    Name ="Image71"
                    Picture ="DrkRedOnLtGrey_Shadow"
                    GridlineColor =0
                    HorizontalAnchor =2

                    LayoutCachedTop =540
                    LayoutCachedWidth =27936
                    LayoutCachedHeight =735
                    TabIndex =10
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18445
                    Top =240
                    Width =689
                    Height =300
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    BorderColor =1643706
                    Name ="UnitCount"
                    ControlSource ="=[ReMarketingMassMod_sub_SUFParked].[Form]![UnitCount]"

                    LayoutCachedLeft =18445
                    LayoutCachedTop =240
                    LayoutCachedWidth =19134
                    LayoutCachedHeight =540
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =255
                            TextAlign =2
                            Left =18420
                            Top =60
                            Width =750
                            Height =540
                            FontSize =8
                            BackColor =1643706
                            BorderColor =1643706
                            Name ="Label77"
                            Caption ="Units"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =18420
                            LayoutCachedTop =60
                            LayoutCachedWidth =19170
                            LayoutCachedHeight =600
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =255
                    Top =60
                    Width =4365
                    Height =570
                    FontSize =20
                    Name ="Label31"
                    Caption ="Mass Modification Form"
                    FontName ="Segoe UI"
                    LayoutCachedTop =60
                    LayoutCachedWidth =4365
                    LayoutCachedHeight =630
                End
                Begin Label
                    OverlapFlags =247
                    Left =30
                    Width =1260
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =8872517
                    Name ="Label294"
                    Caption ="Buyers Order"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =30
                    LayoutCachedWidth =1290
                    LayoutCachedHeight =240
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =26175
                    Top =180
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =4
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =26175
                    LayoutCachedTop =180
                    LayoutCachedWidth =26925
                    LayoutCachedHeight =408
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
                Begin CommandButton
                    OverlapFlags =215
                    Left =25620
                    Top =180
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =5
                    Name ="GoHome"
                    Caption ="Home"
                    ControlTipText ="\"Open Main Menu\""
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
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

                    LayoutCachedLeft =25620
                    LayoutCachedTop =180
                    LayoutCachedWidth =26175
                    LayoutCachedHeight =408
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
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =24780
                    Top =180
                    Width =480
                    Height =210
                    FontSize =8
                    TabIndex =6
                    ForeColor =-2147483615
                    Name ="ClientGroupID"

                    LayoutCachedLeft =24780
                    LayoutCachedTop =180
                    LayoutCachedWidth =25260
                    LayoutCachedHeight =390
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =15540
                    Top =120
                    Width =1927
                    Height =240
                    FontSize =8
                    TabIndex =7
                    Name ="MaketitleRequest"
                    Caption ="Make ALL Title Request"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272b7 ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffd7d7d7ff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffdbdbdbff727272ff ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272bd ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =15540
                    LayoutCachedTop =120
                    LayoutCachedWidth =17467
                    LayoutCachedHeight =360
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
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
                    OverlapFlags =87
                    Left =12960
                    Top =120
                    Width =2467
                    Height =240
                    FontSize =8
                    TabIndex =8
                    Name ="MakeSingleTitleRequest"
                    Caption ="Make SINGLE Title Request"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272b7 ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffd7d7d7ff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffdbdbdbff727272ff ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272bd ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =12960
                    LayoutCachedTop =120
                    LayoutCachedWidth =15427
                    LayoutCachedHeight =360
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
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
                    OverlapFlags =85
                    Left =21780
                    Top =120
                    Width =967
                    Height =240
                    FontSize =8
                    TabIndex =9
                    Name ="MakeBTF"
                    Caption ="Make BTF"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272b7 ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffd7d7d7ff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffff90a468ff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff727272ff ,
                        0x000000000000000000000000b3b3b3ff00000000b3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x00000000000000000000000000000000727272ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffdbdbdbff727272ff ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272bd ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =21780
                    LayoutCachedTop =120
                    LayoutCachedWidth =22747
                    LayoutCachedHeight =360
                    PictureCaptionArrangement =5
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
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
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =13575
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =240
                    Top =315
                    Width =5040
                    Height =270
                    FontSize =8
                    Name ="Label89"
                    Caption ="┌────┬──Search ───────────┬──────────┐"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =240
                    LayoutCachedTop =315
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =585
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Top =525
                    Width =720
                    Height =270
                    FontSize =9
                    TabIndex =10
                    BorderColor =967423
                    Name ="SearchBO"
                    OnChange ="[Event Procedure]"

                    LayoutCachedTop =525
                    LayoutCachedWidth =720
                    LayoutCachedHeight =795
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =150
                    Top =9600
                    Width =5472
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label58"
                    Caption ="┌───┬  Search ────────┬──────┬────┬────────┐"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =150
                    LayoutCachedTop =9600
                    LayoutCachedWidth =5622
                    LayoutCachedHeight =9885
                    BorderThemeColorIndex =4
                    BorderShade =75.0
                    ForeThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Left =3975
                    Top =9630
                    Width =691
                    Height =225
                    FontSize =8
                    FontWeight =700
                    Name ="ClearLabel"
                    Caption ="◄Clear►"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =3975
                    LayoutCachedTop =9630
                    LayoutCachedWidth =4666
                    LayoutCachedHeight =9855
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    ForeTint =60.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =9
                    Left =60
                    Top =10110
                    Width =5730
                    Height =3390
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    BoundColumn =1
                    BackColor =15527148
                    Name ="Units_List"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT InvoiceRequestID as BO,UnitID,UnitNum,VIN,[Status],Make,SubType,[Type],MY"
                        "ear From vw_BuyersOrder Where InvoiceRequestID = 2391 Order BY InvoiceRequestID "
                        "DESC; "
                    ColumnWidths ="0;720;720;2016;288;1152;720;720;0"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =10110
                    LayoutCachedWidth =5790
                    LayoutCachedHeight =13500
                    BorderThemeColorIndex =2
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =780
                    Top =9840
                    Width =720
                    Height =255
                    FontSize =9
                    TabIndex =3
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchUnit"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =9840
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =10095
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =223
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1500
                    Top =9840
                    Width =2010
                    Height =255
                    FontSize =9
                    TabIndex =4
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchVin"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1500
                    LayoutCachedTop =9840
                    LayoutCachedWidth =3510
                    LayoutCachedHeight =10095
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Top =9180
                    Width =5520
                    Height =315
                    FontSize =10
                    TabIndex =1
                    Name ="GetAll"
                    Caption ="▼All 8 Listed Units ►"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =180
                    LayoutCachedTop =9180
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =9495
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =6
                    BackShade =50.0
                    BorderColor =6108695
                    HoverThemeColorIndex =6
                    HoverShade =75.0
                    PressedThemeColorIndex =6
                    PressedShade =50.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =9840
                    Width =720
                    Height =255
                    FontSize =9
                    TabIndex =2
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchAsset"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =9840
                    LayoutCachedWidth =780
                    LayoutCachedHeight =10095
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3510
                    Top =9840
                    Width =285
                    Height =255
                    FontSize =9
                    TabIndex =5
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchStatus"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3510
                    LayoutCachedTop =9840
                    LayoutCachedWidth =3795
                    LayoutCachedHeight =10095
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4950
                    Top =9840
                    Width =825
                    Height =255
                    FontSize =9
                    TabIndex =7
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchType"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4950
                    LayoutCachedTop =9840
                    LayoutCachedWidth =5775
                    LayoutCachedHeight =10095
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =5
                    Top =795
                    Width =5760
                    Height =8325
                    FontSize =9
                    Name ="BuyersOrderList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT InvoiceRequestID as BO ,BuyerShortName AS Buyer ,InvoicePaidDate AS Paid "
                        ",SP ,Count(UnitID) as Units FROM vw_BuyersOrder WHERE WON = True  GROUP BY Invoi"
                        "ceRequestID, BuyerShortName,InvoicePaidDate,SP ORDER BY InvoiceRequestID DESC;"
                    ColumnWidths ="720;2592;1152;576;1152"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedTop =795
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =9120
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    BorderShade =85.0
                End
                Begin Tab
                    OverlapFlags =85
                    Left =5865
                    Width =22050
                    Height =13575
                    FontSize =10
                    TabIndex =9
                    Name ="RMKTTabControl"
                    FontName ="Segoe UI"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =5865
                    LayoutCachedWidth =27915
                    LayoutCachedHeight =13575
                    ThemeFontIndex =-1
                    Shape =0
                    BackColor =16777215
                    BackShade =100.0
                    BorderColor =9211020
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeColor =8210719
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =5940
                            Top =418
                            Width =21900
                            Height =13082
                            Name ="SUF - Parked - Inspection Inv"
                            EventProcPrefix ="SUF___Parked___Inspection_Inv"
                            LayoutCachedLeft =5940
                            LayoutCachedTop =418
                            LayoutCachedWidth =27840
                            LayoutCachedHeight =13500
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    OldBorderStyle =0
                                    Left =5976
                                    Top =418
                                    Width =21510
                                    Height =13080
                                    Name ="ReMarketingMassMod_sub_SUFParked"
                                    SourceObject ="Form.ReMarketingMassMod_sub_SUFParked"
                                    HorizontalAnchor =2
                                    VerticalAnchor =2
                                    FilterOnEmptyMaster =0

                                    LayoutCachedLeft =5976
                                    LayoutCachedTop =418
                                    LayoutCachedWidth =27486
                                    LayoutCachedHeight =13498
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5940
                            Top =420
                            Width =21900
                            Height =13080
                            Name ="Inspection2"
                            Caption ="Inspections && Links"
                            LayoutCachedLeft =5940
                            LayoutCachedTop =420
                            LayoutCachedWidth =27840
                            LayoutCachedHeight =13500
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5970
                                    Top =420
                                    Width =21636
                                    Height =12816
                                    Name ="ReMarketingMassMod_sub_Inspection2"
                                    SourceObject ="Form.ReMarketingMassMod_sub_Inspection2"

                                    LayoutCachedLeft =5970
                                    LayoutCachedTop =420
                                    LayoutCachedWidth =27606
                                    LayoutCachedHeight =13236
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5940
                            Top =418
                            Width =21900
                            Height =13082
                            Name ="SetUnitValues"
                            Caption ="Set Unit Values / Bank Share"
                            LayoutCachedLeft =5940
                            LayoutCachedTop =418
                            LayoutCachedWidth =27840
                            LayoutCachedHeight =13500
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5976
                                    Top =418
                                    Width =19800
                                    Height =12816
                                    Name ="ReMarketingMassMod_sub_FMVPenalty"
                                    SourceObject ="Form.ReMarketingMassMod_sub_FMVPenalty"

                                    LayoutCachedLeft =5976
                                    LayoutCachedTop =418
                                    LayoutCachedWidth =25776
                                    LayoutCachedHeight =13234
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5940
                            Top =418
                            Width =21900
                            Height =13082
                            Name ="Sold - Invoicing"
                            EventProcPrefix ="Sold___Invoicing"
                            LayoutCachedLeft =5940
                            LayoutCachedTop =418
                            LayoutCachedWidth =27840
                            LayoutCachedHeight =13500
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5976
                                    Top =418
                                    Width =19776
                                    Height =12816
                                    Name ="ReMarketingMassMod_sub_SoldInvoice"
                                    SourceObject ="Form.ReMarketingMassMod_sub_SoldInvoice"

                                    LayoutCachedLeft =5976
                                    LayoutCachedTop =418
                                    LayoutCachedWidth =25752
                                    LayoutCachedHeight =13234
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5940
                            Top =465
                            Width =21900
                            Height =13035
                            Name ="Buy Out - Interim Funding"
                            EventProcPrefix ="Buy_Out___Interim_Funding"
                            LayoutCachedLeft =5940
                            LayoutCachedTop =465
                            LayoutCachedWidth =27840
                            LayoutCachedHeight =13500
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =6000
                                    Top =480
                                    Width =21417
                                    Height =12816
                                    Name ="ReMarketingMassMod_sub_BuyOutView"
                                    SourceObject ="Form.ReMarketingMassMod_sub_BuyOutView"

                                    LayoutCachedLeft =6000
                                    LayoutCachedTop =480
                                    LayoutCachedWidth =27417
                                    LayoutCachedHeight =13296
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =5940
                            Top =465
                            Width =21900
                            Height =13035
                            Name ="Titles"
                            Caption ="Titles Check In"
                            LayoutCachedLeft =5940
                            LayoutCachedTop =465
                            LayoutCachedWidth =27840
                            LayoutCachedHeight =13500
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5940
                                    Top =480
                                    Width =21837
                                    Height =12816
                                    Name ="RemarketingMassMod_sub_Titles"
                                    SourceObject ="Form.RemarketingMassMod_sub_Titles"

                                    LayoutCachedLeft =5940
                                    LayoutCachedTop =480
                                    LayoutCachedWidth =27777
                                    LayoutCachedHeight =13296
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =5100
                    Top =540
                    Width =600
                    Height =225
                    FontSize =8
                    FontWeight =700
                    Name ="ClearClientSearch"
                    Caption ="◄Clear"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =5100
                    LayoutCachedTop =540
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =765
                    ForeThemeColorIndex =2
                    ForeTint =60.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3795
                    Top =9840
                    Width =1155
                    Height =255
                    FontSize =9
                    TabIndex =6
                    BackColor =13434879
                    BorderColor =3381759
                    Name ="SearchMake"
                    OnChange ="[Event Procedure]"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3795
                    LayoutCachedTop =9840
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =10095
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =3
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =720
                    Top =525
                    Width =2592
                    Height =270
                    FontSize =9
                    TabIndex =11
                    BorderColor =967423
                    Name ="SearchBuyer"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =720
                    LayoutCachedTop =525
                    LayoutCachedWidth =3312
                    LayoutCachedHeight =795
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3312
                    Top =525
                    Width =1152
                    Height =270
                    FontSize =9
                    TabIndex =12
                    BorderColor =967423
                    Name ="SearchPaid"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =3312
                    LayoutCachedTop =525
                    LayoutCachedWidth =4464
                    LayoutCachedHeight =795
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4464
                    Top =525
                    Width =576
                    Height =270
                    FontSize =9
                    TabIndex =13
                    BorderColor =967423
                    Name ="SearchSP"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =4464
                    LayoutCachedTop =525
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =795
                    BackThemeColorIndex =6
                    BackTint =20.0
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Width =4575
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Label88"
                    Caption ="Select Buyers Order | Double-Click to Edit those Units"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =4695
                    LayoutCachedHeight =285
                    ForeThemeColorIndex =0
                End
            End
        End
        Begin FormFooter
            Height =270
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =23550
                    Width =705
                    Height =225
                    FontSize =9
                    BorderColor =9211020
                    ForeColor =9211020
                    Name ="UnitID"

                    LayoutCachedLeft =23550
                    LayoutCachedWidth =24255
                    LayoutCachedHeight =225
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =22920
                            Width =645
                            Height =270
                            FontSize =9
                            BorderColor =9211020
                            ForeColor =9211020
                            Name ="Label19"
                            Caption ="UnitID:"
                            LayoutCachedLeft =22920
                            LayoutCachedWidth =23565
                            LayoutCachedHeight =270
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "BuyersOrderMassMod.cls"
