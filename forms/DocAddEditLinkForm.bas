Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =11
    ItemSuffix =23
    Left =5145
    Top =3810
    Right =16200
    Bottom =10305
    Filter ="DocID=170"
    RecSrcDt = Begin
        0x2a3c2d3df3a5e540
    End
    RecordSource ="Select * From DocLinks"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =255
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
            Height =570
            BackColor =14211288
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =120
                    Width =3825
                    Height =525
                    FontSize =18
                    BorderColor =0
                    ForeColor =8210719
                    Name ="Label465"
                    Caption ="Add / Edit Linked Doc"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedWidth =3945
                    LayoutCachedHeight =525
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Image
                    PictureTiling = NotDefault
                    BackStyle =1
                    PictureType =2
                    Top =420
                    Width =10800
                    Height =150
                    BorderColor =6108695
                    Name ="Image142"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =420
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =570
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
            End
        End
        Begin Section
            Height =4860
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2580
                    Width =7500
                    Height =585
                    ForeColor =1279872587
                    Name ="dLink"
                    ControlSource ="dLink"
                    StatusBarText ="hyperlink to Document shown here"
                    AfterUpdate ="[Event Procedure]"
                    LeftPadding =180
                    TopPadding =0
                    RightPadding =195
                    BottomPadding =0

                    LayoutCachedLeft =2580
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =585
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =4
                    BackTint =20.0
                    BorderThemeColorIndex =2
                    BorderTint =60.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2580
                    Top =1155
                    Width =7500
                    Height =315
                    TabIndex =2
                    Name ="dName"
                    ControlSource ="dName"
                    LeftPadding =180
                    TopPadding =0
                    RightPadding =195
                    BottomPadding =0

                    LayoutCachedLeft =2580
                    LayoutCachedTop =1155
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =1470
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =1560
                            Top =1155
                            Width =1020
                            Height =315
                            Name ="Label0"
                            Caption ="FileName"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1560
                            LayoutCachedTop =1155
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =1470
                            RowStart =2
                            RowEnd =2
                            BackThemeColorIndex =2
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2580
                    Top =1470
                    Width =7500
                    Height =315
                    TabIndex =3
                    Name ="dtype"
                    ControlSource ="dType"
                    LeftPadding =180
                    TopPadding =0
                    RightPadding =195
                    BottomPadding =0

                    LayoutCachedLeft =2580
                    LayoutCachedTop =1470
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =1785
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Left =1560
                            Top =1470
                            Width =1020
                            Height =315
                            Name ="Label1"
                            Caption ="Doc Type"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1560
                            LayoutCachedTop =1470
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =1785
                            RowStart =3
                            RowEnd =3
                            BackThemeColorIndex =2
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2580
                    Top =1785
                    Width =7500
                    Height =765
                    TabIndex =4
                    Name ="dtags"
                    ControlSource ="dTags"
                    LeftPadding =180
                    TopPadding =0
                    RightPadding =195
                    BottomPadding =0

                    LayoutCachedLeft =2580
                    LayoutCachedTop =1785
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =2550
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =95
                            TextAlign =1
                            Left =1560
                            Top =1785
                            Width =1020
                            Height =765
                            Name ="Label2"
                            Caption ="Dir / Tags"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1560
                            LayoutCachedTop =1785
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =2550
                            RowStart =4
                            RowEnd =4
                            BackThemeColorIndex =2
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =2580
                    Top =690
                    Width =7500
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="GetLink"
                    Caption ="Click HERE to Get Link Details"
                    OnClick ="[Event Procedure]"
                    LeftPadding =180
                    TopPadding =105
                    RightPadding =195
                    BottomPadding =165
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

                    LayoutCachedLeft =2580
                    LayoutCachedTop =690
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =990
                    PictureCaptionArrangement =5
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =8
                    BackTint =100.0
                    BackShade =75.0
                    BorderColor =16777215
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =9
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedThemeColorIndex =8
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    Shadow =2
                    QuickStyle =33
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =120
                    Width =2460
                    Height =390
                    FontSize =14
                    FontWeight =700
                    Name ="Label15"
                    Caption ="1: Paste Link HERE ►"
                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =75.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =60
                    Top =660
                    Width =2460
                    Height =390
                    FontSize =14
                    FontWeight =700
                    Name ="Label16"
                    Caption ="2: Then Click HERE ►"
                    LayoutCachedLeft =60
                    LayoutCachedTop =660
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =1050
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =75.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2580
                    Top =2520
                    Width =7500
                    Height =720
                    TabIndex =5
                    Name ="dDesc"
                    ControlSource ="dDesc"
                    LeftPadding =180
                    TopPadding =0
                    RightPadding =195
                    BottomPadding =0

                    LayoutCachedLeft =2580
                    LayoutCachedTop =2520
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =3240
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =223
                            TextAlign =2
                            Left =1560
                            Top =2520
                            Width =1020
                            Height =720
                            Name ="Label18"
                            Caption ="Descp "
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1560
                            LayoutCachedTop =2520
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =3240
                            RowStart =4
                            RowEnd =4
                            BackThemeColorIndex =6
                            BackShade =75.0
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1560
                    Top =3660
                    Width =1020
                    Height =960
                    FontWeight =700
                    TabIndex =6
                    Name ="TryDirectory"
                    Caption =" Search Sch Dir"
                    OnClick ="[Event Procedure]"
                    GroupTable =4
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff000000000000000090a4681b90a468ff90a468ff90a46848 ,
                        0x00000000ffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff000000000000000090a4682190a4682d90a468ff90a468ff ,
                        0x90a46848ffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff90a4688d90a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468ff00000000ffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff90a468f390a468840000000090a4684290a468ff90a468ff ,
                        0x90a4682dffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff90a468fc0000000090a4682d90a468ff90a468ff90a4682d ,
                        0x00000000ffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff000000000000000000000000000000000000000000000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff00000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xffffffffffffffffffffffff0000000000000000000000000000000000000000 ,
                        0x0000000000000000727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffff0000000090a4682d90a468ff90a468ff90a4682d ,
                        0x0000000090a468fc727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xb3b3b3ffffffffffffffffff90a4682d90a468ff90a468ff90a4684200000000 ,
                        0x90a4688490a468f3727272ffffffffffffffffffb3b3b3ffffffffffffffffff ,
                        0xb3b3b3ffffffffff0000000090a468ff90a468ff90a468ff90a468ff90a468ff ,
                        0x90a468ff90a4688a727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272cf7272720f90a4684890a468ff90a468ff90a4682d90a46818 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272c67272720990a4684890a468ff90a468ff90a4681b ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272bd00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =1560
                    LayoutCachedTop =3660
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =4620
                    PictureCaptionArrangement =3
                    LayoutGroup =3
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =75.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =1
                    BackTint =100.0
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =1
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedThemeColorIndex =1
                    HoverForeThemeColorIndex =2
                    HoverForeTint =100.0
                    HoverForeShade =75.0
                    PressedForeThemeColorIndex =2
                    PressedForeTint =100.0
                    PressedForeShade =75.0
                    GroupTable =4
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                End
                Begin Label
                    OverlapFlags =93
                    Left =2760
                    Top =3480
                    Width =1200
                    Height =315
                    FontWeight =700
                    Name ="Label21"
                    Caption ="Search Text"
                    LayoutCachedLeft =2760
                    LayoutCachedTop =3480
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =3795
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =95
                    Left =3960
                    Top =3480
                    Width =3135
                    Height =315
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="Label22"
                    Caption ="(you can change this)"
                    LayoutCachedLeft =3960
                    LayoutCachedTop =3480
                    LayoutCachedWidth =7095
                    LayoutCachedHeight =3795
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2760
                    Top =3720
                    Width =7200
                    Height =900
                    TabIndex =7
                    BackColor =13434879
                    Name ="ScheduleDesc"

                    LayoutCachedLeft =2760
                    LayoutCachedTop =3720
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =4620
                    BackThemeColorIndex =-1
                End
            End
        End
        Begin FormFooter
            Height =1080
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =8400
                    Top =60
                    Width =1920
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="DeleteLink"
                    Caption ="Delete THIS Link"
                    OnClick ="[Event Procedure]"
                    LeftPadding =120
                    TopPadding =105
                    RightPadding =135
                    BottomPadding =120
                    GridlineColor =0

                    LayoutCachedLeft =8400
                    LayoutCachedTop =60
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =5
                    BackTint =100.0
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =9
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedThemeColorIndex =5
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    Shadow =5
                    QuickStyle =33
                    QuickStyleMask =-629
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2376
                    Top =480
                    Width =624
                    Height =195
                    FontSize =8
                    TabIndex =1
                    Name ="ClientID"
                    ControlSource ="ClientID"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2376
                    LayoutCachedTop =480
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =675
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1740
                            Top =480
                            Width =619
                            Height =195
                            FontSize =8
                            Name ="Label9"
                            Caption ="ClientID"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1740
                            LayoutCachedTop =480
                            LayoutCachedWidth =2359
                            LayoutCachedHeight =675
                            LayoutGroup =1
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4872
                    Top =480
                    Width =612
                    Height =195
                    FontSize =8
                    TabIndex =3
                    Name ="SchID"
                    ControlSource ="SchID"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4872
                    LayoutCachedTop =480
                    LayoutCachedWidth =5484
                    LayoutCachedHeight =675
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =4248
                            Top =480
                            Width =612
                            Height =195
                            FontSize =8
                            Name ="Label10"
                            Caption ="SchID"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =4248
                            LayoutCachedTop =480
                            LayoutCachedWidth =4860
                            LayoutCachedHeight =675
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3636
                    Top =480
                    Width =612
                    Height =195
                    FontSize =8
                    TabIndex =2
                    Name ="MLID"
                    ControlSource ="MLID"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3636
                    LayoutCachedTop =480
                    LayoutCachedWidth =4248
                    LayoutCachedHeight =675
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =1
                            Left =3000
                            Top =480
                            Width =624
                            Height =195
                            FontSize =8
                            Name ="Label12"
                            Caption ="MLID"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =3000
                            LayoutCachedTop =480
                            LayoutCachedWidth =3624
                            LayoutCachedHeight =675
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6120
                    Top =480
                    Width =612
                    Height =195
                    FontSize =8
                    TabIndex =4
                    Name ="GroupID"
                    ControlSource ="GroupID"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6120
                    LayoutCachedTop =480
                    LayoutCachedWidth =6732
                    LayoutCachedHeight =675
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =1
                            Left =5484
                            Top =480
                            Width =624
                            Height =195
                            FontSize =8
                            Name ="Label13"
                            Caption ="GroupID"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =5484
                            LayoutCachedTop =480
                            LayoutCachedWidth =6108
                            LayoutCachedHeight =675
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =1
                            GroupTable =3
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7368
                    Top =480
                    Width =720
                    Height =195
                    FontSize =8
                    TabIndex =5
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7368
                    LayoutCachedTop =480
                    LayoutCachedWidth =8088
                    LayoutCachedHeight =675
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =3
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =1
                            Left =6732
                            Top =480
                            Width =624
                            Height =195
                            FontSize =8
                            Name ="Label14"
                            Caption ="UnitID"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =6732
                            LayoutCachedTop =480
                            LayoutCachedWidth =7356
                            LayoutCachedHeight =675
                            ColumnStart =8
                            ColumnEnd =8
                            LayoutGroup =1
                            GroupTable =3
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4380
                    Top =60
                    Width =2220
                    Height =300
                    TabIndex =6
                    ForeColor =0
                    Name ="AddLink"
                    Caption ="Add Another Link"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000a574451e ,
                        0xa57445b7a57445ffa5744596a5744524a57445a4a57445ffa57445ffa57445ff ,
                        0xa57445b7a574451e0000000000000000000000000000000000000000a57445ba ,
                        0xa57445f3a5744596a5744503a57445ffa57445ffa5744512a5744533a5744547 ,
                        0xa57445f3a57445c00000000000000000000000000000000000000000a57445f3 ,
                        0xa574459900000000a5744599a57445ff00000000a57445ffa574459900000000 ,
                        0xa5744599a57445f90000000000000000000000000000000000000000a57445ba ,
                        0xa57445f3a5744547a5744533a5744512a57445ffa57445ff00000000a574454c ,
                        0xa57445f3a57445c00000000000000000000000000000000000000000a574451b ,
                        0xa57445b1a57445ffa57445ffa57445ffa57445a4a5744524a5744596a57445ff ,
                        0xa57445b1a574451b000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x727272ffffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff727272ff727272ff727272ff727272ff727272ff00000000 ,
                        0x727272ffffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff727272ffffffffffffffffffb3b3b3ffffffffff00000000 ,
                        0x727272ffffffffffffffffffb3b3b3ffffffffffffffffffb3b3b3ffffffffff ,
                        0xffffffff727272ff727272ffffffffffffffffffb3b3b3ffffffffff00000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffb3b3b3ffb3b3b3ffb3b3b3ffb3b3b3ff00000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffffffffffb3b3b3ffffffffff00000000 ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ffffffffffffffffffb3b3b3ffffffffff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff00000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff00000000000000000000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4380
                    LayoutCachedTop =60
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =360
                    PictureCaptionArrangement =5
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =1
                    BackTint =100.0
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =1
                    HoverTint =100.0
                    HoverShade =85.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-113
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =300
                    Top =60
                    Width =3060
                    Height =300
                    TabIndex =7
                    ForeColor =0
                    Name ="SaveandClose"
                    Caption =" Save and  Close"
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

                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =360
                    PictureCaptionArrangement =5
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =1
                    BackTint =100.0
                    BackShade =85.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =1
                    HoverTint =100.0
                    HoverShade =85.0
                    PressedThemeColorIndex =1
                    PressedShade =85.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-113
                End
            End
        End
    End
End
CodeBehindForm
' See "DocAddEditLinkForm.cls"
