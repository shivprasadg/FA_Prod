Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6750
    DatasheetFontHeight =11
    ItemSuffix =119
    Left =7965
    Top =3555
    Right =14715
    Bottom =6510
    RecSrcDt = Begin
        0x279eee8e4d9ee440
    End
    Caption ="Available Printes"
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
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
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
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ToggleButton
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackColor =-1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
            BorderLineStyle =0
            BorderColor =-1
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
        Begin Section
            Height =2970
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =105
                    Top =240
                    Width =2115
                    Height =2670
                    FontSize =9
                    TabIndex =4
                    BoundColumn =1
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="PrinterList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MyPrinters.PrinterName AS [Select Printer], MyPrinters.PrinterNameLong, *"
                        " FROM MyPrinters ORDER BY MyPrinters.PrinterName;"
                    ColumnWidths ="3060;0"
                    FontName ="Segoe UI"
                    GridlineColor =10921638

                    LayoutCachedLeft =105
                    LayoutCachedTop =240
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =2910
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    SpecialEffect =5
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2400
                    Top =510
                    Width =3765
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TargetDocument"
                    FontName ="Segoe UI"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =2400
                    LayoutCachedTop =510
                    LayoutCachedWidth =6165
                    LayoutCachedHeight =750
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2400
                    Top =225
                    Width =3765
                    Height =255
                    FontSize =9
                    BorderColor =8355711
                    Name ="Label4"
                    Caption ="Document Name you're about to print:"
                    FontName ="Segoe UI"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2400
                    LayoutCachedTop =225
                    LayoutCachedWidth =6165
                    LayoutCachedHeight =480
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    SpecialEffect =5
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3090
                    Top =810
                    Width =420
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="NumOfCopies"
                    FontName ="Segoe UI"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =3090
                    LayoutCachedTop =810
                    LayoutCachedWidth =3510
                    LayoutCachedHeight =1095
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2400
                    Top =810
                    Width =630
                    Height =285
                    FontSize =9
                    BorderColor =8355711
                    Name ="Label12"
                    Caption ="Copies"
                    FontName ="Segoe UI"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =2400
                    LayoutCachedTop =810
                    LayoutCachedWidth =3030
                    LayoutCachedHeight =1095
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    SpecialEffect =5
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3570
                    Top =1155
                    Width =420
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="PageStart"
                    Format ="General Number"
                    FontName ="Segoe UI"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =3570
                    LayoutCachedTop =1155
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =1440
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3090
                    Top =1155
                    Width =420
                    Height =285
                    FontSize =9
                    BorderColor =8355711
                    Name ="Label27"
                    Caption ="From"
                    FontName ="Segoe UI"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =3090
                    LayoutCachedTop =1155
                    LayoutCachedWidth =3510
                    LayoutCachedHeight =1440
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    SpecialEffect =5
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4395
                    Top =1155
                    Width =420
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="PageEnd"
                    Format ="General Number"
                    FontName ="Segoe UI"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =4395
                    LayoutCachedTop =1155
                    LayoutCachedWidth =4815
                    LayoutCachedHeight =1440
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4050
                    Top =1155
                    Width =285
                    Height =285
                    FontSize =9
                    BorderColor =8355711
                    Name ="Label29"
                    Caption ="To"
                    FontName ="Segoe UI"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4050
                    LayoutCachedTop =1155
                    LayoutCachedWidth =4335
                    LayoutCachedHeight =1440
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2400
                    Top =1155
                    Width =630
                    Height =285
                    FontSize =9
                    BorderColor =8355711
                    Name ="Label54"
                    Caption ="Pages"
                    FontName ="Segoe UI"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =2400
                    LayoutCachedTop =1155
                    LayoutCachedWidth =3030
                    LayoutCachedHeight =1440
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3570
                    Top =810
                    Width =420
                    Height =285
                    Name ="EmptyCell62"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =3570
                    LayoutCachedTop =810
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =1095
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4050
                    Top =810
                    Width =285
                    Height =285
                    Name ="EmptyCell63"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4050
                    LayoutCachedTop =810
                    LayoutCachedWidth =4335
                    LayoutCachedHeight =1095
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4395
                    Top =810
                    Width =420
                    Height =285
                    Name ="EmptyCell64"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4395
                    LayoutCachedTop =810
                    LayoutCachedWidth =4815
                    LayoutCachedHeight =1095
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4845
                    Top =810
                    Width =1320
                    Height =285
                    Name ="EmptyCell65"
                    GroupTable =1
                    LeftPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =4845
                    LayoutCachedTop =810
                    LayoutCachedWidth =6165
                    LayoutCachedHeight =1095
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =5340
                    Top =2460
                    Width =1320
                    Height =285
                    FontSize =9
                    TabIndex =8
                    Name ="CancelPrint"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    LeftPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =5340
                    LayoutCachedTop =2460
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =2745
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ForeTint =100.0
                    Shape =0
                    BackThemeColorIndex =0
                    BackTint =100.0
                    BorderThemeColorIndex =0
                    BorderTint =100.0
                    HoverThemeColorIndex =0
                    HoverTint =80.0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =2
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    Overlaps =1
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =2355
                    Top =1605
                    Width =2415
                    Height =555
                    TabIndex =5
                    BorderColor =10921638
                    Name ="PickOrientation"
                    DefaultValue ="2"
                    GridlineColor =10921638

                    LayoutCachedLeft =2355
                    LayoutCachedTop =1605
                    LayoutCachedWidth =4770
                    LayoutCachedHeight =2160
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =2400
                            Top =1500
                            Width =1155
                            Height =210
                            FontSize =9
                            BorderColor =8355711
                            Name ="OrientationLabel"
                            Caption ="Orientation"
                            FontName ="Segoe UI"
                            GridlineColor =10921638
                            LayoutCachedLeft =2400
                            LayoutCachedTop =1500
                            LayoutCachedWidth =3555
                            LayoutCachedHeight =1710
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =95
                            Left =2445
                            Top =1800
                            Width =285
                            Height =315
                            OptionValue =2
                            BorderColor =10921638
                            Name ="PortraitOption"
                            GroupTable =3
                            TopPadding =0
                            BottomPadding =0
                            GridlineColor =10921638

                            LayoutCachedLeft =2445
                            LayoutCachedTop =1800
                            LayoutCachedWidth =2730
                            LayoutCachedHeight =2115
                            LayoutGroup =3
                            GroupTable =3
                        End
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =2760
                            Top =1800
                            Width =690
                            Height =315
                            FontSize =9
                            BorderColor =8355711
                            Name ="Label85"
                            Caption ="Portrait"
                            FontName ="Segoe UI"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =2760
                            LayoutCachedTop =1800
                            LayoutCachedWidth =3450
                            LayoutCachedHeight =2115
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =3
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =3
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =95
                            Left =3450
                            Top =1800
                            Width =285
                            Height =315
                            OptionValue =1
                            BorderColor =10921638
                            Name ="LandscapeOption"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638

                            LayoutCachedLeft =3450
                            LayoutCachedTop =1800
                            LayoutCachedWidth =3735
                            LayoutCachedHeight =2115
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =3
                            GroupTable =3
                        End
                        Begin Label
                            OverlapFlags =87
                            TextAlign =1
                            Left =3735
                            Top =1800
                            Width =915
                            Height =315
                            FontSize =9
                            BorderColor =8355711
                            Name ="Label87"
                            Caption ="Landscape"
                            FontName ="Segoe UI"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =3735
                            LayoutCachedTop =1800
                            LayoutCachedWidth =4650
                            LayoutCachedHeight =2115
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =3
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =3
                        End
                        Begin EmptyCell
                            Left =2445
                            Top =2115
                            Width =285
                            Height =45
                            Name ="EmptyCell106"
                            GroupTable =3
                            TopPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =2445
                            LayoutCachedTop =2115
                            LayoutCachedWidth =2730
                            LayoutCachedHeight =2160
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =3
                            GroupTable =3
                        End
                        Begin EmptyCell
                            Left =2760
                            Top =2115
                            Width =690
                            Height =45
                            Name ="EmptyCell111"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =2760
                            LayoutCachedTop =2115
                            LayoutCachedWidth =3450
                            LayoutCachedHeight =2160
                            RowStart =1
                            RowEnd =1
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =3
                            GroupTable =3
                        End
                        Begin EmptyCell
                            Left =3450
                            Top =2115
                            Width =285
                            Height =45
                            Name ="EmptyCell108"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =3450
                            LayoutCachedTop =2115
                            LayoutCachedWidth =3735
                            LayoutCachedHeight =2160
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =3
                            GroupTable =3
                        End
                        Begin EmptyCell
                            Left =3735
                            Top =2115
                            Width =915
                            Height =45
                            Name ="EmptyCell110"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            GridlineColor =10921638
                            LayoutCachedLeft =3735
                            LayoutCachedTop =2115
                            LayoutCachedWidth =4650
                            LayoutCachedHeight =2160
                            RowStart =1
                            RowEnd =1
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =3
                            GroupTable =3
                        End
                    End
                End
                Begin EmptyCell
                    Left =4845
                    Top =1155
                    Width =1320
                    Height =285
                    Name ="EmptyCell88"
                    GroupTable =1
                    LeftPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =4845
                    LayoutCachedTop =1155
                    LayoutCachedWidth =6165
                    LayoutCachedHeight =1440
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =2400
                    Top =2460
                    Height =285
                    FontSize =9
                    TabIndex =6
                    Name ="PrintIt"
                    Caption ="Print"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    LeftPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =2400
                    LayoutCachedTop =2460
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =2745
                    LayoutGroup =2
                    ForeTint =100.0
                    Shape =0
                    BackThemeColorIndex =0
                    BackTint =100.0
                    BorderThemeColorIndex =0
                    BorderTint =100.0
                    HoverThemeColorIndex =0
                    HoverTint =80.0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =2
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3870
                    Top =2460
                    Height =285
                    FontSize =9
                    TabIndex =7
                    Name ="Preview"
                    Caption ="Preview"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    LeftPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =3870
                    LayoutCachedTop =2460
                    LayoutCachedWidth =5310
                    LayoutCachedHeight =2745
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeTint =100.0
                    Shape =0
                    BackThemeColorIndex =0
                    BackTint =100.0
                    BorderThemeColorIndex =0
                    BorderTint =100.0
                    HoverThemeColorIndex =0
                    HoverTint =80.0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    GroupTable =2
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
' See "Printers.cls"
