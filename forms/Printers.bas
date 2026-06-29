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
            TextFontCharSet =177
            TextFontFamily =0
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
                    Name ="PrinterList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MyPrinters.PrinterName AS [Select Printer], MyPrinters.PrinterNameLong, *"
                        " FROM MyPrinters ORDER BY MyPrinters.PrinterName;"
                    ColumnWidths ="3060;0"
                    FontName ="Segoe UI"

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
                    Top =516
                    Width =3780
                    FontSize =9
                    Name ="TargetDocument"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =2400
                    LayoutCachedTop =516
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =756
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
                    Top =228
                    Width =3780
                    Height =252
                    FontSize =9
                    Name ="Label4"
                    Caption ="Document Name you're about to print:"
                    FontName ="Segoe UI"
                    GroupTable =1
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2400
                    LayoutCachedTop =228
                    LayoutCachedWidth =6180
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
                    Left =3096
                    Top =816
                    Width =420
                    Height =288
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="NumOfCopies"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =3096
                    LayoutCachedTop =816
                    LayoutCachedWidth =3516
                    LayoutCachedHeight =1104
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2400
                            Top =816
                            Width =630
                            Height =288
                            FontSize =9
                            Name ="Label12"
                            Caption ="Copies"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =2400
                            LayoutCachedTop =816
                            LayoutCachedWidth =3030
                            LayoutCachedHeight =1104
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =5
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3576
                    Top =1164
                    Width =420
                    Height =288
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="PageStart"
                    Format ="General Number"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =3576
                    LayoutCachedTop =1164
                    LayoutCachedWidth =3996
                    LayoutCachedHeight =1452
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =3096
                            Top =1164
                            Width =420
                            Height =288
                            FontSize =9
                            Name ="Label27"
                            Caption ="From"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =3096
                            LayoutCachedTop =1164
                            LayoutCachedWidth =3516
                            LayoutCachedHeight =1452
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
                    End
                End
                Begin TextBox
                    SpecialEffect =5
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4404
                    Top =1164
                    Width =420
                    Height =288
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="PageEnd"
                    Format ="General Number"
                    FontName ="Segoe UI"
                    GroupTable =1

                    LayoutCachedLeft =4404
                    LayoutCachedTop =1164
                    LayoutCachedWidth =4824
                    LayoutCachedHeight =1452
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =4056
                            Top =1164
                            Width =288
                            Height =288
                            FontSize =9
                            Name ="Label29"
                            Caption ="To"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LayoutCachedLeft =4056
                            LayoutCachedTop =1164
                            LayoutCachedWidth =4344
                            LayoutCachedHeight =1452
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
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2400
                    Top =1164
                    Width =630
                    Height =288
                    FontSize =9
                    Name ="Label54"
                    Caption ="Pages"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LayoutCachedLeft =2400
                    LayoutCachedTop =1164
                    LayoutCachedWidth =3030
                    LayoutCachedHeight =1452
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =3576
                    Top =816
                    Width =420
                    Height =288
                    Name ="EmptyCell62"
                    GroupTable =1
                    LayoutCachedLeft =3576
                    LayoutCachedTop =816
                    LayoutCachedWidth =3996
                    LayoutCachedHeight =1104
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4056
                    Top =816
                    Width =288
                    Height =288
                    Name ="EmptyCell63"
                    GroupTable =1
                    LayoutCachedLeft =4056
                    LayoutCachedTop =816
                    LayoutCachedWidth =4344
                    LayoutCachedHeight =1104
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4404
                    Top =816
                    Width =420
                    Height =288
                    Name ="EmptyCell64"
                    GroupTable =1
                    LayoutCachedLeft =4404
                    LayoutCachedTop =816
                    LayoutCachedWidth =4824
                    LayoutCachedHeight =1104
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =4860
                    Top =816
                    Width =1320
                    Height =288
                    Name ="EmptyCell65"
                    GroupTable =1
                    LeftPadding =0
                    LayoutCachedLeft =4860
                    LayoutCachedTop =816
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =1104
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =5352
                    Top =2460
                    Width =1320
                    Height =285
                    FontSize =9
                    TabIndex =8
                    ForeColor =0
                    Name ="CancelPrint"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    LeftPadding =0

                    LayoutCachedLeft =5352
                    LayoutCachedTop =2460
                    LayoutCachedWidth =6672
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
                    Name ="PickOrientation"
                    DefaultValue ="2"

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
                            Name ="OrientationLabel"
                            Caption ="Orientation"
                            FontName ="Segoe UI"
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
                            Left =2448
                            Top =1800
                            Width =288
                            Height =315
                            OptionValue =2
                            Name ="PortraitOption"
                            GroupTable =3
                            TopPadding =0
                            BottomPadding =0

                            LayoutCachedLeft =2448
                            LayoutCachedTop =1800
                            LayoutCachedWidth =2736
                            LayoutCachedHeight =2115
                            LayoutGroup =3
                            GroupTable =3
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =95
                            Left =3468
                            Top =1800
                            Width =288
                            Height =315
                            OptionValue =1
                            Name ="LandscapeOption"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0

                            LayoutCachedLeft =3468
                            LayoutCachedTop =1800
                            LayoutCachedWidth =3756
                            LayoutCachedHeight =2115
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =3
                            GroupTable =3
                            Begin
                                Begin Label
                                    OverlapFlags =95
                                    TextAlign =1
                                    Left =2772
                                    Top =1800
                                    Width =696
                                    Height =315
                                    FontSize =9
                                    Name ="Label85"
                                    Caption ="Portrait"
                                    FontName ="Segoe UI"
                                    GroupTable =3
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =2772
                                    LayoutCachedTop =1800
                                    LayoutCachedWidth =3468
                                    LayoutCachedHeight =2115
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =3
                                    ThemeFontIndex =-1
                                    ForeThemeColorIndex =2
                                    ForeTint =100.0
                                    GroupTable =3
                                End
                            End
                        End
                        Begin Label
                            OverlapFlags =87
                            TextAlign =1
                            Left =3756
                            Top =1800
                            Width =912
                            Height =315
                            FontSize =9
                            Name ="Label87"
                            Caption ="Landscape"
                            FontName ="Segoe UI"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =3756
                            LayoutCachedTop =1800
                            LayoutCachedWidth =4668
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
                            Left =2448
                            Top =2112
                            Width =288
                            Height =48
                            Name ="EmptyCell106"
                            GroupTable =3
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =2448
                            LayoutCachedTop =2112
                            LayoutCachedWidth =2736
                            LayoutCachedHeight =2160
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =3
                            GroupTable =3
                        End
                        Begin EmptyCell
                            Left =2772
                            Top =2112
                            Width =696
                            Height =48
                            Name ="EmptyCell111"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =2772
                            LayoutCachedTop =2112
                            LayoutCachedWidth =3468
                            LayoutCachedHeight =2160
                            RowStart =1
                            RowEnd =1
                            ColumnStart =1
                            ColumnEnd =1
                            LayoutGroup =3
                            GroupTable =3
                        End
                        Begin EmptyCell
                            Left =3468
                            Top =2112
                            Width =288
                            Height =48
                            Name ="EmptyCell108"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =3468
                            LayoutCachedTop =2112
                            LayoutCachedWidth =3756
                            LayoutCachedHeight =2160
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =3
                            GroupTable =3
                        End
                        Begin EmptyCell
                            Left =3756
                            Top =2112
                            Width =912
                            Height =48
                            Name ="EmptyCell110"
                            GroupTable =3
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =3756
                            LayoutCachedTop =2112
                            LayoutCachedWidth =4668
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
                    Left =4860
                    Top =1164
                    Width =1320
                    Height =288
                    Name ="EmptyCell88"
                    GroupTable =1
                    LeftPadding =0
                    LayoutCachedLeft =4860
                    LayoutCachedTop =1164
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =1452
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
                    ForeColor =0
                    Name ="PrintIt"
                    Caption ="Print"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    LeftPadding =0

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
                    Left =3876
                    Top =2460
                    Height =285
                    FontSize =9
                    TabIndex =7
                    ForeColor =0
                    Name ="Preview"
                    Caption ="Preview"
                    OnClick ="[Event Procedure]"
                    GroupTable =2
                    LeftPadding =0

                    LayoutCachedLeft =3876
                    LayoutCachedTop =2460
                    LayoutCachedWidth =5316
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
