Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12615
    DatasheetFontHeight =11
    ItemSuffix =20
    Left =1320
    Top =2190
    Right =15720
    Bottom =5835
    RecSrcDt = Begin
        0x075672bc7306e540
    End
    RecordSource ="Select * FROM vw_LocationAddressMatrix WHERE ClientGroupID=44 Order By Lstate, L"
        "City"
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
        Begin ComboBox
            AddColon = NotDefault
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
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ToggleButton
            ForeColor =4210752
            FontName ="Calibri"
            GridlineColor =10921638
            ForeTint =75.0
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackColor =-4357809
            OldBorderStyle =0
            BorderLineStyle =0
            BorderColor =-4357809
            ThemeFontIndex =1
            HoverColor =-4357809
            HoverTint =80.0
            PressedColor =-4357809
            PressedShade =80.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =815
            BackColor =14211288
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =4950
                    Top =285
                    Width =1800
                    Height =270
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label1"
                    Caption ="City"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    LayoutCachedLeft =4950
                    LayoutCachedTop =285
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =555
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =6750
                    Top =285
                    Width =570
                    Height =270
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label2"
                    Caption ="State"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    LayoutCachedLeft =6750
                    LayoutCachedTop =285
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =555
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =2505
                    Top =285
                    Width =2445
                    Height =270
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label3"
                    Caption ="Address"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    LayoutCachedLeft =2505
                    LayoutCachedTop =285
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =555
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =60
                    Top =285
                    Width =2475
                    Height =270
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label9"
                    Caption ="  Names Used at Address"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    LayoutCachedLeft =60
                    LayoutCachedTop =285
                    LayoutCachedWidth =2535
                    LayoutCachedHeight =555
                    ColumnStart =8
                    ColumnEnd =8
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =8085
                    Top =285
                    Width =2280
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label297"
                    Caption ="Type of Units @ Locaation"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    LayoutCachedLeft =8085
                    LayoutCachedTop =285
                    LayoutCachedWidth =10365
                    LayoutCachedHeight =540
                    ColumnStart =8
                    ColumnEnd =8
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =223
                    Left =7275
                    Top =510
                    Width =3140
                    Height =290
                    ColumnOrder =2
                    BackColor =15527148
                    BorderColor =0
                    Name ="LocationFilter1"
                    DefaultValue ="1"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7275
                    LayoutCachedTop =510
                    LayoutCachedWidth =10415
                    LayoutCachedHeight =800
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    Begin
                        Begin ToggleButton
                            OverlapFlags =223
                            Left =8040
                            Top =525
                            Width =475
                            Height =255
                            FontSize =8
                            FontWeight =400
                            OptionValue =1
                            Name ="CheckLeased"
                            Caption ="Dom"
                            ControlTipText ="Leased 'A' or 'B' Units"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120
                            GridlineColor =0

                            LayoutCachedLeft =8040
                            LayoutCachedTop =525
                            LayoutCachedWidth =8515
                            LayoutCachedHeight =780
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =95.0
                            GridlineShade =100.0
                            Shape =0
                            Bevel =0
                            BackColor =8210719
                            OldBorderStyle =1
                            BorderThemeColorIndex =4
                            BorderTint =60.0
                            ThemeFontIndex =-1
                            HoverThemeColorIndex =4
                            HoverTint =40.0
                            PressedThemeColorIndex =2
                            PressedShade =100.0
                            HoverForeThemeColorIndex =2
                            HoverForeShade =75.0
                            PressedForeThemeColorIndex =1
                            QuickStyle =30
                            QuickStyleMask =-629
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =223
                            Left =8515
                            Top =525
                            Width =475
                            Height =255
                            FontSize =8
                            FontWeight =400
                            TabIndex =1
                            OptionValue =2
                            Name ="CheckAll"
                            Caption ="Park"
                            ControlTipText ="All Units"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120
                            GridlineColor =0

                            LayoutCachedLeft =8515
                            LayoutCachedTop =525
                            LayoutCachedWidth =8990
                            LayoutCachedHeight =780
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =95.0
                            GridlineShade =100.0
                            Shape =0
                            Bevel =0
                            BackColor =8210719
                            OldBorderStyle =1
                            BorderThemeColorIndex =4
                            BorderTint =60.0
                            ThemeFontIndex =-1
                            HoverThemeColorIndex =4
                            HoverTint =40.0
                            PressedThemeColorIndex =2
                            PressedShade =100.0
                            HoverForeThemeColorIndex =2
                            HoverForeShade =75.0
                            PressedForeThemeColorIndex =1
                            QuickStyle =30
                            QuickStyleMask =-629
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =223
                            Left =8990
                            Top =525
                            Width =475
                            Height =255
                            FontSize =8
                            FontWeight =400
                            TabIndex =2
                            OptionValue =3
                            Name ="CheckRemarketing"
                            Caption ="Regd"
                            ControlTipText ="Remarketing Units"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120
                            GridlineColor =0

                            LayoutCachedLeft =8990
                            LayoutCachedTop =525
                            LayoutCachedWidth =9465
                            LayoutCachedHeight =780
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =95.0
                            GridlineShade =100.0
                            Shape =0
                            Bevel =0
                            BackColor =8210719
                            OldBorderStyle =1
                            BorderThemeColorIndex =4
                            BorderTint =60.0
                            ThemeFontIndex =-1
                            HoverThemeColorIndex =4
                            HoverTint =40.0
                            PressedThemeColorIndex =2
                            PressedShade =100.0
                            HoverForeThemeColorIndex =2
                            HoverForeShade =75.0
                            PressedForeThemeColorIndex =1
                            QuickStyle =30
                            QuickStyleMask =-629
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =223
                            Left =9465
                            Top =525
                            Width =475
                            Height =255
                            FontSize =8
                            FontWeight =400
                            TabIndex =3
                            OptionValue =4
                            Name ="Check51"
                            Caption ="Delv"
                            ControlTipText ="Units To Be Returned"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120
                            GridlineColor =0

                            LayoutCachedLeft =9465
                            LayoutCachedTop =525
                            LayoutCachedWidth =9940
                            LayoutCachedHeight =780
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =95.0
                            GridlineShade =100.0
                            Shape =0
                            Bevel =0
                            BackColor =8210719
                            OldBorderStyle =1
                            BorderThemeColorIndex =4
                            BorderTint =60.0
                            ThemeFontIndex =-1
                            HoverThemeColorIndex =4
                            HoverTint =40.0
                            PressedThemeColorIndex =2
                            PressedShade =100.0
                            HoverForeThemeColorIndex =2
                            HoverForeShade =75.0
                            PressedForeThemeColorIndex =1
                            QuickStyle =30
                            QuickStyleMask =-629
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =215
                            Left =9940
                            Top =525
                            Width =475
                            Height =255
                            FontSize =8
                            FontWeight =400
                            TabIndex =4
                            OptionValue =5
                            Name ="Toggle0"
                            Caption ="Title"
                            ControlTipText ="Tracked Units Only"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120
                            GridlineColor =0

                            LayoutCachedLeft =9940
                            LayoutCachedTop =525
                            LayoutCachedWidth =10415
                            LayoutCachedHeight =780
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =95.0
                            GridlineShade =100.0
                            Shape =0
                            Bevel =0
                            BackColor =8210719
                            OldBorderStyle =1
                            BorderThemeColorIndex =4
                            BorderTint =60.0
                            ThemeFontIndex =-1
                            HoverThemeColorIndex =4
                            HoverTint =40.0
                            PressedThemeColorIndex =2
                            PressedShade =100.0
                            HoverForeThemeColorIndex =2
                            HoverForeShade =75.0
                            PressedForeThemeColorIndex =1
                            QuickStyle =30
                            QuickStyleMask =-629
                            Overlaps =1
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    Left =9780
                    Top =525
                    Width =720
                    Height =290
                    ColumnOrder =4
                    TabIndex =2
                    BackColor =15527148
                    BorderColor =0
                    Name ="LocationFilter2"
                    DefaultValue ="1"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9780
                    LayoutCachedTop =525
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =815
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin ComboBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =540
                    Width =2445
                    Height =255
                    ColumnOrder =6
                    FontSize =9
                    TabIndex =4
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchName"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LocationName FROM Locations WHERE (((Locations.ClientGroupID)=["
                        "Forms]![LocationAddressesEdit]![ClientGroupID])) ORDER BY Locations.LocationName"
                        ";"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =540
                    LayoutCachedWidth =2505
                    LayoutCachedHeight =795
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                End
                Begin ComboBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2505
                    Top =540
                    Width =2445
                    Height =255
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =3
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchAddress"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LAddrLine1 FROM Locations WHERE (((Locations.ClientGroupID)=[Fo"
                        "rms]![LocationAddressesEdit]![ClientGroupID])) GROUP BY Locations.LAddrLine1 ORD"
                        "ER BY Locations.LAddrLine1;"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2505
                    LayoutCachedTop =540
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =795
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                End
                Begin ComboBox
                    OverlapFlags =223
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6750
                    Top =540
                    Width =576
                    Height =255
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =5
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LState FROM Locations WHERE (((ClientGroupID)=[Forms]![Location"
                        "AddressesEdit]![ClientGroupID])) GROUP BY Locations.LState ORDER BY Locations.LS"
                        "tate;"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6750
                    LayoutCachedTop =540
                    LayoutCachedWidth =7326
                    LayoutCachedHeight =795
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4950
                    Top =540
                    Width =1800
                    Height =255
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =1
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchCity"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Locations.LCity FROM Locations WHERE (((ClientGroupID)=[Forms]![LocationA"
                        "ddressesEdit]![ClientGroupID])) GROUP BY Locations.LCity ORDER BY Locations.LCit"
                        "y;"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4950
                    LayoutCachedTop =540
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =795
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =60
                    Width =11070
                    Height =225
                    FontSize =9
                    BorderColor =0
                    Name ="Label88"
                    Caption ="  ┌───           ───────────┬─────────────┬──────────────────────────┬──────────"
                        "───────────────┬───────────────────────────┐"
                    GridlineColor =0
                    LayoutCachedLeft =60
                    LayoutCachedWidth =11130
                    LayoutCachedHeight =225
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9060
                    Width =576
                    ColumnOrder =1
                    FontSize =8
                    TabIndex =6
                    Name ="ClientGroupID"
                    ControlSource ="ClientGroupID"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9060
                    LayoutCachedWidth =9636
                    LayoutCachedHeight =240
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =540
                    Width =795
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    ForeColor =8210719
                    Name ="ClearFilters"
                    Caption ="Clear"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000003255d6ab3255d65a0000000000000000000000000000000000000000 ,
                        0x3255d62d3255d693000000000000000000000000000000000000000000000000 ,
                        0x000000003255d6ae3255d6f93255d6360000000000000000000000003255d62d ,
                        0x3255d6db3255d61e000000000000000000000000000000000000000000000000 ,
                        0x000000003255d6153255d6db3255d6f03255d630000000003255d6303255d6ea ,
                        0x3255d66300000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000003255d61b3255d6c33255d6ed3255d66f3255d6ea3255d6ae ,
                        0x00000000000000000000000000000000000000000000000000000000727272ff ,
                        0x727272ff727272ff000000003255d6033255d6c63255d6ff3255d6de3255d60c ,
                        0x00000000000000000000000000000000000000000000000000000000727272ff ,
                        0x7272727e000000003255d6153255d6ab3255d6ff3255d6cf3255d6bd3255d696 ,
                        0x3255d609000000000000000000000000000000000000000000000000727272ff ,
                        0x000000003255d64e3255d6ed3255d6ff3255d6b73255d60c000000003255d645 ,
                        0x3255d6a53255d6420000000000000000000000000000000000000000727272ff ,
                        0x000000003255d6753255d6de3255d65a00000000000000000000000000000000 ,
                        0x000000003255d6270000000000000000000000000000000000000000727272ff ,
                        0x7272728100000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000007272726c727272ff ,
                        0x727272ff727272ff727272780000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000007272724e727272fc727272ff ,
                        0x727272ff727272ff727272ff7272725a00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000072727236727272f6727272ff727272ff ,
                        0x727272ff727272ff727272ff727272f972727242000000000000000000000000 ,
                        0x00000000000000000000000072727224727272ea727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272f07272722d0000000000000000 ,
                        0x000000000000000000000000727272d2727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272e40000000000000000 ,
                        0x000000000000000000000000727272f0727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =540
                    LayoutCachedWidth =1335
                    LayoutCachedHeight =255
                    PictureCaptionArrangement =5
                    Alignment =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =0
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =12566463
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =8872517
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =8210719
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =8210719
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-885
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7860
                    Width =576
                    ColumnOrder =0
                    FontSize =8
                    TabIndex =8
                    Name ="txtCurrent"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7860
                    LayoutCachedWidth =8436
                    LayoutCachedHeight =240
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =10545
                    Top =285
                    Width =576
                    Height =270
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label12"
                    Caption ="Tax%"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    LayoutCachedLeft =10545
                    LayoutCachedTop =285
                    LayoutCachedWidth =11121
                    LayoutCachedHeight =555
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =11115
                    Top =285
                    Width =360
                    Height =270
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label15"
                    Caption =" RR"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    LayoutCachedLeft =11115
                    LayoutCachedTop =285
                    LayoutCachedWidth =11475
                    LayoutCachedHeight =555
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =7380
                    Top =525
                    Width =595
                    Height =270
                    FontSize =8
                    FontWeight =700
                    Name ="Label19"
                    Caption ="Type"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    LayoutCachedLeft =7380
                    LayoutCachedTop =525
                    LayoutCachedWidth =7975
                    LayoutCachedHeight =795
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =6
                    BackShade =75.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =269
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =9150
                    Width =345
                    FontSize =8
                    TabIndex =1
                    Name ="LocationID"
                    ControlSource ="LocationID"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9150
                    LayoutCachedWidth =9495
                    LayoutCachedHeight =240
                    ColumnStart =14
                    ColumnEnd =14
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6750
                    Width =576
                    FontSize =8
                    Name ="LState"
                    ControlSource ="LState"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6750
                    LayoutCachedWidth =7326
                    LayoutCachedHeight =240
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2505
                    Width =2445
                    FontSize =8
                    TabIndex =2
                    Name ="LAddrLine1"
                    ControlSource ="LAddrLine1"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2505
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =240
                    ColumnStart =2
                    ColumnEnd =2
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8040
                    Width =475
                    FontSize =8
                    TabIndex =3
                    Name ="FAUnit"
                    ControlSource ="Dom"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8040
                    LayoutCachedWidth =8515
                    LayoutCachedHeight =240
                    ColumnStart =3
                    ColumnEnd =3
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8515
                    Width =475
                    FontSize =8
                    TabIndex =4
                    Name ="Track"
                    ControlSource ="Park"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8515
                    LayoutCachedWidth =8990
                    LayoutCachedHeight =240
                    ColumnStart =4
                    ColumnEnd =4
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8990
                    Width =475
                    FontSize =8
                    TabIndex =5
                    Name ="Regd"
                    ControlSource ="Regd"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8990
                    LayoutCachedWidth =9465
                    LayoutCachedHeight =240
                    ColumnStart =5
                    ColumnEnd =5
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9465
                    Width =475
                    FontSize =8
                    TabIndex =6
                    Name ="Delv"
                    ControlSource ="Delv"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9465
                    LayoutCachedWidth =9940
                    LayoutCachedHeight =240
                    ColumnStart =6
                    ColumnEnd =6
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =127
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9940
                    Width =475
                    FontSize =8
                    TabIndex =7
                    Name ="Titled"
                    ControlSource ="Titled"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9940
                    LayoutCachedWidth =10415
                    LayoutCachedHeight =240
                    ColumnStart =7
                    ColumnEnd =7
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =60
                    Width =2445
                    FontSize =8
                    TabIndex =8
                    Name ="LocationName"
                    ControlSource ="LocationName"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =60
                    LayoutCachedWidth =2505
                    LayoutCachedHeight =240
                    ColumnStart =8
                    ColumnEnd =8
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4950
                    Width =1800
                    FontSize =8
                    TabIndex =9
                    Name ="LCity"
                    ControlSource ="LCity"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4950
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =240
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =127
                    IMESentenceMode =3
                    Top =240
                    Width =12600
                    Height =29
                    FontSize =10
                    TabIndex =10
                    BorderColor =12632256
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x0100000096000000010000000100000000000000000000001a00000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004c006f0063006100740069006f006e00490044005d003d005b0074007800 ,
                        0x7400430075007200720065006e0074005d0000000000
                    End
                    GridlineColor =0

                    LayoutCachedTop =240
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =269
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00190000005b00 ,
                        0x4c006f0063006100740069006f006e00490044005d003d005b00740078007400 ,
                        0x430075007200720065006e0074005d0000000000000000000000000000000000 ,
                        0x0000000000
                    End
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    DecimalPlaces =3
                    OldBorderStyle =0
                    OverlapFlags =127
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10500
                    Width =621
                    FontSize =8
                    TabIndex =11
                    BackColor =14151142
                    ForeColor =0
                    Name ="SalesTaxRate"
                    ControlSource ="SalesTaxRate"
                    Format ="Percent"
                    OnClick ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10500
                    LayoutCachedWidth =11121
                    LayoutCachedHeight =240
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =119
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11115
                    Width =360
                    FontSize =8
                    TabIndex =12
                    Name ="RentReceipt"
                    ControlSource ="RentReceipt"
                    FontName ="Segoe UI"
                    OnClick ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11115
                    LayoutCachedWidth =11475
                    LayoutCachedHeight =240
                    ColumnStart =9
                    ColumnEnd =9
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =119
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7380
                    Width =595
                    FontSize =8
                    TabIndex =13
                    Name ="LType"
                    ControlSource ="LType"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7380
                    LayoutCachedWidth =7975
                    LayoutCachedHeight =240
                    ColumnStart =3
                    ColumnEnd =3
                    ThemeFontIndex =-1
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
' See "LocationAddressList.cls"
