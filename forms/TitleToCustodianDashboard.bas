Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =19208
    DatasheetFontHeight =11
    ItemSuffix =313
    Left =4875
    Top =4110
    Right =21765
    Bottom =8265
    RecSrcDt = Begin
        0xb62585344d20e540
    End
    RecordSource ="Select * FROM vw_TitleToCustodian_Pending Order By Client"
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
            Height =1335
            BackColor =14211288
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin ComboBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2940
                    Top =1080
                    Width =1350
                    Height =255
                    ColumnOrder =17
                    FontSize =9
                    TabIndex =21
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchSchNo"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;1440"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2940
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =300
                    Top =558
                    Width =18360
                    Height =225
                    FontSize =9
                    BorderColor =0
                    Name ="Label88"
                    Caption ="  ┌──────────────────────────────────────────────────────────────────────────┴──"
                        "────────────────────────────────────────────────────────────────────────────────"
                        "─────────────────────────────────────┐"
                    GridlineColor =0
                    LayoutCachedLeft =300
                    LayoutCachedTop =558
                    LayoutCachedWidth =18660
                    LayoutCachedHeight =783
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
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7860
                    Width =576
                    ColumnOrder =15
                    FontSize =8
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
                    TextAlign =2
                    Left =5220
                    Top =1080
                    Width =1005
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label3"
                    Caption ="├  Title Log ┤"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5220
                    LayoutCachedTop =1080
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4290
                    Top =1080
                    Width =930
                    Height =255
                    ColumnOrder =16
                    FontSize =9
                    TabIndex =22
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchGroup"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;1440"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4290
                    LayoutCachedTop =1080
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =6225
                    Top =795
                    Height =270
                    ColumnOrder =14
                    FontSize =8
                    FontWeight =400
                    TabIndex =12
                    ForeColor =8210719
                    Name ="SortDaysElapsed1"
                    DefaultValue ="0"
                    Caption ="Days"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =6225
                    LayoutCachedTop =795
                    LayoutCachedWidth =6945
                    LayoutCachedHeight =1065
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    GroupTable =1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =2940
                    Top =795
                    Width =1350
                    Height =270
                    ColumnOrder =13
                    FontSize =8
                    FontWeight =400
                    TabIndex =9
                    ForeColor =8210719
                    Name ="SortSchNo1"
                    DefaultValue ="0"
                    Caption ="Schedule"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =2940
                    LayoutCachedTop =795
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =1065
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    GroupTable =1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =5220
                    Top =795
                    Width =1005
                    Height =270
                    ColumnOrder =11
                    FontSize =8
                    FontWeight =400
                    TabIndex =11
                    ForeColor =8210719
                    Name ="TitleLog1"
                    DefaultValue ="0"
                    Caption ="Last"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =5220
                    LayoutCachedTop =795
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =1065
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    GroupTable =1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =4290
                    Top =795
                    Width =930
                    Height =270
                    ColumnOrder =12
                    FontSize =8
                    FontWeight =400
                    TabIndex =10
                    ForeColor =8210719
                    Name ="SortUnitGroup1"
                    DefaultValue ="0"
                    Caption ="Group"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =4290
                    LayoutCachedTop =795
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =1065
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    GroupTable =1
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =60
                    Top =1080
                    Width =2175
                    Height =255
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =20
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchClient"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_TitleToCustodian_Pending.ClientID, vw_TitleToCustodian_Pending.Client "
                        "FROM vw_TitleToCustodian_Pending GROUP BY vw_TitleToCustodian_Pending.ClientID, "
                        "vw_TitleToCustodian_Pending.Client ORDER BY vw_TitleToCustodian_Pending.Client;"
                    ColumnWidths ="0;1440"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =1080
                    LayoutCachedWidth =2235
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =60
                    Top =795
                    Width =2175
                    Height =270
                    ColumnOrder =9
                    FontSize =8
                    FontWeight =400
                    TabIndex =7
                    ForeColor =8210719
                    Name ="SortClient1"
                    DefaultValue ="0"
                    Caption ="Client"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =60
                    LayoutCachedTop =795
                    LayoutCachedWidth =2235
                    LayoutCachedHeight =1065
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    GroupTable =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =9045
                    Top =1080
                    Width =690
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label94"
                    Caption ="Unit"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9045
                    LayoutCachedTop =1080
                    LayoutCachedWidth =9735
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =10245
                    Top =1080
                    Width =4035
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label103"
                    Caption ="Descp"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10245
                    LayoutCachedTop =1080
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =9045
                    Top =795
                    Width =690
                    Height =270
                    ColumnOrder =8
                    FontSize =8
                    FontWeight =400
                    TabIndex =14
                    ForeColor =8210719
                    Name ="SortAssetID1"
                    DefaultValue ="0"
                    Caption ="AssetID"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =9045
                    LayoutCachedTop =795
                    LayoutCachedWidth =9735
                    LayoutCachedHeight =1065
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    GroupTable =1
                    Overlaps =1
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =14280
                    Top =1080
                    Width =1995
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label123"
                    Caption ="Double-Click to Open Unit"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14280
                    LayoutCachedTop =1080
                    LayoutCachedWidth =16275
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =10245
                    Top =795
                    Width =4035
                    Height =270
                    ColumnOrder =7
                    FontSize =8
                    FontWeight =400
                    TabIndex =16
                    ForeColor =8210719
                    Name ="SortUnitDesc1"
                    DefaultValue ="0"
                    Caption ="Unit Type"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =10245
                    LayoutCachedTop =795
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =1065
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    GroupTable =1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =6945
                    Top =795
                    Width =2100
                    Height =270
                    ColumnOrder =6
                    FontSize =8
                    FontWeight =400
                    TabIndex =13
                    ForeColor =8210719
                    Name ="SortLastEvent1"
                    DefaultValue ="0"
                    Caption ="Last Title Event"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =6945
                    LayoutCachedTop =795
                    LayoutCachedWidth =9045
                    LayoutCachedHeight =1065
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    GroupTable =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =16275
                    Top =1080
                    Width =1110
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label188"
                    Caption ="Assignee"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =16275
                    LayoutCachedTop =1080
                    LayoutCachedWidth =17385
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =17955
                    Top =1080
                    Width =1245
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label190"
                    Caption ="Date"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17955
                    LayoutCachedTop =1080
                    LayoutCachedWidth =19200
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =16275
                    Top =795
                    Width =1110
                    Height =270
                    ColumnOrder =4
                    FontSize =8
                    FontWeight =400
                    TabIndex =18
                    ForeColor =8210719
                    Name ="SortAssignee1"
                    DefaultValue ="0"
                    Caption ="Assignee"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =16275
                    LayoutCachedTop =795
                    LayoutCachedWidth =17385
                    LayoutCachedHeight =1065
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    GroupTable =1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =17955
                    Top =795
                    Width =1245
                    Height =270
                    ColumnOrder =5
                    FontSize =8
                    FontWeight =400
                    TabIndex =19
                    ForeColor =8210719
                    Name ="SortFundDate1"
                    DefaultValue ="0"
                    Caption ="Funded"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =17955
                    LayoutCachedTop =795
                    LayoutCachedWidth =19200
                    LayoutCachedHeight =1065
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    GroupTable =1
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =2
                            Left =17385
                            Top =795
                            Width =570
                            Height =270
                            FontSize =8
                            ForeColor =8210719
                            Name ="Label259"
                            Caption ="Fund"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =15
                            LayoutCachedLeft =17385
                            LayoutCachedTop =795
                            LayoutCachedWidth =17955
                            LayoutCachedHeight =1065
                            ColumnStart =12
                            ColumnEnd =12
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =255
                    Width =19200
                    Height =540
                    BorderColor =14211288
                    Name ="Box49"
                    GridlineColor =0
                    HorizontalAnchor =2
                    LayoutCachedWidth =19200
                    LayoutCachedHeight =540
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =255
                    Left =120
                    Top =60
                    Width =5685
                    Height =570
                    FontSize =20
                    BorderColor =0
                    ForeColor =0
                    Name ="Label31"
                    Caption ="Title Inventory /Custodian"
                    FontName ="Segoe UI"
                    GridlineColor =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =5805
                    LayoutCachedHeight =630
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    OverlapFlags =247
                    Left =2610
                    Top =30
                    Width =2115
                    Height =240
                    FontSize =9
                    FontWeight =700
                    BorderColor =0
                    ForeColor =11830108
                    Name ="Label294"
                    Caption ="Operations - Syndication"
                    FontName ="Segoe UI"
                    GridlineColor =0
                    LayoutCachedLeft =2610
                    LayoutCachedTop =30
                    LayoutCachedWidth =4725
                    LayoutCachedHeight =270
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =6840
                    Top =331
                    Width =750
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
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

                    LayoutCachedLeft =6840
                    LayoutCachedTop =331
                    LayoutCachedWidth =7590
                    LayoutCachedHeight =586
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
                Begin CommandButton
                    OverlapFlags =247
                    Left =12735
                    Top =120
                    Width =1710
                    Height =270
                    FontSize =8
                    TabIndex =2
                    Name ="ExcelLegal"
                    Caption ="  Export"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =12735
                    LayoutCachedTop =120
                    LayoutCachedWidth =14445
                    LayoutCachedHeight =390
                    PictureCaptionArrangement =5
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =2
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13277810
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =2
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =17385
                    Top =1080
                    Width =570
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label260"
                    Caption ="Status"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =17385
                    LayoutCachedTop =1080
                    LayoutCachedWidth =17955
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9985
                    Top =180
                    Width =689
                    Height =300
                    ColumnOrder =3
                    FontSize =12
                    FontWeight =700
                    TabIndex =5
                    BorderColor =14277081
                    ForeColor =0
                    Name ="UnitsShown"
                    ControlSource ="=Count([AssetID])"
                    GridlineColor =0

                    LayoutCachedLeft =9985
                    LayoutCachedTop =180
                    LayoutCachedWidth =10674
                    LayoutCachedHeight =480
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =255
                            TextAlign =2
                            Left =9960
                            Width =750
                            Height =540
                            FontSize =7
                            BackColor =14211288
                            BorderColor =14211288
                            ForeColor =0
                            Name ="Label63"
                            Caption ="Units"
                            FontName ="Segoe UI"
                            GridlineColor =0
                            LayoutCachedLeft =9960
                            LayoutCachedWidth =10710
                            LayoutCachedHeight =540
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =7920
                    Top =120
                    Width =1710
                    Height =270
                    FontSize =8
                    TabIndex =1
                    Name ="TitlesOutReport"
                    Caption =" Titles Out Report"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000048772209487622364774216646732196 ,
                        0x457221c3447020f0446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff00000000000000000000000000000000 ,
                        0x00000000000000004a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f9467321904a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ff497822ffadc29cffbccdaeff477421ff517b2eff ,
                        0xedf2eaffccd8c3ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff8fac78ffffffffff6a8e4bff9bb387ff ,
                        0xffffffff6c8e4fff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff4a7925ffe3eaddffc0d0b3ffe9efe5ff ,
                        0xc4d2b8ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffdfdfcfffdfdfcff ,
                        0x648946ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff82a268fffffffffffbfcfaff ,
                        0x5f8640ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ffdae4d3ffd2ddc8ffe5ebe0ff ,
                        0xbbccaeff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff7a9c5effffffffff88a56fff92ad7cff ,
                        0xfdfdfcff658947ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff91ad7affd4dfcbff4b7726ff4d7829ff ,
                        0xe0e8daffc4d2b8ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ffffffffff467321ff467321ff ,
                        0xffffffff467321ff4a7923ff497822ff487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ffffffffffffffffffffffffffffffffff ,
                        0xfbfcfaff467321f94a7923ae497822db487722ff487622ff477421ff467321ff ,
                        0x457221ff447020ff446f20ff436e20ff467321ff467321ff467321ff467321ff ,
                        0x467321f946732190000000000000000048772209487622334774216646732193 ,
                        0x457221c0447020ed446f20ff436e20ff00000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =7920
                    LayoutCachedTop =120
                    LayoutCachedWidth =9630
                    LayoutCachedHeight =390
                    PictureCaptionArrangement =5
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =2
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =13277810
                    HoverTint =80.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeThemeColorIndex =2
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =2
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =14280
                    Top =795
                    Width =1995
                    Height =270
                    ColumnOrder =2
                    FontSize =8
                    FontWeight =400
                    TabIndex =17
                    ForeColor =8210719
                    Name ="SortVIN1"
                    DefaultValue ="0"
                    Caption ="Unit VIN"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =14280
                    LayoutCachedTop =795
                    LayoutCachedWidth =16275
                    LayoutCachedHeight =1065
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    GroupTable =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =9735
                    Top =1080
                    Width =510
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label295"
                    Caption ="Status"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9735
                    LayoutCachedTop =1080
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2448
                    Left =6945
                    Top =1080
                    Width =2100
                    Height =255
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =23
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchEvent"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="2016;418"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6945
                    LayoutCachedTop =1080
                    LayoutCachedWidth =9045
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =6225
                            Top =1080
                            Width =720
                            Height =255
                            FontSize =8
                            ForeColor =8210719
                            Name ="Label270"
                            Caption ="Elapsed"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =6225
                            LayoutCachedTop =1080
                            LayoutCachedWidth =6945
                            LayoutCachedHeight =1335
                            RowStart =1
                            RowEnd =1
                            ColumnStart =5
                            ColumnEnd =5
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =9735
                    Top =795
                    Width =510
                    Height =270
                    ColumnOrder =0
                    FontSize =8
                    FontWeight =400
                    TabIndex =15
                    ForeColor =8210719
                    Name ="SortUnitStatus1"
                    DefaultValue ="0"
                    Caption ="Unit"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =9735
                    LayoutCachedTop =795
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =1065
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    GroupTable =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =255
                    PictureType =1
                    Left =17685
                    Top =165
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =4
                    ForeColor =0
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =17685
                    LayoutCachedTop =165
                    LayoutCachedWidth =18435
                    LayoutCachedHeight =393
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =9
                    HoverTint =60.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =17130
                    Top =165
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =3
                    ForeColor =0
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

                    LayoutCachedLeft =17130
                    LayoutCachedTop =165
                    LayoutCachedWidth =17685
                    LayoutCachedHeight =393
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =2235
                    Top =1080
                    Width =705
                    Height =255
                    Name ="EmptyCell309"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2235
                    LayoutCachedTop =1080
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ToggleButton
                    OverlapFlags =87
                    Left =2235
                    Top =795
                    Width =705
                    Height =270
                    FontSize =8
                    FontWeight =400
                    TabIndex =8
                    ForeColor =8210719
                    Name ="SortBDEX1"
                    DefaultValue ="0"
                    Caption ="BDEX"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =2235
                    LayoutCachedTop =795
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =1065
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15527148
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderWidth =1
                    BorderColor =10855845
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
                    PressedForeThemeColorIndex =-1
                    GroupTable =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =252
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2940
                    Width =1350
                    Height =223
                    FontSize =8
                    TabIndex =6
                    Name ="SchNo"
                    ControlSource ="Schedule"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2940
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6225
                    Width =720
                    Height =223
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="DaysElapsed"
                    ControlSource ="DaysElapsed"
                    Format ="Standard"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6225
                    LayoutCachedWidth =6945
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Top =223
                    Width =18720
                    Height =29
                    FontSize =10
                    TabIndex =3
                    BorderColor =12632256
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0041007300730065007400490044005d003d005b0074007800740043007500 ,
                        0x7200720065006e0074005d0000000000
                    End
                    GridlineColor =0

                    LayoutCachedTop =223
                    LayoutCachedWidth =18720
                    LayoutCachedHeight =252
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00160000005b00 ,
                        0x41007300730065007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d00000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5220
                    Width =1005
                    Height =223
                    FontSize =8
                    TabIndex =8
                    Name ="EarliestDelvDate"
                    ControlSource ="TitleLog"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5220
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4290
                    Width =930
                    Height =223
                    FontSize =8
                    TabIndex =7
                    Name ="UnitGroup"
                    ControlSource ="UnitGroup"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4290
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9045
                    Width =690
                    Height =223
                    FontSize =8
                    TabIndex =11
                    Name ="AssetID"
                    ControlSource ="AssetID"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9045
                    LayoutCachedWidth =9735
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    ScrollBars =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Width =2175
                    Height =223
                    FontSize =8
                    TabIndex =4
                    Name ="Client"
                    ControlSource ="Client"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =60
                    LayoutCachedWidth =2235
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10245
                    Width =4035
                    Height =223
                    FontSize =8
                    TabIndex =13
                    Name ="UnitDesc"
                    ControlSource ="UnitDesc"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10245
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14280
                    Width =1995
                    Height =223
                    FontSize =8
                    TabIndex =14
                    Name ="VIN"
                    ControlSource ="VIN"
                    Format ="True/False"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14280
                    LayoutCachedWidth =16275
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6945
                    Width =2100
                    Height =223
                    FontSize =8
                    TabIndex =10
                    Name ="LastEvent"
                    ControlSource ="LastEvent"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6945
                    LayoutCachedWidth =9045
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Left =18501
                    Width =96
                    FontSize =8
                    TabIndex =1
                    Name ="schid"
                    ControlSource ="SchID"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18501
                    LayoutCachedWidth =18597
                    LayoutCachedHeight =240
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    ScrollBars =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =16275
                    Width =1110
                    Height =223
                    FontSize =8
                    TabIndex =15
                    Name ="Assignee"
                    ControlSource ="Assignee"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =16275
                    LayoutCachedWidth =17385
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    ScrollBars =2
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =3
                    IMESentenceMode =3
                    Left =17955
                    Width =1245
                    Height =223
                    FontSize =8
                    TabIndex =17
                    Name ="FundedDate"
                    ControlSource ="FundedDate"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =17955
                    LayoutCachedWidth =19200
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =18405
                    Width =96
                    ColumnOrder =15
                    FontSize =8
                    Name ="ClientID"
                    ControlSource ="clientID"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18405
                    LayoutCachedWidth =18501
                    LayoutCachedHeight =240
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =18597
                    Width =96
                    FontSize =8
                    TabIndex =2
                    Name ="SGrpID"
                    ControlSource ="SGrpID"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18597
                    LayoutCachedWidth =18693
                    LayoutCachedHeight =240
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =17385
                    Width =570
                    Height =223
                    FontSize =8
                    FontWeight =700
                    TabIndex =16
                    Name ="Status"
                    ControlSource ="Status"
                    Format ="True/False"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =17385
                    LayoutCachedWidth =17955
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9735
                    Width =510
                    Height =223
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    Name ="Text289"
                    ControlSource ="UnitStatus"
                    Format ="True/False"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9735
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2235
                    Width =705
                    Height =223
                    FontSize =8
                    TabIndex =5
                    Name ="Text306"
                    ControlSource ="BDEX"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2235
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =223
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
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
' See "TitleToCustodianDashboard.cls"
