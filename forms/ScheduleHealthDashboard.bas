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
    Width =20220
    DatasheetFontHeight =11
    ItemSuffix =203
    Left =4440
    Top =5415
    Right =21705
    Bottom =20610
    RecSrcDt = Begin
        0xed7c6623d33be540
    End
    RecordSource ="Select * FROM ScheduleHealthCheckUnFunded "
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
            BackColor =14211288
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin ComboBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1860
                    Top =1080
                    Height =255
                    ColumnOrder =18
                    FontSize =9
                    TabIndex =16
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchSchNo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ScheduleHealthCheck.Schno FROM ScheduleHealthCheck GROUP BY ScheduleHealt"
                        "hCheck.Schno, ScheduleHealthCheck.Lessee ORDER BY ScheduleHealthCheck.Schno;"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1860
                    LayoutCachedTop =1080
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =1080
                    Width =1800
                    Height =255
                    ColumnOrder =16
                    FontSize =9
                    TabIndex =15
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchLessee"
                    RowSourceType ="Table/Query"
                    RowSource ="Select Lessee From ScheduleHealthCheck Group by Lessee Order By Lessee"
                    ColumnWidths ="1800"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =1080
                    LayoutCachedWidth =1860
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
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =300
                    Top =558
                    Width =19185
                    Height =225
                    FontSize =9
                    BorderColor =0
                    Name ="Label88"
                    Caption ="  ┌──────────────────────────────────────────────────────────────────────────┴──"
                        "────────────────────────────────────────────────────────────────────────────────"
                        "────────────────────────────────────────────────────┐"
                    GridlineColor =0
                    LayoutCachedLeft =300
                    LayoutCachedTop =558
                    LayoutCachedWidth =19485
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
                    TabIndex =1
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
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5130
                    Top =1080
                    Width =1335
                    Height =255
                    ColumnOrder =17
                    FontSize =9
                    TabIndex =17
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchSchOutBack"
                    RowSourceType ="Value List"
                    RowSource ="Back;Out;Pending"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5130
                    LayoutCachedTop =1080
                    LayoutCachedWidth =6465
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
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
                            OverlapFlags =87
                            TextAlign =2
                            Left =3300
                            Top =1080
                            Width =1830
                            Height =255
                            FontSize =8
                            ForeColor =8210719
                            Name ="Label3"
                            Caption ="├ Actual Delvery ┤"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =3300
                            LayoutCachedTop =1080
                            LayoutCachedWidth =5130
                            LayoutCachedHeight =1335
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =60
                    Top =795
                    Width =1800
                    Height =270
                    ColumnOrder =14
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    ForeColor =8210719
                    Name ="SortLessee1"
                    DefaultValue ="0"
                    Caption ="Lessee"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =60
                    LayoutCachedTop =795
                    LayoutCachedWidth =1860
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
                Begin ToggleButton
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =1860
                    Top =795
                    Width =1440
                    Height =270
                    ColumnOrder =13
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
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

                    LayoutCachedLeft =1860
                    LayoutCachedTop =795
                    LayoutCachedWidth =3300
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
                Begin ToggleButton
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =3300
                    Top =795
                    Width =936
                    Height =270
                    ColumnOrder =10
                    FontSize =8
                    FontWeight =400
                    TabIndex =4
                    ForeColor =8210719
                    Name ="SortEarliestDelvDate1"
                    DefaultValue ="0"
                    Caption ="Earliest"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =3300
                    LayoutCachedTop =795
                    LayoutCachedWidth =4236
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
                    Left =4230
                    Top =795
                    Width =900
                    Height =270
                    ColumnOrder =11
                    FontSize =8
                    FontWeight =400
                    TabIndex =5
                    ForeColor =8210719
                    Name ="SortLatestDelvDate1"
                    DefaultValue ="0"
                    Caption ="Latest"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =4230
                    LayoutCachedTop =795
                    LayoutCachedWidth =5130
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
                Begin ToggleButton
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =5130
                    Top =795
                    Width =1335
                    Height =270
                    ColumnOrder =12
                    FontSize =8
                    FontWeight =400
                    TabIndex =6
                    ForeColor =8210719
                    Name ="SortSchOutBack1"
                    DefaultValue ="0"
                    Caption ="Sch Out/In"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =5130
                    LayoutCachedTop =795
                    LayoutCachedWidth =6465
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
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7800
                    Top =1080
                    Width =1470
                    Height =255
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =18
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="SearchRALOutBack"
                    RowSourceType ="Value List"
                    RowSource ="Back;Out;Pending"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7800
                    LayoutCachedTop =1080
                    LayoutCachedWidth =9270
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
                            OverlapFlags =87
                            TextAlign =2
                            Left =6465
                            Top =1080
                            Width =1335
                            Height =255
                            FontSize =8
                            ForeColor =8210719
                            Name ="Label94"
                            Caption ="Received"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =6465
                            LayoutCachedTop =1080
                            LayoutCachedWidth =7800
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
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =7800
                    Top =795
                    Width =1470
                    Height =270
                    ColumnOrder =8
                    FontSize =8
                    FontWeight =400
                    TabIndex =7
                    ForeColor =8210719
                    Name ="SortRALOutBack1"
                    DefaultValue ="0"
                    Caption ="RAL Out/Back"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =7800
                    LayoutCachedTop =795
                    LayoutCachedWidth =9270
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
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =2
                            Left =6465
                            Top =795
                            Width =1335
                            Height =270
                            FontSize =8
                            ForeColor =8210719
                            Name ="Label63"
                            Caption ="Final Economics"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =15
                            LayoutCachedLeft =6465
                            LayoutCachedTop =795
                            LayoutCachedWidth =7800
                            LayoutCachedHeight =1065
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
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =9270
                    Top =1080
                    Width =1005
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label103"
                    Caption ="Pkg"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9270
                    LayoutCachedTop =1080
                    LayoutCachedWidth =10275
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
                    Left =10275
                    Top =1080
                    Width =1155
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label119"
                    Caption ="Package OUT"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10275
                    LayoutCachedTop =1080
                    LayoutCachedWidth =11430
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
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =10275
                    Top =795
                    Width =1155
                    Height =270
                    ColumnOrder =7
                    FontSize =8
                    FontWeight =400
                    TabIndex =8
                    ForeColor =8210719
                    Name ="SortPermPkgOut1"
                    DefaultValue ="0"
                    Caption ="Perm Funding"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =10275
                    LayoutCachedTop =795
                    LayoutCachedWidth =11430
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
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =2
                            Left =9270
                            Top =795
                            Width =1005
                            Height =270
                            FontSize =8
                            ForeColor =8210719
                            Name ="Label101"
                            Caption ="OPS"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =15
                            LayoutCachedLeft =9270
                            LayoutCachedTop =795
                            LayoutCachedWidth =10275
                            LayoutCachedHeight =1065
                            ColumnStart =7
                            ColumnEnd =7
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =11430
                    Top =1080
                    Width =420
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label123"
                    Caption ="Title"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11430
                    LayoutCachedTop =1080
                    LayoutCachedWidth =11850
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
                    Left =11850
                    Top =795
                    Width =1185
                    Height =270
                    ColumnOrder =6
                    FontSize =8
                    FontWeight =400
                    TabIndex =9
                    ForeColor =8210719
                    Name ="SortIntermFundBank1"
                    DefaultValue ="0"
                    Caption ="Interim Funding"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =11850
                    LayoutCachedTop =795
                    LayoutCachedWidth =13035
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
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =2
                            Left =11430
                            Top =795
                            Width =420
                            Height =270
                            FontSize =8
                            ForeColor =8210719
                            Name ="Label130"
                            Caption ="FA"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =15
                            LayoutCachedLeft =11430
                            LayoutCachedTop =795
                            LayoutCachedWidth =11850
                            LayoutCachedHeight =1065
                            ColumnStart =9
                            ColumnEnd =9
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =11850
                    Top =1080
                    Width =1185
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label148"
                    Caption ="Bank"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11850
                    LayoutCachedTop =1080
                    LayoutCachedWidth =13035
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
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =13035
                    Top =795
                    Width =1185
                    Height =270
                    ColumnOrder =5
                    FontSize =8
                    FontWeight =400
                    TabIndex =10
                    ForeColor =8210719
                    Name ="SortIntermFund1"
                    DefaultValue ="0"
                    Caption ="Interim Funding"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =13035
                    LayoutCachedTop =795
                    LayoutCachedWidth =14220
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
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =13035
                    Top =1080
                    Width =1185
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label157"
                    Caption ="Date"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13035
                    LayoutCachedTop =1080
                    LayoutCachedWidth =14220
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
                    Left =14220
                    Top =1080
                    Width =1485
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label165"
                    Caption ="Date"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =14220
                    LayoutCachedTop =1080
                    LayoutCachedWidth =15705
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
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =15705
                    Top =1080
                    Width =1485
                    Height =255
                    FontSize =8
                    ForeColor =8210719
                    Name ="Label176"
                    Caption ="Date"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =15705
                    LayoutCachedTop =1080
                    LayoutCachedWidth =17190
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
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =17190
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
                    LayoutCachedLeft =17190
                    LayoutCachedTop =1080
                    LayoutCachedWidth =18300
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =18300
                    Top =1080
                    Width =1515
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
                    LayoutCachedLeft =18300
                    LayoutCachedTop =1080
                    LayoutCachedWidth =19815
                    LayoutCachedHeight =1335
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin ToggleButton
                    OverlapFlags =95
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =14220
                    Top =795
                    Width =1485
                    Height =270
                    ColumnOrder =1
                    FontSize =8
                    FontWeight =400
                    TabIndex =11
                    ForeColor =8210719
                    Name ="SortBLCD1"
                    DefaultValue ="0"
                    Caption ="BLCD"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =14220
                    LayoutCachedTop =795
                    LayoutCachedWidth =15705
                    LayoutCachedHeight =1065
                    ColumnStart =12
                    ColumnEnd =12
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
                    Left =17190
                    Top =795
                    Width =1110
                    Height =270
                    ColumnOrder =3
                    FontSize =8
                    FontWeight =400
                    TabIndex =13
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

                    LayoutCachedLeft =17190
                    LayoutCachedTop =795
                    LayoutCachedWidth =18300
                    LayoutCachedHeight =1065
                    ColumnStart =14
                    ColumnEnd =14
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
                    Left =15705
                    Top =795
                    Width =1485
                    Height =270
                    ColumnOrder =2
                    FontSize =8
                    FontWeight =400
                    TabIndex =12
                    ForeColor =8210719
                    Name ="SortAssignmentDate1"
                    DefaultValue ="0"
                    Caption ="Assignment"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =15705
                    LayoutCachedTop =795
                    LayoutCachedWidth =17190
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
                End
                Begin ToggleButton
                    OverlapFlags =87
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =18300
                    Top =795
                    Width =1515
                    Height =270
                    ColumnOrder =4
                    FontSize =8
                    FontWeight =400
                    TabIndex =14
                    ForeColor =8210719
                    Name ="SortFundDate1"
                    DefaultValue ="0"
                    Caption ="Perm Fund"
                    OnClick ="[Event Procedure]"
                    Tag ="TSort"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =18300
                    LayoutCachedTop =795
                    LayoutCachedWidth =19815
                    LayoutCachedHeight =1065
                    ColumnStart =15
                    ColumnEnd =15
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
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =255
                    Left =15
                    Width =20160
                    Height =540
                    BorderColor =14211288
                    Name ="Box49"
                    GridlineColor =0
                    HorizontalAnchor =2
                    LayoutCachedLeft =15
                    LayoutCachedWidth =20175
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
                    Caption ="Schedule Dashboard"
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
                    Left =1710
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
                    LayoutCachedLeft =1710
                    LayoutCachedWidth =3825
                    LayoutCachedHeight =240
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
                    OverlapFlags =255
                    PictureType =1
                    Left =18795
                    Top =135
                    Width =750
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =19
                    ForeColor =0
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =18795
                    LayoutCachedTop =135
                    LayoutCachedWidth =19545
                    LayoutCachedHeight =363
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =9592887
                    BorderTint =100.0
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    HoverColor =-2
                    HoverThemeColorIndex =-1
                    HoverTint =80.0
                    PressedColor =-2
                    PressedThemeColorIndex =-1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =4
                    HoverForeTint =100.0
                    HoverForeShade =75.0
                    PressedForeThemeColorIndex =4
                    PressedForeTint =100.0
                    PressedForeShade =75.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =18240
                    Top =135
                    Width =555
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =20
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

                    LayoutCachedLeft =18240
                    LayoutCachedTop =135
                    LayoutCachedWidth =18795
                    LayoutCachedHeight =363
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =14211288
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =9592887
                    BorderTint =100.0
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    HoverColor =-2
                    HoverThemeColorIndex =-1
                    HoverTint =80.0
                    PressedColor =-2
                    PressedThemeColorIndex =-1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =4
                    HoverForeTint =100.0
                    HoverForeShade =75.0
                    PressedForeThemeColorIndex =4
                    PressedForeTint =100.0
                    PressedForeShade =75.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =255
                    Left =12735
                    Top =120
                    Width =1710
                    Height =270
                    FontSize =8
                    TabIndex =21
                    Name ="ExcelLegal"
                    Caption ="  Export Legal Size"
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
                Begin CommandButton
                    OverlapFlags =247
                    Left =14565
                    Top =120
                    Width =1845
                    Height =270
                    FontSize =8
                    TabIndex =22
                    Name ="Excel11"
                    Caption ="  Export 11x17 Size"
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

                    LayoutCachedLeft =14565
                    LayoutCachedTop =120
                    LayoutCachedWidth =16410
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
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    Left =8505
                    Top =90
                    Width =1995
                    Height =390
                    ColumnOrder =0
                    TabIndex =23
                    BackColor =15527148
                    BorderColor =0
                    Name ="SchSelection"
                    DefaultValue ="3"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8505
                    LayoutCachedTop =90
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =480
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    Begin
                        Begin ToggleButton
                            OverlapFlags =255
                            TextFontCharSet =177
                            TextFontFamily =0
                            Left =8505
                            Top =150
                            Width =455
                            Height =255
                            FontSize =8
                            OptionValue =1
                            Name ="CheckAll"
                            Caption ="ALL"
                            ControlTipText ="All Schedules"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120
                            GridlineColor =0

                            LayoutCachedLeft =8505
                            LayoutCachedTop =150
                            LayoutCachedWidth =8960
                            LayoutCachedHeight =405
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =95.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Shape =0
                            Bevel =0
                            BackColor =12419407
                            BackTint =100.0
                            OldBorderStyle =1
                            ThemeFontIndex =-1
                            PressedThemeColorIndex =2
                            PressedShade =100.0
                            HoverForeThemeColorIndex =2
                            HoverForeTint =100.0
                            HoverForeShade =75.0
                            QuickStyle =30
                            QuickStyleMask =-629
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =255
                            TextFontCharSet =177
                            TextFontFamily =0
                            Left =8955
                            Top =150
                            Width =605
                            Height =255
                            FontSize =8
                            OptionValue =2
                            Name ="CheckLeased"
                            Caption ="Funded"
                            ControlTipText ="All Funded"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120
                            GridlineColor =0

                            LayoutCachedLeft =8955
                            LayoutCachedTop =150
                            LayoutCachedWidth =9560
                            LayoutCachedHeight =405
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =95.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Shape =0
                            Bevel =0
                            BackColor =12419407
                            BackTint =100.0
                            OldBorderStyle =1
                            ThemeFontIndex =-1
                            PressedThemeColorIndex =2
                            PressedShade =100.0
                            HoverForeThemeColorIndex =2
                            HoverForeTint =100.0
                            HoverForeShade =75.0
                            QuickStyle =30
                            QuickStyleMask =-629
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =247
                            TextFontCharSet =177
                            TextFontFamily =0
                            Left =9555
                            Top =150
                            Width =945
                            Height =255
                            FontSize =8
                            OptionValue =3
                            Name ="CheckRemarketing"
                            Caption ="UN-Funded"
                            ControlTipText ="Un-Funded Schedules"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120
                            GridlineColor =0

                            LayoutCachedLeft =9555
                            LayoutCachedTop =150
                            LayoutCachedWidth =10500
                            LayoutCachedHeight =405
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =95.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Shape =0
                            Bevel =0
                            BackColor =12419407
                            BackTint =100.0
                            OldBorderStyle =1
                            ThemeFontIndex =-1
                            PressedThemeColorIndex =2
                            PressedShade =100.0
                            HoverForeThemeColorIndex =2
                            HoverForeTint =100.0
                            HoverForeShade =75.0
                            QuickStyle =30
                            QuickStyleMask =-629
                            Overlaps =1
                        End
                    End
                End
            End
        End
        Begin Section
            Height =614
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1860
                    Height =598
                    FontSize =8
                    TabIndex =2
                    Name ="SchNo"
                    ControlSource ="Schedule"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1860
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =598
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Width =1800
                    Height =598
                    FontSize =8
                    TabIndex =1
                    Name ="Lessee"
                    ControlSource ="Lessee"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =598
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =127
                    IMESentenceMode =3
                    Top =585
                    Width =20160
                    Height =29
                    FontSize =10
                    TabIndex =16
                    BorderColor =12632256
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008c000000010000000100000000000000000000001500000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00730063006800690064005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d0000000000
                    End
                    GridlineColor =0

                    LayoutCachedTop =585
                    LayoutCachedWidth =20160
                    LayoutCachedHeight =614
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00140000005b00 ,
                        0x730063006800690064005d003d005b0074007800740043007500720072006500 ,
                        0x6e0074005d00000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3300
                    Width =1830
                    Height =598
                    FontSize =8
                    TabIndex =3
                    Name ="EarliestDelvDate"
                    ControlSource ="Delvery"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3300
                    LayoutCachedWidth =5130
                    LayoutCachedHeight =598
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5130
                    Width =1335
                    Height =598
                    FontSize =8
                    TabIndex =4
                    Name ="SchOutBack"
                    ControlSource ="SchOutBack"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5130
                    LayoutCachedWidth =6465
                    LayoutCachedHeight =598
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
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6465
                    Width =1335
                    Height =598
                    FontSize =8
                    TabIndex =5
                    Name ="FinalEconRecv"
                    ControlSource ="FinalEconRecv"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6465
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =598
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    ScrollBars =2
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7800
                    Width =1470
                    Height =598
                    FontSize =8
                    TabIndex =6
                    Name ="RALOutBack"
                    ControlSource ="RALBack"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7800
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =598
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9270
                    Width =1005
                    Height =598
                    FontSize =8
                    TabIndex =7
                    Name ="OPSPkg"
                    ControlSource ="OPSPkg"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9270
                    LayoutCachedWidth =10275
                    LayoutCachedHeight =598
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10275
                    Width =1155
                    Height =598
                    FontSize =8
                    TabIndex =8
                    Name ="PermPkgOut"
                    ControlSource ="PermPkgOut"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10275
                    LayoutCachedWidth =11430
                    LayoutCachedHeight =598
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
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11430
                    Width =420
                    Height =598
                    FontSize =8
                    TabIndex =9
                    Name ="FATitleTrust"
                    ControlSource ="FATitleTrust"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11430
                    LayoutCachedWidth =11850
                    LayoutCachedHeight =598
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
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11850
                    Width =1185
                    Height =598
                    FontSize =8
                    TabIndex =10
                    Name ="IntermFundBank"
                    ControlSource ="IntermFundBank"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11850
                    LayoutCachedWidth =13035
                    LayoutCachedHeight =598
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
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13035
                    Width =1185
                    Height =598
                    FontSize =8
                    TabIndex =11
                    Name ="IntermFund"
                    ControlSource ="IntermFund"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13035
                    LayoutCachedWidth =14220
                    LayoutCachedHeight =598
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =19830
                    Top =225
                    Width =246
                    FontSize =8
                    Name ="schid"
                    ControlSource ="SchID"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19830
                    LayoutCachedTop =225
                    LayoutCachedWidth =20076
                    LayoutCachedHeight =465
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    ScrollBars =2
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14220
                    Width =1485
                    Height =598
                    FontSize =8
                    TabIndex =12
                    Name ="BLCD"
                    ControlSource ="BLCD"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =14220
                    LayoutCachedWidth =15705
                    LayoutCachedHeight =598
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    ScrollBars =2
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15705
                    Width =1485
                    Height =598
                    FontSize =8
                    TabIndex =13
                    Name ="AssignmentDate"
                    ControlSource ="Assigned"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15705
                    LayoutCachedWidth =17190
                    LayoutCachedHeight =598
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    ScrollBars =2
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =17190
                    Width =1110
                    Height =598
                    FontSize =8
                    TabIndex =14
                    Name ="Assignee"
                    ControlSource ="Assignee"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =17190
                    LayoutCachedWidth =18300
                    LayoutCachedHeight =598
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    ScrollBars =2
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =18300
                    Width =1515
                    Height =598
                    FontSize =8
                    TabIndex =15
                    Name ="FundDate"
                    ControlSource ="FundDate"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =18300
                    LayoutCachedWidth =19815
                    LayoutCachedHeight =598
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
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
                    Left =19830
                    Top =15
                    Width =246
                    ColumnOrder =15
                    FontSize =8
                    TabIndex =17
                    Name ="ClientID"
                    ControlSource ="clientID"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19830
                    LayoutCachedTop =15
                    LayoutCachedWidth =20076
                    LayoutCachedHeight =255
                    ColumnStart =1
                    ColumnEnd =1
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
' See "ScheduleHealthDashboard.cls"
