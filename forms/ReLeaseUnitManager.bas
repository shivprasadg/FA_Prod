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
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =21240
    DatasheetFontHeight =11
    ItemSuffix =41
    Left =4215
    Top =2460
    Right =25710
    Bottom =17295
    RecSrcDt = Begin
        0x5079b08c37c5e440
    End
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =780
            BackColor =8210719
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    BackStyle =1
                    PictureType =2
                    Top =570
                    Width =21240
                    Height =150
                    BorderColor =6108695
                    Name ="Image142"
                    Picture ="1_DkBlue_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =570
                    LayoutCachedWidth =21240
                    LayoutCachedHeight =720
                    TabIndex =4
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =255
                    Left =960
                    Top =300
                    Width =3660
                    Height =360
                    FontWeight =700
                    Name ="Label7"
                    Caption ="Plus Title History"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =960
                    LayoutCachedTop =300
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =660
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =247
                    Left =960
                    Width =3660
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="Label8"
                    Caption ="Re-Leased Units"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =960
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =360
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Image
                    SizeMode =0
                    PictureType =1
                    Left =120
                    Width =780
                    Height =540
                    Name ="EventImage"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ControlSource ="ImageType"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =900
                    LayoutCachedHeight =540
                    TabIndex =5
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =8040
                    Top =180
                    Width =1975
                    Height =270
                    FontSize =8
                    FontWeight =700
                    ForeColor =0
                    Name ="OpenUnit"
                    Caption ="Go To Unit's Form"
                    OnClick ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8040
                    LayoutCachedTop =180
                    LayoutCachedWidth =10015
                    LayoutCachedHeight =450
                    ColumnStart =3
                    ColumnEnd =3
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =10015
                    Top =180
                    Width =1975
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="OpenSchedule"
                    Caption ="Go To Schedule Control"
                    OnClick ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10015
                    LayoutCachedTop =180
                    LayoutCachedWidth =11990
                    LayoutCachedHeight =450
                    ColumnStart =4
                    ColumnEnd =4
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =19305
                    Top =225
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =2
                    ForeColor =0
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =19305
                    LayoutCachedTop =225
                    LayoutCachedWidth =20055
                    LayoutCachedHeight =453
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
                    OverlapFlags =215
                    Left =18750
                    Top =225
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

                    LayoutCachedLeft =18750
                    LayoutCachedTop =225
                    LayoutCachedWidth =19305
                    LayoutCachedHeight =453
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =14067
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3720
                    Top =795
                    Width =1005
                    FontSize =9
                    TabIndex =4
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="AssetSearch"
                    OnChange ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3720
                    LayoutCachedTop =795
                    LayoutCachedWidth =4725
                    LayoutCachedHeight =1035
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4800
                    Top =795
                    Width =1950
                    FontSize =9
                    TabIndex =8
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="VINSearch"
                    OnChange ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4800
                    LayoutCachedTop =795
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =1035
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15060
                    Top =795
                    Width =1859
                    FontSize =9
                    TabIndex =5
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="TitleNumSearch"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =15060
                    LayoutCachedTop =795
                    LayoutCachedWidth =16919
                    LayoutCachedHeight =1035
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =1200
                    Top =315
                    Width =15120
                    Height =270
                    FontSize =9
                    BorderColor =0
                    Name ="Label88"
                    Caption =" ─────── Search   ───────────────────┬───────────────┬──────────────────────────"
                        "───────────────────────┬────────────────┬───────────────────────────────────────"
                        "─────┐"
                    GridlineColor =0
                    LayoutCachedLeft =1200
                    LayoutCachedTop =315
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =585
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
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =18745
                    Top =405
                    Width =974
                    Height =300
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    BorderColor =14277081
                    ForeColor =0
                    Name ="EventsShown"
                    GridlineColor =0

                    LayoutCachedLeft =18745
                    LayoutCachedTop =405
                    LayoutCachedWidth =19719
                    LayoutCachedHeight =705
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
                            OverlapFlags =93
                            TextAlign =2
                            Left =18720
                            Top =225
                            Width =1035
                            Height =540
                            FontSize =7
                            BackColor =14211288
                            BorderColor =14211288
                            ForeColor =0
                            Name ="Label63"
                            Caption ="Events"
                            FontName ="Segoe UI"
                            GridlineColor =0
                            LayoutCachedLeft =18720
                            LayoutCachedTop =225
                            LayoutCachedWidth =19755
                            LayoutCachedHeight =765
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
                Begin Subform
                    OverlapFlags =85
                    Top =1107
                    Width =21120
                    Height =12960
                    TabIndex =9
                    Name ="ReLeaseUnitsSub"
                    SourceObject ="Form.ReLeaseUnitsSub"

                    LayoutCachedTop =1107
                    LayoutCachedWidth =21120
                    LayoutCachedHeight =14067
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =1
                    Top =555
                    Width =17085
                    Height =240
                    FontSize =8
                    BorderColor =0
                    Name ="RangeLabel"
                    Caption ="            ┌─── Client───┐                                                     "
                        "                           ┌─ AssetID ─┐           ┌──── VIN ────┐              "
                        "                                                                                "
                        "                     ┌─ Assignee ─┐               ┌─ Date Range ─┐              "
                        "                                                                                "
                        " ┌─ Title Num ─┐  "
                    Tag ="Manager"
                    GridlineColor =0
                    LayoutCachedTop =555
                    LayoutCachedWidth =17085
                    LayoutCachedHeight =795
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11880
                    Top =795
                    Width =975
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    BorderColor =4754549
                    ForeColor =0
                    Name ="EndDate"
                    Format ="Short Date"
                    OnChange ="[Event Procedure]"
                    Tag ="Manager"
                    GridlineColor =0

                    LayoutCachedLeft =11880
                    LayoutCachedTop =795
                    LayoutCachedWidth =12855
                    LayoutCachedHeight =1035
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10860
                    Top =802
                    Width =975
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    BorderColor =4754549
                    ForeColor =0
                    Name ="StartDate"
                    Format ="Short Date"
                    OnChange ="[Event Procedure]"
                    Tag ="Manager"
                    GridlineColor =0

                    LayoutCachedLeft =10860
                    LayoutCachedTop =802
                    LayoutCachedWidth =11835
                    LayoutCachedHeight =1042
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6360
                    Top =315
                    Width =1373
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    ForeColor =0
                    Name ="DoSearch"
                    Caption ="Run Search"
                    OnClick ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6360
                    LayoutCachedTop =315
                    LayoutCachedWidth =7733
                    LayoutCachedHeight =585
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    QuickStyle =2
                    QuickStyleMask =-177
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =60
                    Top =795
                    Width =1859
                    Height =255
                    FontSize =9
                    TabIndex =7
                    BackColor =13434879
                    BorderColor =967423
                    ForeColor =0
                    Name ="ClientSearch"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_ReLeasedUnits_LastTitleHistory.ClientId, vw_ReLeasedUnits_LastTitleHis"
                        "tory.Client FROM vw_ReLeasedUnits_LastTitleHistory GROUP BY vw_ReLeasedUnits_Las"
                        "tTitleHistory.ClientId, vw_ReLeasedUnits_LastTitleHistory.Client ORDER BY vw_ReL"
                        "easedUnits_LastTitleHistory.Client;"
                    ColumnWidths ="0;4320"
                    OnChange ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =795
                    LayoutCachedWidth =1919
                    LayoutCachedHeight =1050
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13668
                    Width =996
                    Height =255
                    FontSize =9
                    TabIndex =12
                    BorderColor =-2147483615
                    Name ="Text34"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13668
                    LayoutCachedWidth =14664
                    LayoutCachedHeight =255
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =85.0
                    GroupTable =3
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12672
                    Width =996
                    Height =255
                    FontSize =9
                    TabIndex =13
                    BorderColor =-2147483615
                    Name ="AssetID"
                    GroupTable =3
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12672
                    LayoutCachedWidth =13668
                    LayoutCachedHeight =255
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =85.0
                    GroupTable =3
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =225
                    Top =300
                    Width =1060
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="ClearFilters"
                    Caption ="Clear All  ►"
                    OnClick ="[Event Procedure]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =225
                    LayoutCachedTop =300
                    LayoutCachedWidth =1285
                    LayoutCachedHeight =540
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =75.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =12419407
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =60.0
                    PressedThemeColorIndex =1
                    PressedShade =80.0
                    HoverForeThemeColorIndex =4
                    HoverForeTint =100.0
                    HoverForeShade =75.0
                    PressedForeThemeColorIndex =4
                    PressedForeTint =100.0
                    PressedForeShade =75.0
                    QuickStyle =2
                    QuickStyleMask =-689
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =24
                    Left =9660
                    Top =795
                    Width =1110
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    BorderColor =4754549
                    ForeColor =0
                    Name ="BankSearch"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_SixKeys.BankID, vw_SixKeys.BankShortName FROM vw_SixKeys GROUP BY vw_S"
                        "ixKeys.BankID, vw_SixKeys.BankShortName HAVING (((vw_SixKeys.BankShortName)<>\"T"
                        "BD\")) ORDER BY vw_SixKeys.BankShortName;"
                    ColumnWidths ="0;4320"
                    OnChange ="[Event Procedure]"
                    Tag ="Manager"
                    GridlineColor =0

                    LayoutCachedLeft =9660
                    LayoutCachedTop =795
                    LayoutCachedWidth =10770
                    LayoutCachedHeight =1050
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =8490
                    Top =315
                    Width =1373
                    Height =270
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    ForeColor =0
                    Name ="ExportToExcel"
                    Caption =" Quick Export"
                    OnClick ="[Event Procedure]"
                    Tag ="UAT"
                    ControlTipText ="Export what is displayed to Excel"
                    LeftPadding =90
                    RightPadding =105
                    BottomPadding =165
                    GridlineColor =0
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

                    LayoutCachedLeft =8490
                    LayoutCachedTop =315
                    LayoutCachedWidth =9863
                    LayoutCachedHeight =585
                    PictureCaptionArrangement =5
                    Alignment =1
                    ColumnStart =20
                    ColumnEnd =20
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =6
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedColor =967423
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    Overlaps =1
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
' See "ReLeaseUnitManager.cls"
