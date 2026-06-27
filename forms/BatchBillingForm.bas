Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =28800
    DatasheetFontHeight =11
    ItemSuffix =635
    Left =13455
    Top =2145
    Right =-23026
    Bottom =17550
    RecSrcDt = Begin
        0xcd1daee3942ee640
    End
    RecordSource ="SELECT BillMakerDates.* FROM BillMakerDates;"
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
        Begin CheckBox
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
            Height =720
            BackColor =10040879
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    BackStyle =1
                    PictureType =2
                    Top =540
                    Width =28800
                    Height =180
                    BorderColor =6108695
                    Name ="Image142"
                    Picture ="1_DkBlue_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =540
                    LayoutCachedWidth =28800
                    LayoutCachedHeight =720
                    TabIndex =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =960
                    Top =60
                    Width =3060
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label8"
                    Caption ="Batch Billing Manager"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =960
                    LayoutCachedTop =60
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =480
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Image
                    PictureType =2
                    Left =120
                    Width =780
                    Height =540
                    Name ="EventImage"
                    Picture ="Documents-Ltblue"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ControlSource ="ImageType"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =900
                    LayoutCachedHeight =540
                    TabIndex =4
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin OptionGroup
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =5160
                    Top =60
                    Width =2645
                    Height =360
                    ColumnOrder =0
                    BackColor =15527148
                    BorderColor =0
                    Name ="UnitSelection"
                    DefaultValue ="1"
                    ControlTipText ="Select the type of fleet lists to show."

                    LayoutCachedLeft =5160
                    LayoutCachedTop =60
                    LayoutCachedWidth =7805
                    LayoutCachedHeight =420
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    PictureType =1
                    Left =21345
                    Top =135
                    Width =750
                    Height =228
                    FontSize =8
                    TabIndex =1
                    ForeColor =0
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =21345
                    LayoutCachedTop =135
                    LayoutCachedWidth =22095
                    LayoutCachedHeight =363
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
                    Left =20790
                    Top =135
                    Width =555
                    Height =228
                    FontSize =8
                    TabIndex =2
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

                    LayoutCachedLeft =20790
                    LayoutCachedTop =135
                    LayoutCachedWidth =21345
                    LayoutCachedHeight =363
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
            Height =14700
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =240
                    Width =11415
                    Height =345
                    FontWeight =700
                    ForeColor =0
                    Name ="Label92"
                    Caption ="❶ Set Date Range  ❷ Select 'Desired' Bill Type (or all)  ❸Press 'Find Client Bil"
                        "ls' Button"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =240
                    LayoutCachedWidth =11655
                    LayoutCachedHeight =345
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =23610
                    Top =1140
                    Width =1290
                    Height =300
                    TabIndex =2
                    ForeColor =0
                    Name ="Billfrom2"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =23610
                    LayoutCachedTop =1140
                    LayoutCachedWidth =24900
                    LayoutCachedHeight =1440
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =40.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
                End
                Begin Label
                    Visible = NotDefault
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =12180
                    Top =60
                    Width =1530
                    Height =225
                    FontSize =8
                    FontWeight =700
                    BackColor =15527148
                    BorderColor =0
                    ForeColor =12349952
                    Name ="ClearSeachBoxText2"
                    Caption ="Clear All Seach Boxes"
                    ControlTipText ="Click to clear ALL of the Search Boxes below"
                    GridlineColor =0
                    LayoutCachedLeft =12180
                    LayoutCachedTop =60
                    LayoutCachedWidth =13710
                    LayoutCachedHeight =285
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
                    OverlapFlags =85
                    Left =15300
                    Top =60
                    Height =270
                    FontSize =8
                    FontWeight =700
                    ForeColor =0
                    Name ="ResetForm"
                    Caption ="Reset Form"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Clear seaches and reset the list of units."
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15300
                    LayoutCachedTop =60
                    LayoutCachedWidth =16740
                    LayoutCachedHeight =330
                    ColumnStart =7
                    ColumnEnd =7
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
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =24900
                    Top =1140
                    Width =1290
                    Height =300
                    TabIndex =3
                    ForeColor =0
                    Name ="BillTo2"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =24900
                    LayoutCachedTop =1140
                    LayoutCachedWidth =26190
                    LayoutCachedHeight =1440
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =40.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =300
                    Top =600
                    Width =2505
                    Height =315
                    FontSize =12
                    FontWeight =700
                    BorderColor =0
                    Name ="Label556"
                    Caption ="Client's with Bills"
                    BottomPadding =150
                    LayoutCachedLeft =300
                    LayoutCachedTop =600
                    LayoutCachedWidth =2805
                    LayoutCachedHeight =915
                    RowStart =1
                    RowEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =300
                    Top =915
                    Width =1290
                    Height =225
                    FontSize =10
                    FontWeight =700
                    BorderColor =0
                    ForeColor =0
                    Name ="Label545"
                    Caption ="Bill From"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =300
                    LayoutCachedTop =915
                    LayoutCachedWidth =1590
                    LayoutCachedHeight =1140
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =1590
                    Top =915
                    Width =1290
                    Height =225
                    FontSize =10
                    FontWeight =700
                    BorderColor =0
                    ForeColor =0
                    Name ="Label538"
                    Caption ="Bill To"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1590
                    LayoutCachedTop =915
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =1140
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2880
                    Top =915
                    Width =1890
                    Height =225
                    FontSize =10
                    FontWeight =700
                    BorderColor =0
                    ForeColor =0
                    Name ="Label549"
                    Caption ="Bill Type"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =0
                    LayoutCachedLeft =2880
                    LayoutCachedTop =915
                    LayoutCachedWidth =4770
                    LayoutCachedHeight =1140
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin EmptyCell
                    Left =4785
                    Top =915
                    Width =2175
                    Height =225
                    Name ="EmptyCell544"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =0
                    LayoutCachedLeft =4785
                    LayoutCachedTop =915
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =1140
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GroupTable =4
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =300
                    Top =1140
                    Width =1290
                    Height =285
                    TabIndex =5
                    ForeColor =0
                    Name ="Billfrom"
                    ControlSource ="Billfrom"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =300
                    LayoutCachedTop =1140
                    LayoutCachedWidth =1590
                    LayoutCachedHeight =1425
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =40.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1590
                    Top =1140
                    Width =1290
                    Height =285
                    TabIndex =6
                    ForeColor =0
                    Name ="Billto"
                    ControlSource ="Billto"
                    Format ="Short Date"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =1590
                    LayoutCachedTop =1140
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =1425
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =40.0
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =4
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2880
                    Top =1140
                    Width =1890
                    Height =285
                    TabIndex =7
                    ForeColor =0
                    Name ="BillTypeID"
                    ControlSource ="BillTypeID"
                    RowSourceType ="Table/Query"
                    RowSource ="Select BillTypeID, BillType Lookup_BillType Order By BillType"
                    ColumnWidths ="0;1008"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15

                    LayoutCachedLeft =2880
                    LayoutCachedTop =1140
                    LayoutCachedWidth =4770
                    LayoutCachedHeight =1425
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =40.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =4
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4785
                    Top =1140
                    Width =2175
                    Height =285
                    FontSize =10
                    TabIndex =8
                    Name ="ApplyDateRange"
                    Caption ="Find Client Bills▼"
                    OnClick ="[Event Procedure]"
                    GroupTable =4
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4785
                    LayoutCachedTop =1140
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =1425
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackColor =4754549
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =5026082
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedThemeColorIndex =6
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    GroupTable =4
                    QuickStyle =32
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    Left =11820
                    Top =1140
                    Width =1005
                    Height =255
                    FontWeight =400
                    TabIndex =9
                    ForeColor =8355711
                    Name ="PickModeToggle"
                    DefaultValue ="0"
                    Caption ="▼Pick On"
                    OnClick ="[Event Procedure]"
                    GridlineColor =0

                    LayoutCachedLeft =11820
                    LayoutCachedTop =1140
                    LayoutCachedWidth =12825
                    LayoutCachedHeight =1395
                    ForeTint =50.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Bevel =0
                    BackColor =15060409
                    BackTint =40.0
                    OldBorderStyle =1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =8
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedColor =0
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =12840
                    Top =1140
                    Width =1470
                    Height =255
                    FontSize =10
                    TabIndex =10
                    Name ="GetAll"
                    Caption ="▼All 9"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =12840
                    LayoutCachedTop =1140
                    LayoutCachedWidth =14310
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackColor =4754549
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =5026082
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedThemeColorIndex =6
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    QuickStyle =32
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =14340
                    Top =1140
                    Width =2520
                    Height =255
                    FontSize =10
                    TabIndex =11
                    Name ="GetSelectedBills"
                    Caption ="▼Selected (0) Bills ►"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =0

                    LayoutCachedLeft =14340
                    LayoutCachedTop =1140
                    LayoutCachedWidth =16860
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =0
                    BackColor =8210719
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =10319446
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =10250042
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =7
                    Left =300
                    Top =1455
                    Width =6660
                    Height =4440
                    FontSize =10
                    TabIndex =12
                    ForeColor =0
                    Name ="ClientList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT S.ClientID ,B.BillTypeID, S.ClientShortName AS Client, B.BillType,Count(B"
                        ".BillId) AS Bills, Min(B.BillFrom) AS FirstBill, Max(B.BillFrom) AS LastBill  FR"
                        "OM Bills AS B  INNER JOIN (Select SchId,ClientID, ClientShortName              F"
                        "ROM vw_SixKeys Group By ClientID,ClientShortName,SchId) AS S ON B.BillSchRef = S"
                        ".SchID  GROUP BY S.ClientID, B.BillTypeID, S.ClientShortName, B.BillType, B.Bill"
                        "Type  HAVING Max(BillFrom) >= #6/1/2024#  AND Min(BillFrom) <= #7/31/2024#  AND "
                        "BillTypeID=6  ORDER BY S.ClientShortName, B.BillType;"
                    ColumnWidths ="0;0;1728;1152;720;1152;1152"
                    OnClick ="[Event Procedure]"
                    BottomPadding =150
                    AllowValueListEdits =0

                    LayoutCachedLeft =300
                    LayoutCachedTop =1455
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =5895
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =40.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =7320
                    Top =1455
                    Height =4440
                    FontSize =10
                    TabIndex =13
                    ForeColor =0
                    Name ="AssigneeList"
                    RowSourceType ="Table/Query"
                    RowSource ="Select BankID,Assignee, Count(BankID) as Bills From  (Select BankID,Assignee FRO"
                        "M Bills as B Inner Join (Select ClientID,SchId,Schedule, Assignee,BankID From vw"
                        "_FiveKeys Group By ClientID,SchId,Schedule, Assignee,BankID)  as F  ON F.SchID=B"
                        ".BillSchRef  WHERE BillTypeID=6  AND ClientID=4971  AND BillFrom >= #6/1/2024#  "
                        "AND Billto <= #7/31/2024#  ) AS B Group BY BankID ,Assignee ORDER BY Assignee "
                    ColumnWidths ="0;1008;432"
                    OnClick ="[Event Procedure]"
                    BottomPadding =150
                    AllowValueListEdits =0

                    LayoutCachedLeft =7320
                    LayoutCachedTop =1455
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =5895
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =40.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =7320
                            Top =915
                            Width =1440
                            Height =240
                            FontSize =10
                            BorderColor =0
                            ForeColor =0
                            Name ="Label511"
                            Caption ="Assignee"
                            BottomPadding =150
                            LayoutCachedLeft =7320
                            LayoutCachedTop =915
                            LayoutCachedWidth =8760
                            LayoutCachedHeight =1155
                            RowStart =1
                            RowEnd =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =9180
                    Top =1455
                    Width =2157
                    Height =4440
                    FontSize =10
                    TabIndex =14
                    ForeColor =0
                    Name ="ScheduleList"
                    RowSourceType ="Table/Query"
                    RowSource ="Select SchID, Schedule, Count(SchID) as Bills  FROM Bills as B Inner Join (Selec"
                        "t ClientID,SchId,Schedule, Assignee,BankID From vw_FiveKeys Group By ClientID,Sc"
                        "hId,Schedule, Assignee,BankID)  as F  ON F.SchID=B.BillSchRef  WHERE BillTypeID="
                        "6  AND ClientID=4971  AND BillFrom >= #6/1/2024#  AND Billto <= #7/31/2024#  GRO"
                        "UP BY SchID, Schedule ORDER BY CDbl(NumInString(Left(Schedule,3))) "
                    ColumnWidths ="0;1440;432"
                    OnClick ="[Event Procedure]"
                    BottomPadding =150
                    AllowValueListEdits =0

                    LayoutCachedLeft =9180
                    LayoutCachedTop =1455
                    LayoutCachedWidth =11337
                    LayoutCachedHeight =5895
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =40.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =9180
                            Top =915
                            Width =2145
                            Height =240
                            FontSize =10
                            BorderColor =0
                            ForeColor =0
                            Name ="Label565"
                            Caption ="Schedules"
                            BottomPadding =150
                            LayoutCachedLeft =9180
                            LayoutCachedTop =915
                            LayoutCachedWidth =11325
                            LayoutCachedHeight =1155
                            RowStart =1
                            RowEnd =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =8
                    Left =11835
                    Top =1455
                    Width =5025
                    Height =4440
                    FontSize =10
                    TabIndex =15
                    ForeColor =0
                    Name ="BillList"
                    RowSourceType ="Table/Query"
                    RowSource ="Select BillID,SchID, BillType,Schedule  ,IIf(NOT b.GroupID IS NULL,(Select SGrpN"
                        "O From SchGrp WHERE sgrpId=b.GroupID),'') as IsGrp ,Assignee ,BillFrom, Billto, "
                        "BankID, ClientID, BillTypeID FROM Bills as B Inner Join (Select ClientID,SchId,S"
                        "chedule, Assignee,BankID From vw_FiveKeys Group By ClientID,SchId,Schedule, Assi"
                        "gnee,BankID)  as F  ON F.SchID=B.BillSchRef  WHERE BillTypeID=6  AND ClientID=49"
                        "71  AND BillFrom >= #6/1/2024#  AND Billto <= #7/31/2024#  AND BankID = 4200  OR"
                        "DER BY Assignee, CDbl(NumInString(Left(Schedule,3))), BillFrom "
                    ColumnWidths ="0;0;0;1008;576;1008;1008;1008"
                    OnClick ="[Event Procedure]"
                    BottomPadding =150
                    AllowValueListEdits =0

                    LayoutCachedLeft =11835
                    LayoutCachedTop =1455
                    LayoutCachedWidth =16860
                    LayoutCachedHeight =5895
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =40.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =11835
                            Top =915
                            Width =1800
                            Height =240
                            FontSize =10
                            BorderColor =0
                            ForeColor =0
                            Name ="Label467"
                            Caption ="Bills"
                            BottomPadding =150
                            LayoutCachedLeft =11835
                            LayoutCachedTop =915
                            LayoutCachedWidth =13635
                            LayoutCachedHeight =1155
                            RowStart =1
                            RowEnd =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =16
                    Left =19500
                    Top =1455
                    Width =8580
                    Height =4440
                    FontSize =9
                    TabIndex =16
                    ForeColor =0
                    Name ="MissingBillsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Sch, Grp, bStatus, Age, Assignee as Bank, BLCD, ActiveUnits as Actv, Last"
                        "Bill, BillType as [Bill Type], FirstPDDate as [1st PD], LastOffLease as Last_EOL"
                        ", clientID, SchID, groupID, uGroupID, BillTypeID  FROM vw_BillsStaleMissing_Batc"
                        "h WHERE clientID=4971  ORDER BY CDbl(NumInString(Left(Sch,3))), GRP "
                    ColumnWidths ="576;432;720;432;1008;1008;432;1008;1008;1008;1008;0;0;0;0;0"
                    BottomPadding =150
                    AllowValueListEdits =0

                    LayoutCachedLeft =19500
                    LayoutCachedTop =1455
                    LayoutCachedWidth =28080
                    LayoutCachedHeight =5895
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =40.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =19500
                            Top =540
                            Width =7560
                            Height =315
                            FontSize =12
                            FontWeight =700
                            BorderColor =0
                            Name ="MissingBillsList_Label"
                            Caption ="Transco 'Stale' or 'Missing' Bills"
                            BottomPadding =150
                            LayoutCachedLeft =19500
                            LayoutCachedTop =540
                            LayoutCachedWidth =27060
                            LayoutCachedHeight =855
                            RowStart =1
                            RowEnd =1
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =2
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =17400
                    Top =3750
                    Width =1680
                    Height =270
                    TabIndex =19
                    Name ="CloneBills"
                    Caption ="Clone Bills"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    GroupTable =5
                    LeftPadding =150
                    TopPadding =105
                    RightPadding =165
                    BottomPadding =210
                    GridlineColor =0
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272b700000000000000000000000000000000727272ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffd7d7d7ff ,
                        0x727272ff00000000727272ffffffffffb3b3b3ffffffffffb3b3b3ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff00000000727272ffffffffffffffffff727272ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff00000000727272ffffffffffffffffff727272ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff00000000727272ffffffffffb3b3b3ffffffffffb3b3b3ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff00000000727272ffffffffffffffffff727272ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff00000000727272ffffffffffffffffff727272ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff00000000727272ffffffffffb3b3b3ffffffffffb3b3b3ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff00000000727272ffffffffffffffffff727272ffffffffffffffffff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff ,
                        0x727272ff00000000727272ffffffffffffffffff727272ffffffffffffffffff ,
                        0x90a468ff90a468ff90a468ff90a468ff90a468ff90a468ff90a468ffffffffff ,
                        0x727272ff00000000727272ffffffffffb3b3b3ffffffffffb3b3b3ffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x727272ff00000000727272ffffffffffffffffff727272ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffdbdbdbff ,
                        0x727272ff00000000727272ffffffffffffffffff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272bd00000000727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffff0000000000000000 ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff0000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =17400
                    LayoutCachedTop =3750
                    LayoutCachedWidth =19080
                    LayoutCachedHeight =4020
                    PictureCaptionArrangement =5
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =3
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
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
                    GroupTable =5
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    PictureType =2
                    Left =17400
                    Top =1470
                    Width =1680
                    Height =1380
                    TabIndex =17
                    Name ="MakeBatchInvoice"
                    Caption ="Make Batch Invoice"
                    OnClick ="[Event Procedure]"
                    Picture ="Documents-Ltblue"
                    GroupTable =5
                    LeftPadding =150
                    TopPadding =105
                    RightPadding =165
                    BottomPadding =210
                    GridlineColor =0

                    LayoutCachedLeft =17400
                    LayoutCachedTop =1470
                    LayoutCachedWidth =19080
                    LayoutCachedHeight =2850
                    PictureCaptionArrangement =2
                    LayoutGroup =3
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
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
                    GroupTable =5
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =300
                    Top =6195
                    Width =16680
                    Height =8400
                    TabIndex =20
                    Name ="BatchBillingMassModSub"
                    SourceObject ="Form.BatchBillingMassModSub"

                    LayoutCachedLeft =300
                    LayoutCachedTop =6195
                    LayoutCachedWidth =16980
                    LayoutCachedHeight =14595
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =17400
                    Top =3165
                    Width =1680
                    Height =270
                    TabIndex =18
                    Name ="ExcelReport"
                    Caption ="Excel Report"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddd87262ddddddd876262626dddddd6262626262262628 ,
                        0x2626262626fffff662ff6ff2626f62f2268f2f8626fffff6626fff62626f62f2 ,
                        0x2628f82626fffff66268f862626f62f2262fff2626fffff6628f6f82626f62f2 ,
                        0x26ff2ff626fffff662626262626f62f22626262626fffff6d872626262626268 ,
                        0xddddd87626dddddd
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    GroupTable =5
                    LeftPadding =150
                    TopPadding =105
                    RightPadding =165
                    BottomPadding =210
                    GridlineColor =0
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17400
                    LayoutCachedTop =3165
                    LayoutCachedWidth =19080
                    LayoutCachedHeight =3435
                    PictureCaptionArrangement =5
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =3
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
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
                    GroupTable =5
                    Shadow =-1
                    QuickStyle =30
                    QuickStyleMask =-625
                    Overlaps =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =26190
                    Top =1140
                    Width =1890
                    Height =300
                    TabIndex =4
                    ForeColor =0
                    Name ="BillType2"
                    ControlSource ="BillTypeID"
                    RowSourceType ="Table/Query"
                    RowSource ="Select BillTypeID, BillType Lookup_BillType Order By BillType"
                    ColumnWidths ="0;1008"
                    DefaultValue ="0"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =26190
                    LayoutCachedTop =1140
                    LayoutCachedWidth =28080
                    LayoutCachedHeight =1440
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =40.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GroupTable =14
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =23610
                    Top =840
                    Width =1290
                    Height =300
                    FontSize =10
                    FontWeight =700
                    BorderColor =0
                    ForeColor =0
                    Name ="Label609"
                    Caption ="Bill From "
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =23610
                    LayoutCachedTop =840
                    LayoutCachedWidth =24900
                    LayoutCachedHeight =1140
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =4
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =24900
                    Top =840
                    Width =1290
                    Height =300
                    FontSize =10
                    FontWeight =700
                    BorderColor =0
                    ForeColor =0
                    Name ="Label610"
                    Caption ="Bill To "
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =24900
                    LayoutCachedTop =840
                    LayoutCachedWidth =26190
                    LayoutCachedHeight =1140
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =4
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =26190
                    Top =840
                    Width =1890
                    Height =300
                    FontSize =10
                    FontWeight =700
                    BorderColor =0
                    ForeColor =0
                    Name ="Label611"
                    Caption ="Type "
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =26190
                    LayoutCachedTop =840
                    LayoutCachedWidth =28080
                    LayoutCachedHeight =1140
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =4
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =19500
                    Top =1140
                    Width =4110
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="MakeMissingBill"
                    Caption ="▼ Creat Selected Bill  ▼"
                    OnClick ="[Event Procedure]"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =19500
                    LayoutCachedTop =1140
                    LayoutCachedWidth =23610
                    LayoutCachedHeight =1440
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =4
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Shape =0
                    Gradient =0
                    BackThemeColorIndex =9
                    BackTint =100.0
                    BackShade =75.0
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =9
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedThemeColorIndex =9
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    GroupTable =14
                    QuickStyle =32
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =19500
                    Top =840
                    Width =4110
                    Height =300
                    Name ="EmptyCell633"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =19500
                    LayoutCachedTop =840
                    LayoutCachedWidth =23610
                    LayoutCachedHeight =1140
                    LayoutGroup =4
                    GroupTable =14
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
' See "BatchBillingForm.cls"
