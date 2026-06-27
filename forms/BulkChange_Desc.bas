Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =22579
    DatasheetFontHeight =11
    ItemSuffix =538
    Left =10875
    Top =5925
    Right =-32236
    Bottom =18630
    ShortcutMenuBar ="SortClip"
    OrderBy ="[Bulk_Update_SelectedUnits].[LState] DESC"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0x9431b42e1528e640
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Bulk_Update_SelectedUnits"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1155
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =10170
                    Top =15
                    Width =4035
                    Height =285
                    FontSize =8
                    Name ="Label407"
                    Caption ="┌──────────────────  Linked ──────────────────┐"
                    LayoutCachedLeft =10170
                    LayoutCachedTop =15
                    LayoutCachedWidth =14205
                    LayoutCachedHeight =300
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    Left =7110
                    Top =15
                    Width =1125
                    Height =285
                    FontSize =8
                    Name ="Label126"
                    Caption ="┌── Linked ──┐"
                    LayoutCachedLeft =7110
                    LayoutCachedTop =15
                    LayoutCachedWidth =8235
                    LayoutCachedHeight =300
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Top =960
                    Width =22579
                    Height =195
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =960
                    LayoutCachedWidth =22579
                    LayoutCachedHeight =1155
                    TabIndex =25
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =30
                    Top =825
                    Width =1140
                    Height =255
                    FontSize =9
                    Name ="Label73"
                    Caption ="♦AssetID"
                    RightPadding =38
                    LayoutCachedLeft =30
                    LayoutCachedTop =825
                    LayoutCachedWidth =1170
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    Left =2400
                    Top =825
                    Width =2085
                    Height =255
                    FontSize =9
                    Name ="Label74"
                    Caption ="VIN"
                    LayoutCachedLeft =2400
                    LayoutCachedTop =825
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =1185
                    Top =825
                    Width =1155
                    Height =255
                    FontSize =9
                    Name ="Label75"
                    Caption ="Unit #"
                    RightPadding =38
                    LayoutCachedLeft =1185
                    LayoutCachedTop =825
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =5175
                    Top =825
                    Width =750
                    Height =255
                    FontSize =9
                    Name ="Label77"
                    Caption ="Year"
                    LayoutCachedLeft =5175
                    LayoutCachedTop =825
                    LayoutCachedWidth =5925
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =5970
                    Top =825
                    Width =1575
                    Height =255
                    FontSize =9
                    Name ="Label78"
                    Caption ="Make"
                    LayoutCachedLeft =5970
                    LayoutCachedTop =825
                    LayoutCachedWidth =7545
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =7575
                    Top =825
                    Width =1815
                    Height =255
                    FontSize =9
                    Name ="Label79"
                    Caption ="Model"
                    LayoutCachedLeft =7575
                    LayoutCachedTop =825
                    LayoutCachedWidth =9390
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =9450
                    Top =825
                    Width =1440
                    Height =255
                    FontSize =9
                    Name ="Label80"
                    Caption ="Type"
                    LayoutCachedLeft =9450
                    LayoutCachedTop =825
                    LayoutCachedWidth =10890
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =14970
                    Top =825
                    Width =690
                    Height =255
                    FontSize =9
                    Name ="Label81"
                    Caption ="Axle"
                    LayoutCachedLeft =14970
                    LayoutCachedTop =825
                    LayoutCachedWidth =15660
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    Left =5985
                    Top =510
                    Width =1545
                    Height =255
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =12
                    BackColor =15858167
                    Name ="MakeCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitMakes.UnitMake, LookUp_UnitMakes.MakeID FROM LookUp_UnitMakes "
                        "GROUP BY LookUp_UnitMakes.UnitMake, LookUp_UnitMakes.MakeID ORDER BY LookUp_Unit"
                        "Makes.UnitMake;"
                    ColumnWidths ="1440;0"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Medit"

                    LayoutCachedLeft =5985
                    LayoutCachedTop =510
                    LayoutCachedWidth =7530
                    LayoutCachedHeight =765
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =20
                    ListWidth =3240
                    Left =7590
                    Top =510
                    Width =1800
                    Height =255
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =13
                    BackColor =15858167
                    Name ="ModelCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitModels.UnitModel, LookUp_UnitModels.MakeID, LookUp_UnitModels."
                        "UnitMake FROM LookUp_UnitModels WHERE (((LookUp_UnitModels.UnitMake)=[Forms]![Bu"
                        "lkChange_Form]![BulkChange_Desc].[Form]![MakeCbo])) ORDER BY LookUp_UnitModels.U"
                        "nitModel;"
                    ColumnWidths ="2736;0;0"
                    OnGotFocus ="[Event Procedure]"
                    Tag ="Medit"

                    LayoutCachedLeft =7590
                    LayoutCachedTop =510
                    LayoutCachedWidth =9390
                    LayoutCachedHeight =765
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =9450
                    Top =510
                    Height =255
                    ColumnOrder =11
                    FontSize =9
                    TabIndex =14
                    Name ="TypeCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitTypes.UnitType, LookUp_UnitTypes.UnitTypeDesc AS [Desc], LookU"
                        "p_UnitTypes.UnitTypeID FROM LookUp_UnitTypes ORDER BY LookUp_UnitTypes.UnitTypeI"
                        "D;"
                    ColumnWidths ="1152;1728;0"
                    StatusBarText ="UNIT TYPE"
                    OnClick ="[Event Procedure]"
                    Tag ="Medit"
                    AllowValueListEdits =0

                    LayoutCachedLeft =9450
                    LayoutCachedTop =510
                    LayoutCachedWidth =10890
                    LayoutCachedHeight =765
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5175
                    Top =510
                    Width =750
                    Height =255
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =11
                    Name ="YearCbo"
                    RowSourceType ="Value List"
                    RowSource ="\"2010\";\"2011\";\"2012\";\"2013\";\"2014\";\"2015\";\"2016\";\"2017\";\"2018\""
                        ";\"2019\";\"2020\";\"2021\";\"2022\";\"2023\";\"2024\";\"2025\";\"2026\";\"2027\""
                        ";\"2028\""
                    Tag ="Medit"

                    LayoutCachedLeft =5175
                    LayoutCachedTop =510
                    LayoutCachedWidth =5925
                    LayoutCachedHeight =765
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =9450
                    Top =180
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="TypeSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9450
                    LayoutCachedTop =180
                    LayoutCachedWidth =10890
                    LayoutCachedHeight =420
                    ColumnStart =5
                    ColumnEnd =5
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =7605
                    Top =180
                    Width =1800
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="ModelSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7605
                    LayoutCachedTop =180
                    LayoutCachedWidth =9405
                    LayoutCachedHeight =420
                    ColumnStart =4
                    ColumnEnd =4
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6000
                    Top =180
                    Width =1545
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="MakeSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6000
                    LayoutCachedTop =180
                    LayoutCachedWidth =7545
                    LayoutCachedHeight =420
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5190
                    Top =180
                    Width =750
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="YearSubmit"
                    Caption ="▼"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5190
                    LayoutCachedTop =180
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =420
                    ColumnStart =2
                    ColumnEnd =2
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14970
                    Top =510
                    Width =690
                    Height =255
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =18
                    BackColor =15523798
                    Name ="AxleCbo"
                    RowSourceType ="Value List"
                    RowSource ="\"TBD\";\"S/A\";\"T/A\";\"3/A\";\"4/A\""
                    Tag ="Medit"

                    LayoutCachedLeft =14970
                    LayoutCachedTop =510
                    LayoutCachedWidth =15660
                    LayoutCachedHeight =765
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14970
                    Top =180
                    Width =690
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="AxleSubmit"
                    Caption ="▼"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =14970
                    LayoutCachedTop =180
                    LayoutCachedWidth =15660
                    LayoutCachedHeight =420
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =3420
                    Top =480
                    Width =765
                    Height =210
                    ColumnOrder =6
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =480
                    LayoutCachedWidth =4185
                    LayoutCachedHeight =690
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3480
                    Top =180
                    Width =720
                    Height =300
                    ColumnOrder =7
                    FontWeight =700
                    TabIndex =1
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =3480
                    LayoutCachedTop =180
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =2
                    Left =4545
                    Top =825
                    Width =585
                    Height =255
                    FontSize =9
                    Name ="Label189"
                    Caption ="Status"
                    LayoutCachedLeft =4545
                    LayoutCachedTop =825
                    LayoutCachedWidth =5130
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4515
                    Top =180
                    Width =585
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    Name ="UnitStatusSubmit"
                    Caption ="▼"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4515
                    LayoutCachedTop =180
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =420
                    ColumnStart =14
                    ColumnEnd =14
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1872
                    Left =4515
                    Top =510
                    Width =585
                    Height =255
                    ColumnOrder =13
                    FontSize =9
                    TabIndex =19
                    Name ="UnitStatusCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitStatus.Unitstatus, LookUp_UnitStatus.UnitStatusDesc FROM LookU"
                        "p_UnitStatus ORDER BY LookUp_UnitStatus.Unitstatus;"
                    ColumnWidths ="432;1440"
                    Tag ="Medit"

                    LayoutCachedLeft =4515
                    LayoutCachedTop =510
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =765
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =60
                    Width =3165
                    Height =540
                    FontSize =20
                    FontWeight =700
                    Name ="Label115"
                    Caption ="Description | Spec"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =3225
                    LayoutCachedHeight =540
                    ForeThemeColorIndex =0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10935
                    Top =825
                    Width =1440
                    Height =255
                    FontSize =9
                    Name ="Label381"
                    Caption ="Sub Type"
                    LayoutCachedLeft =10935
                    LayoutCachedTop =825
                    LayoutCachedWidth =12375
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =3168
                    Left =10935
                    Top =510
                    Height =255
                    ColumnOrder =5
                    FontSize =9
                    TabIndex =15
                    Name ="SubTypeCbo"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="1152;1728;0"
                    StatusBarText ="UNIT TYPE"
                    OnGotFocus ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Medit"
                    AllowValueListEdits =0

                    LayoutCachedLeft =10935
                    LayoutCachedTop =510
                    LayoutCachedWidth =12375
                    LayoutCachedHeight =765
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =10935
                    Top =180
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="SubTypeSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10935
                    LayoutCachedTop =180
                    LayoutCachedWidth =12375
                    LayoutCachedHeight =420
                    ColumnStart =6
                    ColumnEnd =6
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =3168
                    Left =12435
                    Top =510
                    Width =1185
                    Height =255
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =16
                    Name ="ConfigCBO"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="1152;1728;0"
                    StatusBarText ="UNIT TYPE"
                    OnGotFocus ="[Event Procedure]"
                    Tag ="Medit"
                    AllowValueListEdits =0

                    LayoutCachedLeft =12435
                    LayoutCachedTop =510
                    LayoutCachedWidth =13620
                    LayoutCachedHeight =765
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =12435
                    Top =180
                    Width =1185
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="UnitConfigSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =12435
                    LayoutCachedTop =180
                    LayoutCachedWidth =13620
                    LayoutCachedHeight =420
                    ColumnStart =6
                    ColumnEnd =6
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =12435
                    Top =825
                    Width =1185
                    Height =255
                    FontSize =9
                    Name ="Label498"
                    Caption ="Config"
                    LayoutCachedLeft =12435
                    LayoutCachedTop =825
                    LayoutCachedWidth =13620
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =13680
                    Top =825
                    Width =1230
                    Height =255
                    FontSize =9
                    Name ="UOMLabel"
                    Caption ="UOM"
                    LayoutCachedLeft =13680
                    LayoutCachedTop =825
                    LayoutCachedWidth =14910
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =13695
                    Top =180
                    Width =1200
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="UnitUOMSubmit"
                    Caption ="▼ Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =13695
                    LayoutCachedTop =180
                    LayoutCachedWidth =14895
                    LayoutCachedHeight =420
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1584
                    Left =13680
                    Top =510
                    Width =1230
                    Height =255
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =17
                    Name ="UOMCbo"
                    RowSourceType ="Value List"
                    ColumnWidths ="0;1"
                    StatusBarText ="Unit UOM"
                    OnGotFocus ="[Event Procedure]"
                    Tag ="Cloned"
                    ControlTipText ="A Daycab Tractor doesn't have a UOM"
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =13680
                    LayoutCachedTop =510
                    LayoutCachedWidth =14910
                    LayoutCachedHeight =765
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =15750
                    Top =825
                    Width =1440
                    Height =255
                    FontSize =9
                    Name ="Label527"
                    Caption ="Axle Property"
                    LayoutCachedLeft =15750
                    LayoutCachedTop =825
                    LayoutCachedWidth =17190
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15750
                    Top =180
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =20
                    Name ="AxlePropertySubmit"
                    Caption ="▼ Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15750
                    LayoutCachedTop =180
                    LayoutCachedWidth =17190
                    LayoutCachedHeight =420
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15750
                    Top =480
                    Height =270
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =21
                    BackColor =15523798
                    Name ="AxlePropertyCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="LookUp_AxleProp"
                    Tag ="Cloned"
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =15750
                    LayoutCachedTop =480
                    LayoutCachedWidth =17190
                    LayoutCachedHeight =750
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =17220
                    Top =825
                    Width =3900
                    Height =255
                    FontSize =9
                    Name ="Label531"
                    Caption ="Unit Descriptor"
                    LayoutCachedLeft =17220
                    LayoutCachedTop =825
                    LayoutCachedWidth =21120
                    LayoutCachedHeight =1080
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =17235
                    Top =180
                    Width =3885
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =23
                    Name ="DescriptorSubmit"
                    Caption ="▼ Submit"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17235
                    LayoutCachedTop =180
                    LayoutCachedWidth =21120
                    LayoutCachedHeight =420
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =17235
                    Top =480
                    Width =3885
                    Height =270
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =22
                    Name ="DescriptorCbo"
                    Tag ="Cloned"
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =17235
                    LayoutCachedTop =480
                    LayoutCachedWidth =21120
                    LayoutCachedHeight =750
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =247
                    Left =3210
                    Top =510
                    Width =1275
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =24
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
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3210
                    LayoutCachedTop =510
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =765
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =6
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =-2147483633
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                    HoverThemeColorIndex =6
                    HoverShade =75.0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =374
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =30
                    Top =30
                    Width =1140
                    Height =255
                    FontSize =9
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    OnDblClick ="[Event Procedure]"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =1170
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2400
                    Top =30
                    Width =2085
                    Height =255
                    FontSize =9
                    TabIndex =12
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    StatusBarText ="VIN number must be unique in table"
                    BottomPadding =38

                    LayoutCachedLeft =2400
                    LayoutCachedTop =30
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =8
                    ColumnEnd =8
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1200
                    Top =30
                    Width =1140
                    Height =255
                    FontSize =9
                    TabIndex =1
                    Name ="UnitUnitNum"
                    ControlSource ="UnitUnitNum"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1200
                    LayoutCachedTop =30
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5175
                    Top =30
                    Width =750
                    Height =255
                    FontSize =9
                    TabIndex =2
                    Name ="UnitYr"
                    ControlSource ="UnitYr"
                    BottomPadding =38

                    LayoutCachedLeft =5175
                    LayoutCachedTop =30
                    LayoutCachedWidth =5925
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5985
                    Top =30
                    Width =1545
                    Height =255
                    FontSize =9
                    TabIndex =3
                    BackColor =15858167
                    Name ="UnitMake"
                    ControlSource ="UnitMake"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitMakes.UnitMake, LookUp_UnitMakes.MakeID FROM LookUp_UnitMakes "
                        "GROUP BY LookUp_UnitMakes.UnitMake, LookUp_UnitMakes.MakeID ORDER BY LookUp_Unit"
                        "Makes.UnitMake;"
                    ColumnWidths ="1440;0"
                    StatusBarText ="unit make"
                    AfterUpdate ="[Event Procedure]"
                    BottomPadding =38
                    AllowValueListEdits =0

                    LayoutCachedLeft =5985
                    LayoutCachedTop =30
                    LayoutCachedWidth =7530
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =3240
                    Left =7590
                    Top =30
                    Width =1800
                    Height =255
                    FontSize =9
                    TabIndex =4
                    BackColor =15858167
                    Name ="UnitModel"
                    ControlSource ="UnitModel"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitModels.UnitModel, LookUp_UnitModels.MakeID, LookUp_UnitModels."
                        "UnitMake FROM LookUp_UnitModels WHERE (((LookUp_UnitModels.UnitMake)=[Forms]![Bu"
                        "lkChange_Form]![BulkChange_Desc].[Form]![UnitMake])) ORDER BY LookUp_UnitModels."
                        "UnitModel;"
                    ColumnWidths ="2736;0;0"
                    StatusBarText ="unit model"
                    BottomPadding =38
                    AllowValueListEdits =0

                    LayoutCachedLeft =7590
                    LayoutCachedTop =30
                    LayoutCachedWidth =9390
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =4
                    ColumnEnd =4
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =3456
                    Left =9450
                    Top =30
                    Height =255
                    FontSize =9
                    TabIndex =5
                    Name ="UnitType"
                    ControlSource ="UnitType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitTypes.UnitType, LookUp_UnitTypes.UnitTypeDesc AS [Desc], LookU"
                        "p_UnitTypes.UnitTypeID FROM LookUp_UnitTypes ORDER BY LookUp_UnitTypes.UnitTypeI"
                        "D;"
                    ColumnWidths ="1728;1728;0"
                    StatusBarText ="UNIT TYPE"
                    OnClick ="[Event Procedure]"
                    BottomPadding =38
                    AllowValueListEdits =0

                    LayoutCachedLeft =9450
                    LayoutCachedTop =30
                    LayoutCachedWidth =10890
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Top =302
                    Width =22579
                    Height =72
                    TabIndex =11
                    ForeColor =16777215
                    Name ="Marker"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0xffffff001f497d00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0055006e0069007400490044005d003d005b00740078007400430075007200 ,
                        0x720065006e0074005d0000000000
                    End
                    HorizontalAnchor =2

                    LayoutCachedTop =302
                    LayoutCachedWidth =22579
                    LayoutCachedHeight =374
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1872
                    Left =4545
                    Top =30
                    Width =585
                    Height =255
                    FontSize =9
                    TabIndex =10
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitStatus.Unitstatus, LookUp_UnitStatus.UnitStatusDesc FROM LookU"
                        "p_UnitStatus ORDER BY LookUp_UnitStatus.Unitstatus;"
                    ColumnWidths ="720;1440"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    BeforeUpdate ="[Event Procedure]"
                    BottomPadding =38

                    LayoutCachedLeft =4545
                    LayoutCachedTop =30
                    LayoutCachedWidth =5130
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =14
                    ColumnEnd =14
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =3456
                    Left =10935
                    Top =30
                    Height =255
                    FontSize =9
                    TabIndex =6
                    Name ="UnitSubType"
                    ControlSource ="UnitSubType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_UnitType_Matrix.UnitSubType, vw_UnitType_Matrix.UnitSubTypeDesc, vw_Un"
                        "itType_Matrix.UnitSubTypeID, vw_UnitType_Matrix.UnitType FROM vw_UnitType_Matrix"
                        " GROUP BY vw_UnitType_Matrix.UnitSubType, vw_UnitType_Matrix.UnitSubTypeDesc, vw"
                        "_UnitType_Matrix.UnitSubTypeID, vw_UnitType_Matrix.UnitType HAVING (((vw_UnitTyp"
                        "e_Matrix.UnitType)=[Forms]![BulkChange_Form]![BulkChange_Desc].[Form]![UnitType]"
                        ")) ORDER BY vw_UnitType_Matrix.UnitSubType;"
                    ColumnWidths ="1728;1728;0"
                    StatusBarText ="UNIT TYPE"
                    OnClick ="[Event Procedure]"
                    BottomPadding =38
                    AllowValueListEdits =0

                    LayoutCachedLeft =10935
                    LayoutCachedTop =30
                    LayoutCachedWidth =12375
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListWidth =3024
                    Left =12435
                    Top =30
                    Width =1185
                    Height =255
                    FontSize =9
                    TabIndex =7
                    Name ="UnitConfiguration"
                    ControlSource ="UnitConfiguration"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_UnitType_Matrix.UnitConfig, vw_UnitType_Matrix.UnitConfigDesc, vw_Unit"
                        "Type_Matrix.UnitConfigID, vw_UnitType_Matrix.UnitSubType, vw_UnitType_Matrix.Uni"
                        "tType FROM vw_UnitType_Matrix GROUP BY vw_UnitType_Matrix.UnitConfig, vw_UnitTyp"
                        "e_Matrix.UnitConfigDesc, vw_UnitType_Matrix.UnitConfigID, vw_UnitType_Matrix.Uni"
                        "tSubType, vw_UnitType_Matrix.UnitType HAVING (((vw_UnitType_Matrix.UnitSubType)="
                        "[Forms]![BulkChange_Form]![BulkChange_Desc].[Form]![UnitSubType]) AND ((vw_UnitT"
                        "ype_Matrix.UnitType)=[Forms]![BulkChange_Form]![BulkChange_Desc].[Form]![UnitTyp"
                        "e]));"
                    ColumnWidths ="2880"
                    StatusBarText ="UNIT TYPE"
                    BottomPadding =38
                    AllowValueListEdits =0

                    LayoutCachedLeft =12435
                    LayoutCachedTop =30
                    LayoutCachedWidth =13620
                    LayoutCachedHeight =285
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1584
                    Left =13680
                    Top =30
                    Width =1230
                    Height =255
                    FontSize =9
                    TabIndex =8
                    Name ="UnitUOM"
                    ControlSource ="UnitUOM"
                    RowSourceType ="Value List"
                    RowSource ="5;5 Car;6;6 Car;7;7 Car;8;8 Car;9;9 Car;10;10 Car"
                    ColumnWidths ="0;1"
                    StatusBarText ="Unit UOM"
                    OnGotFocus ="[Event Procedure]"
                    ControlTipText ="This Unit's UOM, can be inches, feet, gals, cars"
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =13680
                    LayoutCachedTop =30
                    LayoutCachedWidth =14910
                    LayoutCachedHeight =285
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14970
                    Top =30
                    Width =690
                    Height =255
                    FontSize =9
                    TabIndex =9
                    BackColor =15523798
                    Name ="Axle"
                    ControlSource ="Axle"
                    RowSourceType ="Value List"
                    RowSource ="TBD;S/A;T/A;3/A;4/A"

                    LayoutCachedLeft =14970
                    LayoutCachedTop =30
                    LayoutCachedWidth =15660
                    LayoutCachedHeight =285
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15750
                    Top =30
                    Height =270
                    FontSize =9
                    TabIndex =13
                    BackColor =15523798
                    Name ="UnitAxleProperty"
                    ControlSource ="UnitAxleProperty"
                    RowSourceType ="Table/Query"
                    RowSource ="LookUp_AxleProp"
                    Tag ="Cloned"
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =15750
                    LayoutCachedTop =30
                    LayoutCachedWidth =17190
                    LayoutCachedHeight =300
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =17250
                    Top =30
                    Width =3885
                    Height =270
                    FontSize =9
                    TabIndex =14
                    Name ="UnitDescriptor"
                    ControlSource ="UnitDescriptor"
                    Tag ="Cloned"
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =17250
                    LayoutCachedTop =30
                    LayoutCachedWidth =21135
                    LayoutCachedHeight =300
                    ColumnEnd =4
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
' See "BulkChange_Desc.cls"
