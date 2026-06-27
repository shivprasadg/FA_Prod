Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =2
    GridX =24
    GridY =24
    Width =21345
    DatasheetFontHeight =11
    ItemSuffix =5056
    Left =4680
    Top =5190
    Right =22395
    Bottom =18420
    Filter ="UnitID=11575"
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x7a7e66bd7325e540
    End
    RecordSource ="UnitsForm"
    Caption ="ATLAAS Core"
    OnCurrent ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
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
        Begin NavigationControl
            BorderWidth =1
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin NavigationButton
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            HoverThemeColorIndex =2
            HoverTint =20.0
            PressedThemeColorIndex =2
            PressedTint =60.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            BackThemeColorIndex =1
            OldBorderStyle =0
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
            FontName ="Calibri"
            FontWeight =400
            FontSize =11
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =2168
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Top =435
                    Width =21312
                    Height =210
                    Name ="Image582"
                    Picture ="LtGrey2_Shadow"
                    HorizontalAnchor =2

                    LayoutCachedTop =435
                    LayoutCachedWidth =21312
                    LayoutCachedHeight =645
                    TabIndex =25
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2730
                    Top =120
                    Width =14250
                    Height =420
                    ColumnOrder =14
                    FontSize =14
                    Name ="UnitSouceInfo"
                    ControlSource ="=\"<B>\" & UCase([clientshnm]) & \"</b> Lease # <b>\" & [mlno] & \" </b>Sch.<b> "
                        "\" & [schno] & \" </b>Grp.<b>  \" & [sgrpno] & \"</b>\""
                    TextFormat =1

                    LayoutCachedLeft =2730
                    LayoutCachedTop =120
                    LayoutCachedWidth =16980
                    LayoutCachedHeight =540
                    BackThemeColorIndex =1
                    BackShade =85.0
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =247
                    Left =90
                    Width =2190
                    Height =480
                    FontSize =20
                    FontWeight =700
                    Name ="Label522"
                    Caption ="Units Form"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =90
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =480
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =19635
                    Top =150
                    Width =555
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
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

                    LayoutCachedLeft =19635
                    LayoutCachedTop =150
                    LayoutCachedWidth =20190
                    LayoutCachedHeight =378
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderThemeColorIndex =0
                    HoverColor =0
                    PressedColor =-2
                    PressedShade =80.0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    PictureType =1
                    Left =20235
                    Top =150
                    Width =750
                    Height =228
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="CloseMe"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    TopPadding =15
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =20235
                    LayoutCachedTop =150
                    LayoutCachedWidth =20985
                    LayoutCachedHeight =378
                    PictureCaptionArrangement =5
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderThemeColorIndex =0
                    HoverColor =0
                    PressedColor =-2
                    PressedShade =80.0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =660
                    Top =1530
                    Width =720
                    Height =270
                    ColumnOrder =13
                    FontWeight =700
                    TabIndex =11
                    BackColor =11271167
                    Name ="UnitYr"
                    ControlSource ="UnitYr"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Model Year is Required"
                    Tag ="Requried"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =660
                    LayoutCachedTop =1530
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =1800
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =1530
                            Width =569
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label1"
                            Caption ="Year"
                            GroupTable =113
                            LayoutCachedLeft =60
                            LayoutCachedTop =1530
                            LayoutCachedWidth =629
                            LayoutCachedHeight =1800
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1860
                    Top =1530
                    Width =2130
                    Height =270
                    ColumnOrder =15
                    FontWeight =700
                    TabIndex =12
                    BackColor =11271167
                    Name ="UnitMake"
                    ControlSource ="UnitMake"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitMakes.UnitMake, LookUp_UnitMakes.MakeID FROM LookUp_UnitMakes "
                        "GROUP BY LookUp_UnitMakes.UnitMake, LookUp_UnitMakes.MakeID ORDER BY LookUp_Unit"
                        "Makes.UnitMake;"
                    ColumnWidths ="1440;0"
                    StatusBarText ="unit make"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Unit make is required"
                    OnClick ="[Event Procedure]"
                    Tag ="Requried"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =1860
                    LayoutCachedTop =1530
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =1800
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =2
                            Left =1380
                            Top =1530
                            Width =480
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label2"
                            Caption ="Make"
                            GroupTable =113
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =1380
                            LayoutCachedTop =1530
                            LayoutCachedWidth =1860
                            LayoutCachedHeight =1800
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4770
                    Top =1530
                    Width =2100
                    Height =270
                    ColumnOrder =16
                    FontWeight =700
                    TabIndex =13
                    BackColor =11271167
                    Name ="UnitModel"
                    ControlSource ="UnitModel"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitModels.UnitModel, LookUp_UnitModels.MakeID, LookUp_UnitModels."
                        "UnitMake FROM LookUp_UnitModels WHERE (((LookUp_UnitModels.UnitMake)=[forms]![Un"
                        "itsForm]![unitmake])) ORDER BY LookUp_UnitModels.UnitModel;"
                    ColumnWidths ="2880;0"
                    StatusBarText ="unit model"
                    Tag ="Requried"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =4770
                    LayoutCachedTop =1530
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =1800
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =2
                            Left =3990
                            Top =1530
                            Width =780
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label3"
                            Caption ="Model"
                            GroupTable =113
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =3990
                            LayoutCachedTop =1530
                            LayoutCachedWidth =4770
                            LayoutCachedHeight =1800
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =7650
                    Top =1530
                    Width =2100
                    Height =270
                    ColumnWidth =3045
                    ColumnOrder =17
                    FontWeight =700
                    TabIndex =14
                    BackColor =11271167
                    Name ="UnitType"
                    ControlSource ="UnitType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitTypes.UnitType, LookUp_UnitTypes.UnitTypeDesc AS [Desc], LookU"
                        "p_UnitTypes.UnitTypeID FROM LookUp_UnitTypes ORDER BY LookUp_UnitTypes.UnitTypeI"
                        "D;"
                    ColumnWidths ="2160;2160;0"
                    StatusBarText ="UNIT TYPE"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Unit Type is required"
                    OnClick ="[Event Procedure]"
                    Tag ="Requried"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =7650
                    LayoutCachedTop =1530
                    LayoutCachedWidth =9750
                    LayoutCachedHeight =1800
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =2
                            Left =6870
                            Top =1530
                            Width =780
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label4"
                            Caption ="Type"
                            GroupTable =113
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =6870
                            LayoutCachedTop =1530
                            LayoutCachedWidth =7650
                            LayoutCachedHeight =1800
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =5904
                    Left =660
                    Top =1860
                    Width =3330
                    Height =270
                    ColumnOrder =21
                    FontSize =10
                    FontWeight =700
                    TabIndex =20
                    BackColor =11271167
                    Name ="VendorREF"
                    ControlSource ="VendorREF"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName AS Name, Vendors.VendorSNm AS ShortN"
                        "ame, Vendors.VENDORDBA AS DBA FROM Vendors WHERE (((Vendors.VendorObsolete)=No))"
                        " ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;2160;1440;2160"
                    StatusBarText ="lookup from Vendor table"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Vendor is required"
                    OnDblClick ="[Event Procedure]"
                    Tag ="Requried"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =660
                    LayoutCachedTop =1860
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =2130
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =60
                            Top =1860
                            Width =569
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label14"
                            Caption ="Vendor"
                            GroupTable =113
                            LayoutCachedLeft =60
                            LayoutCachedTop =1860
                            LayoutCachedWidth =629
                            LayoutCachedHeight =2130
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4770
                    Top =1860
                    Width =2100
                    Height =270
                    ColumnOrder =22
                    FontWeight =700
                    TabIndex =21
                    BackColor =11271167
                    Name ="UnitFinalCost"
                    ControlSource ="UnitFinalCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Cost of Unit, (not including 'Child Parts')"
                    ValidationRule =">=0"
                    ValidationText ="Unit Cost is required (even zero)"
                    DefaultValue ="0"
                    Tag ="Requried"
                    ControlTipText ="Cost of Unit, (not including 'Child Parts')"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =4770
                    LayoutCachedTop =1860
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =2130
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =2
                            Left =3990
                            Top =1860
                            Width =780
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label15"
                            Caption ="Unit Cost"
                            GroupTable =113
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =3990
                            LayoutCachedTop =1860
                            LayoutCachedWidth =4770
                            LayoutCachedHeight =2130
                            RowStart =1
                            RowEnd =1
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =720
                    Width =1290
                    Height =345
                    ColumnOrder =8
                    FontSize =14
                    TabIndex =3
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    GroupTable =23
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =780
                    LayoutCachedTop =720
                    LayoutCachedWidth =2070
                    LayoutCachedHeight =1065
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =23
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =60
                            Top =720
                            Width =660
                            Height =345
                            FontSize =10
                            FontWeight =700
                            Name ="Label751"
                            Caption ="AssetID"
                            GroupTable =23
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =720
                            LayoutCachedWidth =720
                            LayoutCachedHeight =1065
                            RowEnd =1
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =23
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2730
                    Top =720
                    Width =1995
                    Height =345
                    ColumnOrder =18
                    FontSize =14
                    TabIndex =4
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="UnitUnitNum"
                    ControlSource ="UnitUnitNum"
                    GroupTable =23
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2730
                    LayoutCachedTop =720
                    LayoutCachedWidth =4725
                    LayoutCachedHeight =1065
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =23
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2130
                            Top =720
                            Width =540
                            Height =345
                            FontSize =10
                            FontWeight =700
                            Name ="Label6"
                            Caption ="Unit #"
                            GroupTable =23
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =2130
                            LayoutCachedTop =720
                            LayoutCachedWidth =2670
                            LayoutCachedHeight =1065
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =23
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5595
                    Top =720
                    Width =4125
                    Height =345
                    ColumnWidth =2070
                    ColumnOrder =19
                    FontSize =14
                    TabIndex =5
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="UnitVIN"
                    ControlSource ="UnitVIN"
                    GroupTable =23
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5595
                    LayoutCachedTop =720
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =1065
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    BackShade =95.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =23
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4785
                            Top =720
                            Width =750
                            Height =345
                            FontSize =10
                            FontWeight =700
                            Name ="Label7"
                            Caption ="VIN "
                            GroupTable =23
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =4785
                            LayoutCachedTop =720
                            LayoutCachedWidth =5535
                            LayoutCachedHeight =1065
                            RowEnd =1
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =2
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =23
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =11760
                    Top =770
                    Width =330
                    ColumnOrder =20
                    TabIndex =7
                    Name ="UnitVINLocked"
                    ControlSource ="UnitVINLocked"
                    OnClick ="[Event Procedure]"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =11760
                    LayoutCachedTop =770
                    LayoutCachedWidth =12090
                    LayoutCachedHeight =1010
                    ColumnStart =9
                    ColumnEnd =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2730
                    Top =1080
                    Width =1995
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =10
                    BackColor =14211288
                    BorderColor =14211288
                    Name ="AltUnitnum"
                    ControlSource ="UnitUnitNum2"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="This field is used to match Auto Carriers, the Unit Number inserted would be the"
                        " other half of the pair."
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2730
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4725
                    LayoutCachedHeight =1320
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1710
                            Top =1095
                            Width =960
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="AltLabel"
                            Caption ="Alt Unit #"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1710
                            LayoutCachedTop =1095
                            LayoutCachedWidth =2670
                            LayoutCachedHeight =1335
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =2
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =17745
                    Top =1530
                    Width =1005
                    Height =270
                    ColumnOrder =12
                    FontSize =10
                    FontWeight =700
                    TabIndex =18
                    BackColor =15523798
                    Name ="Axle"
                    ControlSource ="Axle"
                    RowSourceType ="Value List"
                    RowSource ="TBD;S/A;T/A;3/A"
                    Tag ="Cloned"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =17745
                    LayoutCachedTop =1530
                    LayoutCachedWidth =18750
                    LayoutCachedHeight =1800
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =17145
                            Top =1530
                            Width =600
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label460"
                            Caption ="Type"
                            GroupTable =113
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =17145
                            LayoutCachedTop =1530
                            LayoutCachedWidth =17745
                            LayoutCachedHeight =1800
                            ColumnStart =14
                            ColumnEnd =14
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =119
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10530
                    Top =720
                    Width =1230
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =6
                    Name ="UnitVINVer"
                    ControlSource ="UnitVINVer"
                    Format ="Short Date"
                    StatusBarText ="date vin verified"
                    TopPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10530
                    LayoutCachedTop =720
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =960
                    ColumnStart =7
                    ColumnEnd =7
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =2
                            Left =9720
                            Top =720
                            Width =750
                            Height =240
                            FontSize =10
                            FontWeight =700
                            Name ="Label333"
                            Caption ="Verified"
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =9720
                            LayoutCachedTop =720
                            LayoutCachedWidth =10470
                            LayoutCachedHeight =960
                            ColumnStart =6
                            ColumnEnd =6
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =11940
                    Top =720
                    Width =675
                    Height =240
                    FontSize =10
                    FontWeight =700
                    ForeColor =255
                    Name ="Label1806"
                    Caption =" Locked"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11940
                    LayoutCachedTop =720
                    LayoutCachedWidth =12615
                    LayoutCachedHeight =960
                    ColumnStart =8
                    ColumnEnd =8
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10515
                    Top =1860
                    Width =2175
                    Height =270
                    ColumnOrder =7
                    TabIndex =23
                    BackColor =10053120
                    Name ="OEC"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Unit Cost + Child Parts (summed)"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =10515
                    LayoutCachedTop =1860
                    LayoutCachedWidth =12690
                    LayoutCachedHeight =2130
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =10
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =9750
                            Top =1860
                            Width =765
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label2269"
                            Caption ="Total Cost"
                            GroupTable =113
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =9750
                            LayoutCachedTop =1860
                            LayoutCachedWidth =10515
                            LayoutCachedHeight =2130
                            RowStart =1
                            RowEnd =1
                            ColumnStart =8
                            ColumnEnd =8
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =87
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7650
                    Top =1860
                    Width =2100
                    Height =270
                    ColumnOrder =6
                    TabIndex =22
                    Name ="ChildPartCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    ControlTipText ="Child Part Cost (summed)"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =7650
                    LayoutCachedTop =1860
                    LayoutCachedWidth =9750
                    LayoutCachedHeight =2130
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =10
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =1
                            Left =6870
                            Top =1860
                            Width =780
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label3643"
                            Caption ="Child Cost"
                            GroupTable =113
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =6870
                            LayoutCachedTop =1860
                            LayoutCachedWidth =7650
                            LayoutCachedHeight =2130
                            RowStart =1
                            RowEnd =1
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =19020
                    Top =840
                    Width =1995
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="OpenUnitsGroup"
                    Caption ="View This Unit's Group"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =19020
                    LayoutCachedTop =840
                    LayoutCachedWidth =21015
                    LayoutCachedHeight =1095
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =4
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =14211288
                    BorderThemeColorIndex =0
                    HoverColor =0
                    PressedColor =-2
                    PressedShade =80.0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =16305
                    Top =915
                    Width =720
                    Height =495
                    ColumnWidth =1965
                    ColumnOrder =11
                    FontSize =18
                    FontWeight =700
                    TabIndex =9
                    BackColor =11271167
                    Name ="UnitStatus"
                    ControlSource ="UnitStatus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_UnitStatus.Unitstatus, LookUp_UnitStatus.UnitStatusDesc FROM LookU"
                        "p_UnitStatus ORDER BY LookUp_UnitStatus.Unitstatus;"
                    ColumnWidths ="648;3600"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Unit status is required"
                    BeforeUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    Tag ="Requried"
                    ControlTipText ="'The current Status the unit is in, Being Built, Active, Off lease, etc"
                    AllowValueListEdits =0
                    InheritValueList =0

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =16305
                    LayoutCachedTop =915
                    LayoutCachedWidth =17025
                    LayoutCachedHeight =1410
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =16305
                            Top =705
                            Width =720
                            Height =195
                            FontSize =9
                            FontWeight =700
                            BackColor =0
                            Name ="Label436"
                            Caption ="Status"
                            LayoutCachedLeft =16305
                            LayoutCachedTop =705
                            LayoutCachedWidth =17025
                            LayoutCachedHeight =900
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4896
                    Left =10515
                    Top =1530
                    Width =2175
                    Height =270
                    ColumnOrder =5
                    FontWeight =700
                    TabIndex =15
                    BackColor =11271167
                    Name ="UnitSubType"
                    ControlSource ="UnitSubType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_UnitType_Matrix.UnitSubType, vw_UnitType_Matrix.UnitSubTypeDesc, vw_Un"
                        "itType_Matrix.UnitSubTypeID, vw_UnitType_Matrix.UnitType FROM vw_UnitType_Matrix"
                        " GROUP BY vw_UnitType_Matrix.UnitSubType, vw_UnitType_Matrix.UnitSubTypeDesc, vw"
                        "_UnitType_Matrix.UnitSubTypeID, vw_UnitType_Matrix.UnitType HAVING (((vw_UnitTyp"
                        "e_Matrix.UnitType)=[Forms]![UnitsForm]![UnitType]));"
                    ColumnWidths ="2448;2448;0"
                    StatusBarText ="UNIT TYPE"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Unit Type is required"
                    OnClick ="[Event Procedure]"
                    Tag ="Requried"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =10515
                    LayoutCachedTop =1530
                    LayoutCachedWidth =12690
                    LayoutCachedHeight =1800
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =2
                            Left =9750
                            Top =1530
                            Width =765
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label4289"
                            Caption ="Sub Type"
                            GroupTable =113
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =9750
                            LayoutCachedTop =1530
                            LayoutCachedWidth =10515
                            LayoutCachedHeight =1800
                            ColumnStart =8
                            ColumnEnd =8
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4896
                    Left =13230
                    Top =1530
                    Width =1620
                    Height =270
                    ColumnOrder =3
                    FontSize =10
                    FontWeight =700
                    TabIndex =16
                    BackColor =15523798
                    Name ="UnitConfiguration"
                    ControlSource ="UnitConfiguration"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_UnitType_Matrix.UnitConfig, vw_UnitType_Matrix.UnitConfigDesc, vw_Unit"
                        "Type_Matrix.UnitConfigID, vw_UnitType_Matrix.UnitSubType, vw_UnitType_Matrix.Uni"
                        "tType FROM vw_UnitType_Matrix GROUP BY vw_UnitType_Matrix.UnitConfig, vw_UnitTyp"
                        "e_Matrix.UnitConfigDesc, vw_UnitType_Matrix.UnitConfigID, vw_UnitType_Matrix.Uni"
                        "tSubType, vw_UnitType_Matrix.UnitType HAVING (((vw_UnitType_Matrix.UnitSubType)="
                        "[Forms]![UnitsForm]![UnitSubType]) AND ((vw_UnitType_Matrix.UnitType)=[Forms]![U"
                        "nitsForm]![UnitType]));"
                    ColumnWidths ="1728;1728;0"
                    StatusBarText ="UNIT TYPE"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Unit Type is required"
                    Tag ="Requried"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =13230
                    LayoutCachedTop =1530
                    LayoutCachedWidth =14850
                    LayoutCachedHeight =1800
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =12690
                            Top =1530
                            Width =540
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label4397"
                            Caption ="Config"
                            GroupTable =113
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =12690
                            LayoutCachedTop =1530
                            LayoutCachedWidth =13230
                            LayoutCachedHeight =1800
                            ColumnStart =10
                            ColumnEnd =10
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1584
                    Left =15705
                    Top =1530
                    Height =270
                    ColumnOrder =4
                    FontSize =10
                    FontWeight =700
                    TabIndex =17
                    BackColor =15523798
                    Name ="UOMCbo"
                    ControlSource ="UnitUOM"
                    RowSourceType ="Value List"
                    ColumnWidths ="0;1"
                    StatusBarText ="Unit UOM"
                    OnGotFocus ="[Event Procedure]"
                    Tag ="Cloned"
                    ControlTipText ="A Daycab Tractor doesn't have a UOM"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =15705
                    LayoutCachedTop =1530
                    LayoutCachedWidth =17145
                    LayoutCachedHeight =1800
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =14850
                            Top =1530
                            Width =855
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="UOMLabel"
                            Caption ="Measure"
                            GroupTable =113
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =14850
                            LayoutCachedTop =1530
                            LayoutCachedWidth =15705
                            LayoutCachedHeight =1800
                            ColumnStart =12
                            ColumnEnd =12
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
                Begin EmptyCell
                    Left =13230
                    Top =1860
                    Width =1620
                    Height =270
                    Name ="EmptyCell4552"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =13230
                    LayoutCachedTop =1860
                    LayoutCachedWidth =14850
                    LayoutCachedHeight =2130
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =10
                    GroupTable =113
                End
                Begin EmptyCell
                    Left =12690
                    Top =1860
                    Width =540
                    Height =270
                    Name ="EmptyCell4554"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =12690
                    LayoutCachedTop =1860
                    LayoutCachedWidth =13230
                    LayoutCachedHeight =2130
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =10
                    GroupTable =113
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =17145
                    Top =1335
                    Width =600
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label4591"
                    Caption ="Axle"
                    LayoutCachedLeft =17145
                    LayoutCachedTop =1335
                    LayoutCachedWidth =17745
                    LayoutCachedHeight =1575
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =12690
                    Top =1335
                    Width =540
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label4596"
                    Caption ="Unit"
                    LayoutCachedLeft =12690
                    LayoutCachedTop =1335
                    LayoutCachedWidth =13230
                    LayoutCachedHeight =1575
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =14850
                    Top =1335
                    Width =750
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label4597"
                    Caption ="Unit of "
                    LayoutCachedLeft =14850
                    LayoutCachedTop =1335
                    LayoutCachedWidth =15600
                    LayoutCachedHeight =1575
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =18750
                    Top =1320
                    Width =555
                    Height =240
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label4598"
                    Caption ="Axle"
                    LayoutCachedLeft =18750
                    LayoutCachedTop =1320
                    LayoutCachedWidth =19305
                    LayoutCachedHeight =1560
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =19305
                    Top =1530
                    Height =270
                    ColumnOrder =2
                    FontSize =10
                    FontWeight =700
                    TabIndex =19
                    BackColor =15523798
                    Name ="UnitAxleProperty"
                    ControlSource ="UnitAxleProperty"
                    RowSourceType ="Table/Query"
                    RowSource ="LookUp_AxleProp"
                    Tag ="Cloned"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =19305
                    LayoutCachedTop =1530
                    LayoutCachedWidth =20745
                    LayoutCachedHeight =1800
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =18750
                            Top =1530
                            Width =555
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label4404"
                            Caption ="Config"
                            GroupTable =113
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =18750
                            LayoutCachedTop =1530
                            LayoutCachedWidth =19305
                            LayoutCachedHeight =1800
                            ColumnStart =16
                            ColumnEnd =16
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15705
                    Top =1860
                    Width =5040
                    Height =270
                    ColumnOrder =0
                    FontSize =10
                    FontWeight =700
                    TabIndex =24
                    BackColor =15523798
                    Name ="UnitDescriptor"
                    ControlSource ="UnitDescriptor"
                    Tag ="Cloned"
                    GroupTable =113
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =15705
                    LayoutCachedTop =1860
                    LayoutCachedWidth =20745
                    LayoutCachedHeight =2130
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =17
                    LayoutGroup =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =113
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14850
                            Top =1860
                            Width =855
                            Height =270
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label4614"
                            Caption ="Descriptor"
                            GroupTable =113
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =14850
                            LayoutCachedTop =1860
                            LayoutCachedWidth =15705
                            LayoutCachedHeight =2130
                            RowStart =1
                            RowEnd =1
                            ColumnStart =12
                            ColumnEnd =12
                            LayoutGroup =10
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =113
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =11925
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Top =3735
                    Width =18120
                    Height =8190
                    TabIndex =53
                    Name ="UnitsTabControl"

                    LayoutCachedTop =3735
                    LayoutCachedWidth =18120
                    LayoutCachedHeight =11925
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =135
                            Top =4215
                            Width =17850
                            Height =7575
                            BorderColor =0
                            Name ="Locations"
                            Caption ="Locations | Title Log"
                            GridlineColor =0
                            LayoutCachedLeft =135
                            LayoutCachedTop =4215
                            LayoutCachedWidth =17985
                            LayoutCachedHeight =11790
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureAlignment =0
                                    PictureType =2
                                    Left =5325
                                    Top =5415
                                    Width =4020
                                    Height =240
                                    Name ="Image1894"
                                    Picture ="1_WhiteShadow"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =5325
                                    LayoutCachedTop =5415
                                    LayoutCachedWidth =9345
                                    LayoutCachedHeight =5655
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =1
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureAlignment =0
                                    PictureType =2
                                    Left =360
                                    Top =5415
                                    Width =4020
                                    Height =240
                                    Name ="Image1851"
                                    Picture ="1_WhiteShadow"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =360
                                    LayoutCachedTop =5415
                                    LayoutCachedWidth =4380
                                    LayoutCachedHeight =5655
                                    TabIndex =1
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =1
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureAlignment =0
                                    PictureType =2
                                    Left =9690
                                    Top =5355
                                    Width =8219
                                    Height =240
                                    Name ="Image4175"
                                    Picture ="1_WhiteShadow"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =9690
                                    LayoutCachedTop =5355
                                    LayoutCachedWidth =17909
                                    LayoutCachedHeight =5595
                                    TabIndex =2
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =1
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureAlignment =0
                                    PictureType =2
                                    Left =360
                                    Top =8550
                                    Width =4020
                                    Height =240
                                    Name ="Image1896"
                                    Picture ="1_WhiteShadow"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =360
                                    LayoutCachedTop =8550
                                    LayoutCachedWidth =4380
                                    LayoutCachedHeight =8790
                                    TabIndex =3
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =1
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureAlignment =0
                                    PictureType =2
                                    Left =13950
                                    Top =8055
                                    Width =4020
                                    Height =240
                                    Name ="Image4247"
                                    Picture ="1_WhiteShadow"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =13950
                                    LayoutCachedTop =8055
                                    LayoutCachedWidth =17970
                                    LayoutCachedHeight =8295
                                    TabIndex =4
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =1
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureAlignment =0
                                    PictureType =2
                                    Left =5310
                                    Top =8550
                                    Width =4020
                                    Height =240
                                    Name ="Image1982"
                                    Picture ="1_WhiteShadow"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =5310
                                    LayoutCachedTop =8550
                                    LayoutCachedWidth =9330
                                    LayoutCachedHeight =8790
                                    TabIndex =5
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =1
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureAlignment =0
                                    PictureType =2
                                    Left =9690
                                    Top =8535
                                    Width =8234
                                    Height =240
                                    Name ="Image4177"
                                    Picture ="1_WhiteShadow"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =9690
                                    LayoutCachedTop =8535
                                    LayoutCachedWidth =17924
                                    LayoutCachedHeight =8775
                                    TabIndex =6
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =1
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =360
                                    Top =4440
                                    Width =9015
                                    Height =435
                                    FontSize =18
                                    Name ="Label2035"
                                    Caption ="LOCATIONS "
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =360
                                    LayoutCachedTop =4440
                                    LayoutCachedWidth =9375
                                    LayoutCachedHeight =4875
                                    ForeThemeColorIndex =2
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =360
                                    Top =4920
                                    Width =3990
                                    Height =240
                                    FontSize =8
                                    Name ="Label1860"
                                    Caption ="Delivery "
                                    FontName ="Segoe UI"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =360
                                    LayoutCachedTop =4920
                                    LayoutCachedWidth =4350
                                    LayoutCachedHeight =5160
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =9690
                                    Top =5145
                                    Width =8219
                                    Height =285
                                    FontSize =8
                                    BackColor =9732411
                                    Name ="Label4176"
                                    Caption ="Location Change History"
                                    FontName ="Segoe UI"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =9690
                                    LayoutCachedTop =5145
                                    LayoutCachedWidth =17909
                                    LayoutCachedHeight =5430
                                    ColumnEnd =1
                                    BorderThemeColorIndex =0
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin ComboBox
                                    ColumnHeads = NotDefault
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =6
                                    ListWidth =7200
                                    Left =360
                                    Top =5160
                                    Width =3990
                                    Height =300
                                    FontWeight =700
                                    TabIndex =7
                                    BackColor =15523798
                                    Name ="DeliveryCbo"
                                    ControlSource ="UnitDeliveryLocationRef"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LAddrLine1, Locat"
                                        "ions.LCity AS City, Locations.LState AS St, Locations.LZip AS Zip FROM Locations"
                                        " WHERE (((Locations.ClientGroupID)=[FORMS]![UnitsForm]![ClientGroupId]) AND ((Lo"
                                        "cations.LocObsolete)=No)) ORDER BY Locations.LocationName;"
                                    ColumnWidths ="0;2160;2160;1440;720;720"
                                    StatusBarText ="Link to Contacts table"
                                    OnGotFocus ="[Event Procedure]"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =360
                                    LayoutCachedTop =5160
                                    LayoutCachedWidth =4350
                                    LayoutCachedHeight =5460
                                    RowStart =1
                                    RowEnd =1
                                    BorderThemeColorIndex =4
                                    BorderTint =20.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin ComboBox
                                    ColumnHeads = NotDefault
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =6
                                    ListWidth =7200
                                    Left =5325
                                    Top =5160
                                    Width =4020
                                    Height =300
                                    ColumnOrder =9
                                    FontWeight =700
                                    TabIndex =8
                                    BackColor =15523798
                                    Name ="UnitGarageLocationRef"
                                    ControlSource ="UnitGarageLocationRef"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Locations.LocationID, Locations.LocationName, Locations.LAddrLine1, Locat"
                                        "ions.LCity, Locations.LState, Locations.LocObsolete, Locations.LocObsolete, Loca"
                                        "tions.LocObsdate, Locations.LocObsCmt FROM Locations WHERE (((Locations.LocObsol"
                                        "ete)=No) AND ((Locations.ClientGroupID)=[FORMS]![UnitsForm]![CLIENTGROUPID])) OR"
                                        "DER BY Locations.LocationName, Locations.LocationType;"
                                    ColumnWidths ="0;2160;2160;1440;720;720"
                                    StatusBarText ="Link to Contacts table"
                                    ValidationRule ="Is Not Null And >0"
                                    ValidationText ="Domicile location required if unknown  (setup TBD address)"
                                    AfterUpdate ="[Event Procedure]"
                                    OnGotFocus ="[Event Procedure]"
                                    OnDirty ="[Event Procedure]"

                                    LayoutCachedLeft =5325
                                    LayoutCachedTop =5160
                                    LayoutCachedWidth =9345
                                    LayoutCachedHeight =5460
                                    BorderThemeColorIndex =4
                                    BorderTint =20.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =5325
                                            Top =4920
                                            Width =1335
                                            Height =240
                                            FontSize =8
                                            Name ="Label112"
                                            Caption ="Garage "
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =5325
                                            LayoutCachedTop =4920
                                            LayoutCachedWidth =6660
                                            LayoutCachedHeight =5160
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =5
                                    Left =9630
                                    Top =5490
                                    Width =8325
                                    Height =2250
                                    TabIndex =9
                                    Name ="UnitLocationChangeHistoryForm"
                                    SourceObject ="Form.UnitLocationChangeHistoryForm"
                                    LinkChildFields ="AssetId"
                                    LinkMasterFields ="UnitID"

                                    LayoutCachedLeft =9630
                                    LayoutCachedTop =5490
                                    LayoutCachedWidth =17955
                                    LayoutCachedHeight =7740
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =360
                                    Top =5595
                                    Width =4020
                                    Height =930
                                    FontSize =9
                                    TabIndex =10
                                    Name ="DeliveryAddr"
                                    ControlSource ="=\"<B>\" & [DsT1] & \"</b><Br />\" & [Dst2] & \"<Br />\" & IIf(IsNull([DSUB]),Nu"
                                        "ll,[DSUB] & \"<Br />\") & [DCity] & \", \" & [Dst] & \" \" & [Dzip]"
                                    FontName ="Segoe UI"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    TextFormat =1

                                    LayoutCachedLeft =360
                                    LayoutCachedTop =5595
                                    LayoutCachedWidth =4380
                                    LayoutCachedHeight =6525
                                    RowStart =3
                                    RowEnd =3
                                    ColumnEnd =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =5325
                                    Top =5595
                                    Width =4020
                                    Height =930
                                    FontSize =9
                                    TabIndex =11
                                    Name ="Text1893"
                                    ControlSource ="=\"<B>\" & [GsT1] & \"</b><Br />\" & [Gst2] & \"<Br />\" & IIf(IsNull([GSUB]),Nu"
                                        "ll,[GSUB] & \"<Br />\") & [GCity] & \", \" & [Gst] & \" \" & [Gzip] & \" |   <b>"
                                        "County</b> \" & [GCounty]"
                                    FontName ="Segoe UI"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    TextFormat =1

                                    LayoutCachedLeft =5325
                                    LayoutCachedTop =5595
                                    LayoutCachedWidth =9345
                                    LayoutCachedHeight =6525
                                    RowStart =3
                                    RowEnd =3
                                    ColumnEnd =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    ScrollBars =2
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =5300
                                    Top =6570
                                    Width =2415
                                    Height =1065
                                    ColumnOrder =13
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =12
                                    BackColor =15523798
                                    Name ="UnitTaxNotes"
                                    ControlSource ="UnitTaxNotes"
                                    StatusBarText ="8/6/10 trucnate to list box"

                                    LayoutCachedLeft =5300
                                    LayoutCachedTop =6570
                                    LayoutCachedWidth =7715
                                    LayoutCachedHeight =7635
                                    BorderThemeColorIndex =4
                                    BorderTint =20.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            Vertical = NotDefault
                                            BackStyle =1
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =5025
                                            Top =6555
                                            Width =285
                                            Height =1050
                                            FontSize =8
                                            Name ="Label132"
                                            Caption =" └─Tax Notes"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =5025
                                            LayoutCachedTop =6555
                                            LayoutCachedWidth =5310
                                            LayoutCachedHeight =7605
                                            BackThemeColorIndex =1
                                            BorderThemeColorIndex =9
                                            BorderShade =75.0
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =238
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8565
                                    Top =6585
                                    Width =855
                                    FontSize =9
                                    TabIndex =13
                                    Name ="UnitTaxMethod"
                                    ControlSource ="StTaxType"
                                    GroupTable =99

                                    LayoutCachedLeft =8565
                                    LayoutCachedTop =6585
                                    LayoutCachedWidth =9420
                                    LayoutCachedHeight =6825
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =5
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =99
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =238
                                            TextAlign =1
                                            Left =7755
                                            Top =6585
                                            Width =750
                                            Height =240
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label212"
                                            Caption ="Method"
                                            GroupTable =99
                                            LayoutCachedLeft =7755
                                            LayoutCachedTop =6585
                                            LayoutCachedWidth =8505
                                            LayoutCachedHeight =6825
                                            LayoutGroup =5
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =99
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    DecimalPlaces =3
                                    OverlapFlags =215
                                    TextFontCharSet =238
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =8565
                                    Top =6885
                                    Width =855
                                    FontSize =9
                                    TabIndex =14
                                    Name ="UnitTaxRate"
                                    ControlSource ="LTaxRate"
                                    Format ="Percent"
                                    StatusBarText ="8/6/10 being moved to locations"
                                    GroupTable =99

                                    LayoutCachedLeft =8565
                                    LayoutCachedTop =6885
                                    LayoutCachedWidth =9420
                                    LayoutCachedHeight =7125
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =5
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =99
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =238
                                            TextAlign =3
                                            Left =7755
                                            Top =6885
                                            Width =750
                                            Height =240
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label213"
                                            Caption ="Tax Rate "
                                            GroupTable =99
                                            LayoutCachedLeft =7755
                                            LayoutCachedTop =6885
                                            LayoutCachedWidth =8505
                                            LayoutCachedHeight =7125
                                            RowStart =1
                                            RowEnd =1
                                            LayoutGroup =5
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =99
                                        End
                                    End
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =215
                                    Left =8565
                                    Top =7185
                                    Width =855
                                    ColumnOrder =12
                                    TabIndex =15
                                    BorderColor =62207
                                    Name ="UnitTaxExempt"
                                    ControlSource ="UnitTaxExempt"
                                    GroupTable =99

                                    LayoutCachedLeft =8565
                                    LayoutCachedTop =7185
                                    LayoutCachedWidth =9420
                                    LayoutCachedHeight =7425
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =5
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =99
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =215
                                            TextFontCharSet =238
                                            TextAlign =1
                                            Left =7755
                                            Top =7185
                                            Width =750
                                            Height =240
                                            FontSize =9
                                            FontWeight =700
                                            Name ="Label131"
                                            Caption ="Exempt"
                                            GroupTable =99
                                            LayoutCachedLeft =7755
                                            LayoutCachedTop =7185
                                            LayoutCachedWidth =8505
                                            LayoutCachedHeight =7425
                                            RowStart =2
                                            RowEnd =2
                                            LayoutGroup =5
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =99
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =13950
                                    Top =7800
                                    Width =4020
                                    Height =300
                                    FontWeight =700
                                    TabIndex =16
                                    BoundColumn =1
                                    BackColor =15523798
                                    Name ="UCCC"
                                    ControlSource ="UCCC"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT DISTINCT Locations.LocationID, Locations.LocationName FROM Locations WHER"
                                        "E (((Locations.ClientGroupID)=[FORMS]![UnitsForm]![CLIENTgroupID]) AND ((Locatio"
                                        "ns.LocationType)=\"garage\"));"
                                    ColumnWidths ="0;1872"
                                    StatusBarText ="Unit Accounting Code - to Add new code simply type in value"
                                    AfterUpdate ="[Event Procedure]"
                                    ShortcutMenuBar ="Unit Accounting Code - to Add new code simply type in value"
                                    OnDirty ="[Event Procedure]"

                                    LayoutCachedLeft =13950
                                    LayoutCachedTop =7800
                                    LayoutCachedWidth =17970
                                    LayoutCachedHeight =8100
                                    BorderThemeColorIndex =4
                                    BorderTint =20.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =1
                                            Left =11535
                                            Top =7830
                                            Width =2370
                                            Height =255
                                            FontSize =8
                                            Name ="UCCLabel"
                                            Caption ="Unit Location for Accounting"
                                            FontName ="Segoe UI"
                                            ControlTipText ="Unit Cost Accounting Code - to Add new code simply type in value"
                                            LayoutCachedLeft =11535
                                            LayoutCachedTop =7830
                                            LayoutCachedWidth =13905
                                            LayoutCachedHeight =8085
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin ComboBox
                                    ColumnHeads = NotDefault
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =7
                                    ListWidth =7200
                                    Left =360
                                    Top =8295
                                    Width =4020
                                    Height =300
                                    ColumnOrder =10
                                    FontWeight =700
                                    TabIndex =17
                                    BackColor =15523798
                                    Name ="UnitRegLocationRef"
                                    ControlSource ="UnitRegLocationRef"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Locations.LocationID, Locations.Locationsn, Locations.LocationType, Locat"
                                        "ions.LAddrLine1, Locations.LState, Locations.LZip, Locations.LCity FROM Location"
                                        "s WHERE (((Locations.LocationType)<>\"billing\") AND ((Locations.ClientGroupID)="
                                        "[FORMS]![UnitsForm]![CLIENTgroupID]) AND ((Locations.LocObsolete)=No)) ORDER BY "
                                        "Locations.Locationsn, Locations.LocationType, Locations.LState, Locations.LCity,"
                                        " Locations.LocationName;"
                                    ColumnWidths ="0;1440;1440;2160;1440;720;720"
                                    StatusBarText ="Link to Contacts table"
                                    OnGotFocus ="[Event Procedure]"
                                    Tag ="Cloned"

                                    LayoutCachedLeft =360
                                    LayoutCachedTop =8295
                                    LayoutCachedWidth =4380
                                    LayoutCachedHeight =8595
                                    BorderThemeColorIndex =4
                                    BorderTint =20.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =1
                                            Left =360
                                            Top =8025
                                            Width =1725
                                            Height =240
                                            FontSize =8
                                            Name ="Label114"
                                            Caption ="Registration "
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =360
                                            LayoutCachedTop =8025
                                            LayoutCachedWidth =2085
                                            LayoutCachedHeight =8265
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin ComboBox
                                    ColumnHeads = NotDefault
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =6
                                    ListWidth =7200
                                    Left =5310
                                    Top =8295
                                    Width =3300
                                    Height =300
                                    ColumnOrder =15
                                    FontWeight =700
                                    TabIndex =18
                                    BackColor =15523798
                                    Name ="UnitTitleLocationRef"
                                    ControlSource ="UnitTitleLocationRef"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Locations.LocationID, Locations.Locationsn, Locations.LocationType, Locat"
                                        "ions.LAddrLine1, Locations.LCity, Locations.LState, Locations.LZip FROM Location"
                                        "s WHERE (((Locations.LocationType)<>\"Billing\") AND ((Locations.ClientGroupID)="
                                        "[FORMS]![UnitsForm]![CLIENTgroupID]) AND ((Locations.LocObsolete)=No)) ORDER BY "
                                        "Locations.Locationsn, Locations.LocationType, Locations.LCity, Locations.Locatio"
                                        "nName;"
                                    ColumnWidths ="0;1440;1440;2160;1440;720;720"
                                    StatusBarText ="Link to Contacts table"
                                    OnGotFocus ="[Event Procedure]"

                                    LayoutCachedLeft =5310
                                    LayoutCachedTop =8295
                                    LayoutCachedWidth =8610
                                    LayoutCachedHeight =8595
                                    BorderThemeColorIndex =4
                                    BorderTint =20.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =1
                                            Left =5310
                                            Top =8025
                                            Width =1890
                                            Height =240
                                            FontSize =8
                                            Name ="Label113"
                                            Caption ="Location Title Mailed To"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =5310
                                            LayoutCachedTop =8025
                                            LayoutCachedWidth =7200
                                            LayoutCachedHeight =8265
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =8685
                                    Top =8295
                                    Width =660
                                    Height =315
                                    FontWeight =700
                                    TabIndex =19
                                    BackColor =15523798
                                    Name ="TitelST"
                                    ControlSource ="tST"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT LookUp_States.StateAbbrev AS ST FROM LookUp_States;"
                                    StatusBarText ="lookup in lookup_state table"
                                    Format =">"

                                    LayoutCachedLeft =8685
                                    LayoutCachedTop =8295
                                    LayoutCachedWidth =9345
                                    LayoutCachedHeight =8610
                                    BorderThemeColorIndex =4
                                    BorderTint =20.0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =8700
                                            Top =7890
                                            Width =675
                                            Height =405
                                            FontSize =8
                                            Name ="Label445"
                                            Caption ="Title \015\012State"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =8700
                                            LayoutCachedTop =7890
                                            LayoutCachedWidth =9375
                                            LayoutCachedHeight =8295
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =9690
                                    Top =8325
                                    Width =8234
                                    Height =285
                                    FontSize =8
                                    BackColor =9732411
                                    Name ="Label4178"
                                    Caption ="Title Log History"
                                    FontName ="Segoe UI"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    LayoutCachedLeft =9690
                                    LayoutCachedTop =8325
                                    LayoutCachedWidth =17924
                                    LayoutCachedHeight =8610
                                    ColumnEnd =1
                                    BorderThemeColorIndex =0
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =13950
                                    Top =8355
                                    Width =1890
                                    Height =225
                                    FontSize =8
                                    TabIndex =20
                                    Name ="UTlogq"
                                    Caption ="Unit Title Log"
                                    FontName ="Segoe UI"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120
                                    OnClickEmMacro = Begin
                                        Version =196611
                                        ColumnsShown =8
                                        Begin
                                            Action ="OpenReport"
                                            Argument ="UnitTitleLog"
                                            Argument ="2"
                                            Argument =""
                                            Argument =""
                                            Argument ="0"
                                        End
                                        Begin
                                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                                "nterfaceMacro For=\"UTlogq\" xmlns=\"http://schemas.microsoft.com/office/accesss"
                                                "ervices/2009/11/application\"><Statements><Action Name=\"OpenReport\"><Argument "
                                                "Name=\"ReportName\">UnitTitleLog<"
                                        End
                                        Begin
                                            Comment ="_AXL:/Argument><Argument Name=\"View\">Print Preview</Argument></Action></Statem"
                                                "ents></UserInterfaceMacro>"
                                        End
                                    End

                                    LayoutCachedLeft =13950
                                    LayoutCachedTop =8355
                                    LayoutCachedWidth =15840
                                    LayoutCachedHeight =8580
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =1
                                    Gradient =25
                                    BackColor =9732411
                                    BorderThemeColorIndex =8
                                    BorderShade =50.0
                                    HoverThemeColorIndex =7
                                    HoverShade =75.0
                                    PressedThemeColorIndex =8
                                    PressedShade =75.0
                                    PressedForeThemeColorIndex =1
                                    QuickStyle =30
                                    QuickStyleMask =-117
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =15975
                                    Top =8355
                                    Width =1890
                                    Height =225
                                    FontSize =8
                                    TabIndex =21
                                    Name ="AddTitleLogEntry"
                                    Caption ="Update Title Log"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Segoe UI"
                                    ControlTipText ="Open Form"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120

                                    LayoutCachedLeft =15975
                                    LayoutCachedTop =8355
                                    LayoutCachedWidth =17865
                                    LayoutCachedHeight =8580
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =1
                                    Gradient =25
                                    BackColor =9732411
                                    BorderThemeColorIndex =8
                                    BorderShade =50.0
                                    HoverThemeColorIndex =7
                                    HoverShade =75.0
                                    PressedThemeColorIndex =8
                                    PressedShade =75.0
                                    PressedForeThemeColorIndex =1
                                    QuickStyle =30
                                    QuickStyleMask =-117
                                    Overlaps =1
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =5
                                    Left =9660
                                    Top =8655
                                    Width =8325
                                    Height =3135
                                    TabIndex =22
                                    Name ="UNittitleLogForm"
                                    SourceObject ="Form.UNittitleLogForm"
                                    LinkChildFields ="UnitID"
                                    LinkMasterFields ="UnitID"

                                    LayoutCachedLeft =9660
                                    LayoutCachedTop =8655
                                    LayoutCachedWidth =17985
                                    LayoutCachedHeight =11790
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =360
                                    Top =8715
                                    Width =4020
                                    Height =930
                                    FontSize =9
                                    TabIndex =23
                                    Name ="Text1895"
                                    ControlSource ="=\"<B>\" & [Rs1] & \"</b><Br />\" & [Rs2] & \"<Br />\" & IIf(IsNull([RSUB]),Null"
                                        ",[RSUB] & \"<Br />\") & [RCity] & \", \" & [Rst] & \" \" & [Rzip]"
                                    FontName ="Segoe UI"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    TextFormat =1

                                    LayoutCachedLeft =360
                                    LayoutCachedTop =8715
                                    LayoutCachedWidth =4380
                                    LayoutCachedHeight =9645
                                    RowStart =3
                                    RowEnd =3
                                    ColumnEnd =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =1230
                                    Top =9915
                                    Width =1170
                                    Height =270
                                    FontSize =10
                                    TabIndex =24
                                    BackColor =15523798
                                    Name ="UnitIRPAcct"
                                    ControlSource ="UnitIRPAcct"
                                    Format ="mm/dd/yy"
                                    StatusBarText ="IRP Acct Number"
                                    Tag ="Cloned"
                                    GroupTable =88

                                    LayoutCachedLeft =1230
                                    LayoutCachedTop =9915
                                    LayoutCachedWidth =2400
                                    LayoutCachedHeight =10185
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =4
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =88
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =1
                                            Left =390
                                            Top =9915
                                            Width =780
                                            Height =270
                                            FontSize =8
                                            Name ="Label148"
                                            Caption ="IRP Acct"
                                            FontName ="Segoe UI"
                                            GroupTable =88
                                            LayoutCachedLeft =390
                                            LayoutCachedTop =9915
                                            LayoutCachedWidth =1170
                                            LayoutCachedHeight =10185
                                            LayoutGroup =4
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =88
                                        End
                                    End
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =215
                                    Left =2460
                                    Top =9915
                                    Width =285
                                    Height =270
                                    TabIndex =25
                                    BorderColor =12632256
                                    Name ="UnitIRP"
                                    ControlSource ="UnitIRP"
                                    StatusBarText ="yes unit is IRP'ed"
                                    GroupTable =88

                                    LayoutCachedLeft =2460
                                    LayoutCachedTop =9915
                                    LayoutCachedWidth =2745
                                    LayoutCachedHeight =10185
                                    ColumnStart =2
                                    ColumnEnd =2
                                    LayoutGroup =4
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =88
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =1
                                    Left =2805
                                    Top =9915
                                    Width =420
                                    Height =270
                                    FontSize =8
                                    Name ="Label168"
                                    Caption ="IRP"
                                    FontName ="Segoe UI"
                                    GroupTable =88
                                    LayoutCachedLeft =2805
                                    LayoutCachedTop =9915
                                    LayoutCachedWidth =3225
                                    LayoutCachedHeight =10185
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =4
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =88
                                End
                                Begin EmptyCell
                                    Left =3285
                                    Top =9915
                                    Width =1110
                                    Height =270
                                    Name ="EmptyCell4271"
                                    GroupTable =88
                                    LayoutCachedLeft =3285
                                    LayoutCachedTop =9915
                                    LayoutCachedWidth =4395
                                    LayoutCachedHeight =10185
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =4
                                    GroupTable =88
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =1230
                                    Top =10245
                                    Width =1170
                                    Height =255
                                    FontSize =10
                                    TabIndex =26
                                    BackColor =15523798
                                    Name ="Plateno"
                                    ControlSource ="UnitPlateNum"
                                    Tag ="Cloned"
                                    GroupTable =88

                                    LayoutCachedLeft =1230
                                    LayoutCachedTop =10245
                                    LayoutCachedWidth =2400
                                    LayoutCachedHeight =10500
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =4
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =88
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =1
                                            Left =390
                                            Top =10245
                                            Width =780
                                            Height =255
                                            FontSize =8
                                            Name ="Label170"
                                            Caption ="Plate #"
                                            FontName ="Segoe UI"
                                            GroupTable =88
                                            LayoutCachedLeft =390
                                            LayoutCachedTop =10245
                                            LayoutCachedWidth =1170
                                            LayoutCachedHeight =10500
                                            RowStart =1
                                            RowEnd =1
                                            LayoutGroup =4
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =88
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3285
                                    Top =10245
                                    Width =1110
                                    Height =255
                                    FontSize =10
                                    TabIndex =27
                                    BackColor =15523798
                                    Name ="Units.UnitRegExp"
                                    ControlSource ="Units.UnitRegExp"
                                    Format ="Short Date"
                                    StatusBarText ="Registration Expiration"
                                    InputMask ="99/99/00;"
                                    Tag ="Cloned"
                                    EventProcPrefix ="Units_UnitRegExp"
                                    GroupTable =88

                                    LayoutCachedLeft =3285
                                    LayoutCachedTop =10245
                                    LayoutCachedWidth =4395
                                    LayoutCachedHeight =10500
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =4
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =88
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =1
                                            Left =2460
                                            Top =10245
                                            Width =765
                                            Height =255
                                            FontSize =8
                                            Name ="Label217"
                                            Caption ="Reg Exp"
                                            FontName ="Segoe UI"
                                            GroupTable =88
                                            LayoutCachedLeft =2460
                                            LayoutCachedTop =10245
                                            LayoutCachedWidth =3225
                                            LayoutCachedHeight =10500
                                            RowStart =1
                                            RowEnd =1
                                            ColumnStart =2
                                            ColumnEnd =3
                                            LayoutGroup =4
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =88
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1230
                                    Top =10560
                                    Width =1170
                                    ColumnOrder =56
                                    FontSize =10
                                    TabIndex =28
                                    BackColor =15523798
                                    Name ="UnitRegW"
                                    ControlSource ="UnitRegW"
                                    Format ="Standard"
                                    StatusBarText ="weight unit registered at"
                                    ValidationRule =">=0"
                                    ValidationText ="Reg Weight must be >=0"
                                    DefaultValue ="0"
                                    Tag ="Cloned"
                                    GroupTable =88

                                    LayoutCachedLeft =1230
                                    LayoutCachedTop =10560
                                    LayoutCachedWidth =2400
                                    LayoutCachedHeight =10800
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =4
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =88
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =1
                                            Left =390
                                            Top =10560
                                            Width =780
                                            Height =240
                                            FontSize =8
                                            Name ="Label24"
                                            Caption ="Weight"
                                            FontName ="Segoe UI"
                                            GroupTable =88
                                            LayoutCachedLeft =390
                                            LayoutCachedTop =10560
                                            LayoutCachedWidth =1170
                                            LayoutCachedHeight =10800
                                            RowStart =2
                                            RowEnd =2
                                            LayoutGroup =4
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =88
                                        End
                                    End
                                End
                                Begin EmptyCell
                                    Left =2460
                                    Top =10560
                                    Width =765
                                    Name ="EmptyCell4250"
                                    GroupTable =88
                                    LayoutCachedLeft =2460
                                    LayoutCachedTop =10560
                                    LayoutCachedWidth =3225
                                    LayoutCachedHeight =10800
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =2
                                    ColumnEnd =3
                                    LayoutGroup =4
                                    GroupTable =88
                                End
                                Begin EmptyCell
                                    Left =3285
                                    Top =10560
                                    Width =1110
                                    Name ="EmptyCell4254"
                                    GroupTable =88
                                    LayoutCachedLeft =3285
                                    LayoutCachedTop =10560
                                    LayoutCachedWidth =4395
                                    LayoutCachedHeight =10800
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =4
                                    ColumnEnd =4
                                    LayoutGroup =4
                                    GroupTable =88
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1230
                                    Top =10860
                                    Width =1170
                                    FontSize =10
                                    TabIndex =29
                                    BackColor =15523798
                                    Name ="FATitleLog"
                                    ControlSource ="AssetID_FATitleLog"
                                    Format ="General Number"
                                    StatusBarText ="weight unit registered at"
                                    ValidationRule =">=0"
                                    ValidationText ="Reg Weight must be >=0"
                                    DefaultValue ="0"
                                    Tag ="Cloned"
                                    GroupTable =88

                                    LayoutCachedLeft =1230
                                    LayoutCachedTop =10860
                                    LayoutCachedWidth =2400
                                    LayoutCachedHeight =11100
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =4
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =88
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =1
                                            Left =390
                                            Top =10860
                                            Width =780
                                            Height =240
                                            FontSize =8
                                            Name ="Label4732"
                                            Caption ="TitleLog#"
                                            FontName ="Segoe UI"
                                            GroupTable =88
                                            LayoutCachedLeft =390
                                            LayoutCachedTop =10860
                                            LayoutCachedWidth =1170
                                            LayoutCachedHeight =11100
                                            RowStart =3
                                            RowEnd =3
                                            LayoutGroup =4
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =88
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =1
                                    Left =2460
                                    Top =10860
                                    Width =1935
                                    Height =240
                                    FontSize =8
                                    Name ="Label4735"
                                    Caption ="< linked to prev-lease"
                                    FontName ="Segoe UI"
                                    GroupTable =88
                                    LayoutCachedLeft =2460
                                    LayoutCachedTop =10860
                                    LayoutCachedWidth =4395
                                    LayoutCachedHeight =11100
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =2
                                    ColumnEnd =4
                                    LayoutGroup =4
                                    ForeThemeColorIndex =2
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =88
                                End
                                Begin CommandButton
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    Left =3150
                                    Top =11340
                                    Width =1890
                                    Height =285
                                    FontSize =10
                                    TabIndex =30
                                    Name ="LocationChanges"
                                    Caption ="Location Change Detail"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =3150
                                    LayoutCachedTop =11340
                                    LayoutCachedWidth =5040
                                    LayoutCachedHeight =11625
                                    ForeThemeColorIndex =1
                                    UseTheme =1
                                    Gradient =25
                                    BackThemeColorIndex =4
                                    BorderThemeColorIndex =1
                                    HoverThemeColorIndex =4
                                    HoverTint =80.0
                                    PressedThemeColorIndex =4
                                    PressedShade =80.0
                                    HoverForeThemeColorIndex =1
                                    PressedForeThemeColorIndex =1
                                    Shadow =-1
                                    QuickStyle =30
                                    QuickStyleMask =-65
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    Visible = NotDefault
                                    OverlapFlags =215
                                    Left =975
                                    Top =11355
                                    Width =1920
                                    Height =285
                                    FontSize =10
                                    TabIndex =31
                                    Name ="Command467"
                                    Caption ="Acct Cd Change Detail"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120
                                    OnClickEmMacro = Begin
                                        Version =196611
                                        ColumnsShown =8
                                        Begin
                                            Action ="OpenReport"
                                            Argument ="ADMAcctCDChq"
                                            Argument ="2"
                                            Argument =""
                                            Argument =""
                                            Argument ="0"
                                        End
                                        Begin
                                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                                "nterfaceMacro For=\"Command467\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                                "essservices/2009/11/application\"><Statements><Action Name=\"OpenReport\"><Argum"
                                                "ent Name=\"ReportName\">ADMAcctCD"
                                        End
                                        Begin
                                            Comment ="_AXL:Chq</Argument><Argument Name=\"View\">Print Preview</Argument></Action></St"
                                                "atements></UserInterfaceMacro>"
                                        End
                                    End
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =975
                                    LayoutCachedTop =11355
                                    LayoutCachedWidth =2895
                                    LayoutCachedHeight =11640
                                    ForeThemeColorIndex =1
                                    UseTheme =1
                                    Gradient =25
                                    BackThemeColorIndex =4
                                    BorderThemeColorIndex =1
                                    HoverThemeColorIndex =4
                                    HoverTint =80.0
                                    PressedThemeColorIndex =4
                                    PressedShade =80.0
                                    HoverForeThemeColorIndex =1
                                    PressedForeThemeColorIndex =1
                                    Shadow =-1
                                    QuickStyle =30
                                    QuickStyleMask =-65
                                    Overlaps =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =4215
                            Width =17850
                            Height =7575
                            BorderColor =0
                            Name ="InterimFundingNotes"
                            Caption ="Notes - Legal"
                            GridlineColor =0
                            LayoutCachedLeft =135
                            LayoutCachedTop =4215
                            LayoutCachedWidth =17985
                            LayoutCachedHeight =11790
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =255
                                    Top =4470
                                    Width =9525
                                    Height =495
                                    FontSize =20
                                    Name ="Label2155"
                                    Caption ="Unit Notes and Documentation - Legal Desc"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =255
                                    LayoutCachedTop =4470
                                    LayoutCachedWidth =9780
                                    LayoutCachedHeight =4965
                                    BackThemeColorIndex =1
                                    ForeThemeColorIndex =6
                                    ForeShade =75.0
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =180
                                    Top =5085
                                    Width =9660
                                    Height =6615
                                    Name ="AddUnitNtView"
                                    SourceObject ="Form.AddUnitNtView"
                                    LinkChildFields ="UnitRef"
                                    LinkMasterFields ="UnitID"

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =5085
                                    LayoutCachedWidth =9840
                                    LayoutCachedHeight =11700
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =9960
                                    Top =5745
                                    Width =7860
                                    Height =540
                                    FontSize =10
                                    TabIndex =1
                                    BackColor =15523798
                                    Name ="LegalDescription"
                                    ControlSource ="LegalDescription"
                                    Tag ="Cloned"
                                    ControlTipText ="\"Legal/Documentation. The value held here will superceed Atlaas' generated Unit"
                                        " Description\""

                                    LayoutCachedLeft =9960
                                    LayoutCachedTop =5745
                                    LayoutCachedWidth =17820
                                    LayoutCachedHeight =6285
                                    ColumnStart =1
                                    ColumnEnd =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =9960
                                            Top =5505
                                            Width =7860
                                            Height =270
                                            FontSize =10
                                            FontWeight =700
                                            Name ="Label4603"
                                            Caption ="Legal / CofA Unit Description"
                                            LeftPadding =0
                                            RightPadding =0
                                            LayoutCachedLeft =9960
                                            LayoutCachedTop =5505
                                            LayoutCachedWidth =17820
                                            LayoutCachedHeight =5775
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =9960
                                    Top =6525
                                    Width =7860
                                    Height =270
                                    FontSize =10
                                    TabIndex =2
                                    BackColor =15523798
                                    Name ="Text4610"
                                    ControlSource ="=DLookUp(\"DescriptionLong\",\"LookUp_UnitDesc\",\"AssetId=\" & [UnitID])"

                                    LayoutCachedLeft =9960
                                    LayoutCachedTop =6525
                                    LayoutCachedWidth =17820
                                    LayoutCachedHeight =6795
                                    ColumnStart =1
                                    ColumnEnd =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =9960
                                            Top =6825
                                            Width =7860
                                            Height =240
                                            FontSize =9
                                            Name ="Label4611"
                                            Caption ="Unit's Printed Description as entered"
                                            LeftPadding =0
                                            RightPadding =0
                                            LayoutCachedLeft =9960
                                            LayoutCachedTop =6825
                                            LayoutCachedWidth =17820
                                            LayoutCachedHeight =7065
                                            ThemeFontIndex =1
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                                Begin TextBox
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =9960
                                    Top =7485
                                    Width =7860
                                    Height =540
                                    FontSize =10
                                    TabIndex =3
                                    BackColor =14151142
                                    Name ="DocsSUFSubtitle"
                                    ControlSource ="DocsSUFSubtitle"
                                    Tag ="Cloned"
                                    ControlTipText ="\"Legal/Documentation uses this field to match the Unit Description on the CofA,"
                                        " in case it is different that Atlaas' generated Unit Description\""

                                    LayoutCachedLeft =9960
                                    LayoutCachedTop =7485
                                    LayoutCachedWidth =17820
                                    LayoutCachedHeight =8025
                                    ColumnStart =1
                                    ColumnEnd =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =9960
                                            Top =7245
                                            Width =7860
                                            Height =270
                                            FontSize =10
                                            FontWeight =700
                                            Name ="Label4738"
                                            Caption ="SUF Sub Title - This will appear on the SUF as a Sub title (75 Character Limit)"
                                            LeftPadding =0
                                            RightPadding =0
                                            LayoutCachedLeft =9960
                                            LayoutCachedTop =7245
                                            LayoutCachedWidth =17820
                                            LayoutCachedHeight =7515
                                            ForeThemeColorIndex =6
                                            ForeShade =50.0
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =60
                            Top =4215
                            Width =17925
                            Height =7575
                            BorderColor =0
                            Name ="ChildParts"
                            Caption ="Child Parts"
                            GridlineColor =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =4215
                            LayoutCachedWidth =17985
                            LayoutCachedHeight =11790
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =60
                                    Top =4365
                                    Width =17760
                                    Height =6300
                                    Name ="ChildParts_sub"
                                    SourceObject ="Form.ChildParts_sub"
                                    LinkChildFields ="UnitREF"
                                    LinkMasterFields ="UnitID"

                                    LayoutCachedLeft =60
                                    LayoutCachedTop =4365
                                    LayoutCachedWidth =17820
                                    LayoutCachedHeight =10665
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =4215
                            Width =17850
                            Height =7575
                            BorderColor =0
                            Name ="VINDecoded"
                            Caption ="VIN Decoded | SERIALS"
                            GridlineColor =0
                            LayoutCachedLeft =135
                            LayoutCachedTop =4215
                            LayoutCachedWidth =17985
                            LayoutCachedHeight =11790
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureAlignment =0
                                    PictureType =2
                                    Left =11775
                                    Top =4845
                                    Width =3375
                                    Height =240
                                    Name ="Image4125"
                                    Picture ="1_WhiteShadow"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =11775
                                    LayoutCachedTop =4845
                                    LayoutCachedWidth =15150
                                    LayoutCachedHeight =5085
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =1
                                End
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureAlignment =0
                                    PictureType =2
                                    Left =210
                                    Top =4845
                                    Width =11487
                                    Height =240
                                    Name ="Image4124"
                                    Picture ="1_WhiteShadow"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =210
                                    LayoutCachedTop =4845
                                    LayoutCachedWidth =11697
                                    LayoutCachedHeight =5085
                                    TabIndex =1
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =1
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =255
                                    Left =210
                                    Top =4485
                                    Width =11490
                                    Height =435
                                    FontSize =18
                                    Name ="Label2227"
                                    Caption =" VIN Decoded"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =210
                                    LayoutCachedTop =4485
                                    LayoutCachedWidth =11700
                                    LayoutCachedHeight =4920
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    ForeThemeColorIndex =2
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =11775
                                    Top =4485
                                    Width =3375
                                    Height =435
                                    FontSize =18
                                    Name ="Label2255"
                                    Caption ="Component Serial #"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =11775
                                    LayoutCachedTop =4485
                                    LayoutCachedWidth =15150
                                    LayoutCachedHeight =4920
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    ForeThemeColorIndex =2
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =9540
                                    Top =4575
                                    Width =1995
                                    Height =255
                                    FontSize =8
                                    FontWeight =700
                                    TabIndex =8
                                    Name ="DecodeMyVin"
                                    Caption ="Decode This Unit's VIN"
                                    OnClick ="[Event Procedure]"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =9540
                                    LayoutCachedTop =4575
                                    LayoutCachedWidth =11535
                                    LayoutCachedHeight =4830
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =4
                                    ForeThemeColorIndex =0
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    UseTheme =1
                                    BackColor =14211288
                                    BorderThemeColorIndex =0
                                    HoverColor =0
                                    PressedColor =-2
                                    PressedShade =80.0
                                    PressedForeThemeColorIndex =0
                                    QuickStyle =22
                                    QuickStyleMask =-629
                                    Overlaps =1
                                End
                                Begin Subform
                                    CanGrow = NotDefault
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =165
                                    Top =5070
                                    Width =11235
                                    Height =3855
                                    TabIndex =2
                                    Name ="VINDecode"
                                    SourceObject ="Form.VINDecode"
                                    LinkChildFields ="UNITID"
                                    LinkMasterFields ="UnitID"

                                    LayoutCachedLeft =165
                                    LayoutCachedTop =5070
                                    LayoutCachedWidth =11400
                                    LayoutCachedHeight =8925
                                End
                                Begin TextBox
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =13170
                                    Top =5190
                                    Width =1770
                                    Height =285
                                    FontSize =10
                                    TabIndex =3
                                    Name ="Text293"
                                    ControlSource ="UnitEngineSerial"
                                    GroupTable =103

                                    LayoutCachedLeft =13170
                                    LayoutCachedTop =5190
                                    LayoutCachedWidth =14940
                                    LayoutCachedHeight =5475
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =6
                                    BackThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =103
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =11910
                                            Top =5190
                                            Width =1200
                                            Height =285
                                            FontSize =8
                                            FontWeight =700
                                            Name ="Label294"
                                            Caption ="ENGINE "
                                            GroupTable =103
                                            LayoutCachedLeft =11910
                                            LayoutCachedTop =5190
                                            LayoutCachedWidth =13110
                                            LayoutCachedHeight =5475
                                            LayoutGroup =6
                                            ThemeFontIndex =1
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =103
                                        End
                                    End
                                End
                                Begin TextBox
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =13170
                                    Top =5535
                                    Width =1770
                                    Height =285
                                    FontSize =10
                                    TabIndex =4
                                    Name ="Text295"
                                    ControlSource ="UnitTransSerial"
                                    GroupTable =103

                                    LayoutCachedLeft =13170
                                    LayoutCachedTop =5535
                                    LayoutCachedWidth =14940
                                    LayoutCachedHeight =5820
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =6
                                    BackThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =103
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =11910
                                            Top =5535
                                            Width =1200
                                            Height =285
                                            FontSize =8
                                            FontWeight =700
                                            Name ="Label296"
                                            Caption ="TRANSMISSION"
                                            GroupTable =103
                                            LayoutCachedLeft =11910
                                            LayoutCachedTop =5535
                                            LayoutCachedWidth =13110
                                            LayoutCachedHeight =5820
                                            RowStart =1
                                            RowEnd =1
                                            LayoutGroup =6
                                            ThemeFontIndex =1
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =103
                                        End
                                    End
                                End
                                Begin TextBox
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =13170
                                    Top =5880
                                    Width =1770
                                    Height =255
                                    FontSize =10
                                    TabIndex =5
                                    Name ="Text297"
                                    ControlSource ="UnitFASerial"
                                    StatusBarText ="front axle serial"
                                    GroupTable =103

                                    LayoutCachedLeft =13170
                                    LayoutCachedTop =5880
                                    LayoutCachedWidth =14940
                                    LayoutCachedHeight =6135
                                    RowStart =2
                                    RowEnd =2
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =6
                                    BackThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =103
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =11910
                                            Top =5880
                                            Width =1200
                                            Height =255
                                            FontSize =8
                                            FontWeight =700
                                            Name ="Label298"
                                            Caption ="FRONT AXLE "
                                            GroupTable =103
                                            LayoutCachedLeft =11910
                                            LayoutCachedTop =5880
                                            LayoutCachedWidth =13110
                                            LayoutCachedHeight =6135
                                            RowStart =2
                                            RowEnd =2
                                            LayoutGroup =6
                                            ThemeFontIndex =1
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =103
                                        End
                                    End
                                End
                                Begin TextBox
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =13170
                                    Top =6195
                                    Width =1770
                                    FontSize =10
                                    TabIndex =6
                                    Name ="Text301"
                                    ControlSource ="UnitRAFSerial"
                                    StatusBarText ="rear axle forward serial"
                                    GroupTable =103

                                    LayoutCachedLeft =13170
                                    LayoutCachedTop =6195
                                    LayoutCachedWidth =14940
                                    LayoutCachedHeight =6435
                                    RowStart =3
                                    RowEnd =3
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =6
                                    BackThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =103
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =11910
                                            Top =6195
                                            Width =1200
                                            Height =240
                                            FontSize =8
                                            FontWeight =700
                                            Name ="Label302"
                                            Caption ="REAR AXLE FWD"
                                            GroupTable =103
                                            LayoutCachedLeft =11910
                                            LayoutCachedTop =6195
                                            LayoutCachedWidth =13110
                                            LayoutCachedHeight =6435
                                            RowStart =3
                                            RowEnd =3
                                            LayoutGroup =6
                                            ThemeFontIndex =1
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =103
                                        End
                                    End
                                End
                                Begin TextBox
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =13170
                                    Top =6495
                                    Width =1770
                                    Height =285
                                    FontSize =10
                                    TabIndex =7
                                    Name ="Text299"
                                    ControlSource ="UnitRARSerial"
                                    StatusBarText ="rear axle rear serial"
                                    GroupTable =103

                                    LayoutCachedLeft =13170
                                    LayoutCachedTop =6495
                                    LayoutCachedWidth =14940
                                    LayoutCachedHeight =6780
                                    RowStart =4
                                    RowEnd =4
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =6
                                    BackThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =103
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =11910
                                            Top =6495
                                            Width =1200
                                            Height =285
                                            FontSize =8
                                            FontWeight =700
                                            Name ="Label300"
                                            Caption ="REAR AXLE REAR"
                                            GroupTable =103
                                            LayoutCachedLeft =11910
                                            LayoutCachedTop =6495
                                            LayoutCachedWidth =13110
                                            LayoutCachedHeight =6780
                                            RowStart =4
                                            RowEnd =4
                                            LayoutGroup =6
                                            ThemeFontIndex =1
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =103
                                        End
                                    End
                                End
                                Begin TextBox
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =13275
                                    Top =7320
                                    Width =3075
                                    Height =285
                                    FontSize =10
                                    Name ="UnitTransmissionType"
                                    ControlSource ="UnitTransmissionType"
                                    StatusBarText ="rear axle rear serial"
                                    GroupTable =115

                                    LayoutCachedLeft =13275
                                    LayoutCachedTop =7320
                                    LayoutCachedWidth =16350
                                    LayoutCachedHeight =7605
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =12
                                    BackThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =115
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =11895
                                            Top =7320
                                            Width =1320
                                            Height =285
                                            FontSize =8
                                            FontWeight =700
                                            Name ="Label4885"
                                            Caption ="Transmission Type"
                                            GroupTable =115
                                            LayoutCachedLeft =11895
                                            LayoutCachedTop =7320
                                            LayoutCachedWidth =13215
                                            LayoutCachedHeight =7605
                                            LayoutGroup =12
                                            ThemeFontIndex =1
                                            BackThemeColorIndex =2
                                            ForeThemeColorIndex =1
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =115
                                        End
                                    End
                                End
                                Begin TextBox
                                    SpecialEffect =5
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =13275
                                    Top =7665
                                    Width =3075
                                    Height =285
                                    FontSize =10
                                    Name ="UnitEngineType"
                                    ControlSource ="UnitEngineType"
                                    StatusBarText ="rear axle rear serial"
                                    GroupTable =115

                                    LayoutCachedLeft =13275
                                    LayoutCachedTop =7665
                                    LayoutCachedWidth =16350
                                    LayoutCachedHeight =7950
                                    RowStart =1
                                    RowEnd =1
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =12
                                    BackThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =115
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =11895
                                            Top =7665
                                            Width =1320
                                            Height =285
                                            FontSize =8
                                            FontWeight =700
                                            Name ="Label4889"
                                            Caption ="Engine Type"
                                            GroupTable =115
                                            LayoutCachedLeft =11895
                                            LayoutCachedTop =7665
                                            LayoutCachedWidth =13215
                                            LayoutCachedHeight =7950
                                            RowStart =1
                                            RowEnd =1
                                            LayoutGroup =12
                                            ThemeFontIndex =1
                                            BackThemeColorIndex =2
                                            ForeThemeColorIndex =1
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =115
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =60
                            Top =4215
                            Width =17925
                            Height =7575
                            BorderColor =0
                            Name ="Miles/Hours"
                            EventProcPrefix ="Miles_Hours"
                            GridlineColor =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =4215
                            LayoutCachedWidth =17985
                            LayoutCachedHeight =11790
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureAlignment =0
                                    PictureType =2
                                    Left =120
                                    Top =4740
                                    Width =17789
                                    Height =225
                                    Name ="Image4333"
                                    Picture ="1_WhiteShadow"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =120
                                    LayoutCachedTop =4740
                                    LayoutCachedWidth =17909
                                    LayoutCachedHeight =4965
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =1
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =60
                                    Top =4890
                                    Width =17820
                                    Height =6855
                                    TabIndex =1
                                    Name ="ViewMiNew"
                                    SourceObject ="Form.ViewMiNew"
                                    LinkChildFields ="UnitID"
                                    LinkMasterFields ="UnitID"

                                    LayoutCachedLeft =60
                                    LayoutCachedTop =4890
                                    LayoutCachedWidth =17880
                                    LayoutCachedHeight =11745
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =120
                                    Top =4410
                                    Width =13544
                                    Height =390
                                    FontSize =12
                                    Name ="Label4332"
                                    Caption ="Mileage Readings"
                                    FontName ="Segoe UI"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    HorizontalAnchor =2
                                    LayoutCachedLeft =120
                                    LayoutCachedTop =4410
                                    LayoutCachedWidth =13664
                                    LayoutCachedHeight =4800
                                    ColumnEnd =1
                                    BackThemeColorIndex =6
                                    BackShade =75.0
                                    BorderThemeColorIndex =6
                                    BorderShade =75.0
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin Label
                                    BackStyle =1
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =12825
                                    Top =4410
                                    Width =5069
                                    Height =390
                                    FontSize =12
                                    Name ="Label4334"
                                    Caption ="MPG Goals"
                                    FontName ="Segoe UI"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0
                                    HorizontalAnchor =2
                                    LayoutCachedLeft =12825
                                    LayoutCachedTop =4410
                                    LayoutCachedWidth =17894
                                    LayoutCachedHeight =4800
                                    ColumnEnd =1
                                    BackThemeColorIndex =6
                                    BackShade =75.0
                                    BorderThemeColorIndex =6
                                    BorderShade =75.0
                                    ForeThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =15045
                                    Top =4485
                                    Width =675
                                    Height =270
                                    FontSize =10
                                    TabIndex =2
                                    Name ="Text4335"
                                    ControlSource ="DrivingMpgGoal"
                                    GroupTable =110

                                    LayoutCachedLeft =15045
                                    LayoutCachedTop =4485
                                    LayoutCachedWidth =15720
                                    LayoutCachedHeight =4755
                                    ColumnStart =1
                                    ColumnEnd =1
                                    LayoutGroup =8
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =110
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =14355
                                            Top =4485
                                            Width =630
                                            Height =270
                                            FontSize =8
                                            Name ="Label4336"
                                            Caption ="Driving"
                                            FontName ="Segoe UI"
                                            GroupTable =110
                                            LayoutCachedLeft =14355
                                            LayoutCachedTop =4485
                                            LayoutCachedWidth =14985
                                            LayoutCachedHeight =4755
                                            LayoutGroup =8
                                            ForeThemeColorIndex =1
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =110
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =16500
                                    Top =4485
                                    Width =690
                                    Height =270
                                    FontSize =10
                                    TabIndex =3
                                    Name ="Text4337"
                                    ControlSource ="MpgGoal"
                                    GroupTable =110

                                    LayoutCachedLeft =16500
                                    LayoutCachedTop =4485
                                    LayoutCachedWidth =17190
                                    LayoutCachedHeight =4755
                                    ColumnStart =3
                                    ColumnEnd =3
                                    LayoutGroup =8
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    GroupTable =110
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =15780
                                            Top =4485
                                            Width =660
                                            Height =270
                                            FontSize =8
                                            Name ="Label4338"
                                            Caption ="Overall"
                                            FontName ="Segoe UI"
                                            GroupTable =110
                                            LayoutCachedLeft =15780
                                            LayoutCachedTop =4485
                                            LayoutCachedWidth =16440
                                            LayoutCachedHeight =4755
                                            ColumnStart =2
                                            ColumnEnd =2
                                            LayoutGroup =8
                                            ForeThemeColorIndex =1
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                            GroupTable =110
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =4215
                            Width =17850
                            Height =7575
                            BorderColor =0
                            Name ="Recalls Etc."
                            EventProcPrefix ="Recalls_Etc_"
                            GridlineColor =0
                            LayoutCachedLeft =135
                            LayoutCachedTop =4215
                            LayoutCachedWidth =17985
                            LayoutCachedHeight =11790
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =240
                                    Top =4665
                                    Width =16545
                                    Height =5835
                                    Name ="UnitRcallsub"
                                    SourceObject ="Form.UnitRcallsub"
                                    LinkChildFields ="UnitID"
                                    LinkMasterFields ="UnitID"

                                    LayoutCachedLeft =240
                                    LayoutCachedTop =4665
                                    LayoutCachedWidth =16785
                                    LayoutCachedHeight =10500
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =60
                            Top =4215
                            Width =17925
                            Height =7575
                            BorderColor =0
                            Name ="Remarketing"
                            Caption ="Remarketing Info"
                            GridlineColor =0
                            LayoutCachedLeft =60
                            LayoutCachedTop =4215
                            LayoutCachedWidth =17985
                            LayoutCachedHeight =11790
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =60
                                    Top =4440
                                    Width =13290
                                    Height =7050
                                    Name ="Units_RMKT_Sub"
                                    SourceObject ="Form.Units_RMKT_Sub"
                                    LinkChildFields ="UnitRef"
                                    LinkMasterFields ="UnitID"

                                    LayoutCachedLeft =60
                                    LayoutCachedTop =4440
                                    LayoutCachedWidth =13350
                                    LayoutCachedHeight =11490
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =13410
                                    Top =4395
                                    Width =4410
                                    Height =7035
                                    TabIndex =1
                                    Name ="InspectionsSubForm"
                                    SourceObject ="Form.InspectionsSubForm"
                                    LinkChildFields ="AssetID"
                                    LinkMasterFields ="UnitID"

                                    LayoutCachedLeft =13410
                                    LayoutCachedTop =4395
                                    LayoutCachedWidth =17820
                                    LayoutCachedHeight =11430
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =4215
                            Width =17850
                            Height =7575
                            BorderColor =0
                            Name ="UnitChanges"
                            Caption ="Unit Changes"
                            GridlineColor =0
                            LayoutCachedLeft =135
                            LayoutCachedTop =4215
                            LayoutCachedWidth =17985
                            LayoutCachedHeight =11790
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =150
                                    Top =4410
                                    Width =13560
                                    Height =7335
                                    Name ="UnitRevisionLog"
                                    SourceObject ="Form.UnitRevisionLog"
                                    LinkChildFields ="AssetID"
                                    LinkMasterFields ="UnitID"

                                    LayoutCachedLeft =150
                                    LayoutCachedTop =4410
                                    LayoutCachedWidth =13710
                                    LayoutCachedHeight =11745
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =4215
                            Width =17850
                            Height =7575
                            BorderColor =0
                            Name ="Hidden"
                            GridlineColor =0
                            LayoutCachedLeft =135
                            LayoutCachedTop =4215
                            LayoutCachedWidth =17985
                            LayoutCachedHeight =11790
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            Begin
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =450
                                    Top =8535
                                    Width =1725
                                    Height =315
                                    FontSize =10
                                    Name ="CACarbCompDT"
                                    ControlSource ="CACarbCompDT"
                                    StatusBarText ="DT :tractors must comply with the CA CARB Compliance Law:the tractors become obs"
                                        "olete in CA or the must have a diesel particulate filter installed."

                                    LayoutCachedLeft =450
                                    LayoutCachedTop =8535
                                    LayoutCachedWidth =2175
                                    LayoutCachedHeight =8850
                                    BackThemeColorIndex =1
                                    GridlineThemeColorIndex =1
                                    GridlineShade =65.0
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =450
                                            Top =8295
                                            Width =1725
                                            Height =255
                                            FontSize =8
                                            Name ="Label387"
                                            Caption ="CA ARB Compliance"
                                            FontName ="Segoe UI"
                                            LayoutCachedLeft =450
                                            LayoutCachedTop =8295
                                            LayoutCachedWidth =2175
                                            LayoutCachedHeight =8550
                                            ForeThemeColorIndex =2
                                            GridlineThemeColorIndex =1
                                            GridlineShade =65.0
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =60
                            Top =4215
                            Width =17925
                            Height =7575
                            Name ="LeaseEndOptions"
                            Caption ="EXIT Options"
                            LayoutCachedLeft =60
                            LayoutCachedTop =4215
                            LayoutCachedWidth =17985
                            LayoutCachedHeight =11790
                            Begin
                                Begin Image
                                    PictureTiling = NotDefault
                                    SizeMode =1
                                    PictureAlignment =0
                                    PictureType =2
                                    Left =75
                                    Top =4710
                                    Width =17877
                                    Height =240
                                    Name ="Image4347"
                                    Picture ="1_WhiteShadow"
                                    LeftPadding =0
                                    TopPadding =0
                                    RightPadding =0
                                    BottomPadding =0

                                    LayoutCachedLeft =75
                                    LayoutCachedTop =4710
                                    LayoutCachedWidth =17952
                                    LayoutCachedHeight =4950
                                    RowStart =2
                                    RowEnd =2
                                    ColumnEnd =1
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =60
                                    Top =4365
                                    Width =17880
                                    Height =435
                                    FontSize =18
                                    Name ="Label4348"
                                    Caption ="EXIT Options"
                                    FontName ="Segoe UI"
                                    LayoutCachedLeft =60
                                    LayoutCachedTop =4365
                                    LayoutCachedWidth =17940
                                    LayoutCachedHeight =4800
                                    BackThemeColorIndex =1
                                    BackShade =85.0
                                    ForeThemeColorIndex =2
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =180
                                    Top =5025
                                    Width =4766
                                    Height =6719
                                    TabIndex =1
                                    Name ="Units_Exit_ExchangeIT"
                                    SourceObject ="Form.Units_Exit_ExchangeIT"
                                    LinkChildFields ="UnitID"
                                    LinkMasterFields ="UnitID"

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =5025
                                    LayoutCachedWidth =4946
                                    LayoutCachedHeight =11744
                                    BorderThemeColorIndex =1
                                    BorderShade =95.0
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =5040
                                    Top =5025
                                    Width =4766
                                    Height =6720
                                    TabIndex =2
                                    Name ="Units_Exit_Extend"
                                    SourceObject ="Form.Units_Exit_Extend"
                                    LinkChildFields ="UnitID"
                                    LinkMasterFields ="UnitID"

                                    LayoutCachedLeft =5040
                                    LayoutCachedTop =5025
                                    LayoutCachedWidth =9806
                                    LayoutCachedHeight =11745
                                    BorderThemeColorIndex =1
                                    BorderShade =95.0
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =9960
                                    Top =5025
                                    Width =4766
                                    Height =6719
                                    TabIndex =3
                                    Name ="Units_Exit_Terminate"
                                    SourceObject ="Form.Units_Exit_Terminate"
                                    LinkChildFields ="UnitID"
                                    LinkMasterFields ="UnitID"

                                    LayoutCachedLeft =9960
                                    LayoutCachedTop =5025
                                    LayoutCachedWidth =14726
                                    LayoutCachedHeight =11744
                                    BorderThemeColorIndex =1
                                    BorderShade =95.0
                                End
                            End
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =16425
                    Top =690
                    Width =630
                    Height =960
                    Name ="Box3883"
                    LayoutCachedLeft =16425
                    LayoutCachedTop =690
                    LayoutCachedWidth =17055
                    LayoutCachedHeight =1650
                    BorderThemeColorIndex =6
                    BorderShade =50.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =223
                    Left =16272
                    Top =660
                    Width =315
                    Height =945
                    Name ="Box3882"
                    LayoutCachedLeft =16272
                    LayoutCachedTop =660
                    LayoutCachedWidth =16587
                    LayoutCachedHeight =1605
                    BackThemeColorIndex =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =150
                    Top =2145
                    Width =4185
                    Height =240
                    FontSize =8
                    BackColor =12349952
                    Name ="Label3384"
                    Caption =" Purchase Lease Back (acquisition)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =150
                    LayoutCachedTop =2145
                    LayoutCachedWidth =4335
                    LayoutCachedHeight =2385
                    ColumnEnd =1
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Left =18075
                    Width =174
                    Height =11895
                    BorderColor =16777215
                    Name ="Image2243"
                    Picture ="LtGrey1_VertShadow"
                    VerticalAnchor =2

                    LayoutCachedLeft =18075
                    LayoutCachedWidth =18249
                    LayoutCachedHeight =11895
                    TabIndex =77
                    BorderShade =100.0
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureType =2
                    Width =18147
                    Height =210
                    BorderColor =16777215
                    Name ="Image767"
                    Picture ="LtGrey2_Shadow"

                    LayoutCachedWidth =18147
                    LayoutCachedHeight =210
                    TabIndex =76
                    BorderShade =100.0
                End
                Begin Image
                    SizeMode =1
                    PictureType =2
                    Left =17865
                    Top =15
                    Width =375
                    Height =465
                    BorderColor =16777215
                    Name ="Image2244"
                    Picture ="1_LtGrey1_CornerShadow"

                    LayoutCachedLeft =17865
                    LayoutCachedTop =15
                    LayoutCachedWidth =18240
                    LayoutCachedHeight =480
                    TabIndex =78
                    BorderShade =100.0
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =127
                    Left =18240
                    Width =3105
                    Height =11895
                    Name ="Box2167"
                    HorizontalAnchor =2
                    VerticalAnchor =2
                    LayoutCachedLeft =18240
                    LayoutCachedWidth =21345
                    LayoutCachedHeight =11895
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3000
                    Top =1200
                    Width =1320
                    ColumnOrder =21
                    FontSize =10
                    TabIndex =15
                    Name ="UnitGVW"
                    ControlSource ="UnitGVW"
                    Format ="Standard"
                    StatusBarText ="GVW from MSO"
                    ValidationRule =">=0"
                    ValidationText ="GVW must be >= 0"
                    DefaultValue ="0"
                    GroupTable =66
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =3000
                    LayoutCachedTop =1200
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2520
                            Top =1200
                            Width =480
                            Height =240
                            FontSize =8
                            Name ="Label9"
                            Caption ="GVW"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2520
                            LayoutCachedTop =1200
                            LayoutCachedWidth =3000
                            LayoutCachedHeight =1440
                            RowStart =2
                            RowEnd =2
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9690
                    Top =570
                    Width =2248
                    Height =270
                    ColumnOrder =25
                    FontSize =10
                    TabIndex =4
                    Name ="UnitInvoiceNum"
                    ControlSource ="UnitInvoiceNum"
                    GroupTable =66

                    LayoutCachedLeft =9690
                    LayoutCachedTop =570
                    LayoutCachedWidth =11938
                    LayoutCachedHeight =840
                    ColumnStart =13
                    ColumnEnd =14
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8640
                            Top =570
                            Width =1027
                            Height =270
                            FontSize =8
                            Name ="Label16"
                            Caption ="Inv Num"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =8640
                            LayoutCachedTop =570
                            LayoutCachedWidth =9667
                            LayoutCachedHeight =840
                            ColumnStart =12
                            ColumnEnd =12
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13050
                    Top =570
                    Width =1140
                    Height =270
                    ColumnOrder =22
                    FontSize =10
                    TabIndex =5
                    Name ="UnitTASent"
                    ControlSource ="UnitTASent"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =66

                    LayoutCachedLeft =13050
                    LayoutCachedTop =570
                    LayoutCachedWidth =14190
                    LayoutCachedHeight =840
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =12135
                            Top =570
                            Width =885
                            Height =270
                            FontSize =8
                            Name ="Label21"
                            Caption ="App Sent"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =12135
                            LayoutCachedTop =570
                            LayoutCachedWidth =13020
                            LayoutCachedHeight =840
                            ColumnStart =16
                            ColumnEnd =16
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13050
                    Top =900
                    Width =1140
                    ColumnOrder =23
                    FontSize =10
                    TabIndex =12
                    Name ="UnitPOA"
                    ControlSource ="UnitPOA"
                    Format ="Short Date"
                    StatusBarText ="lookup from Contacts table"
                    GroupTable =66

                    LayoutCachedLeft =13050
                    LayoutCachedTop =900
                    LayoutCachedWidth =14190
                    LayoutCachedHeight =1140
                    RowStart =1
                    RowEnd =1
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =12135
                            Top =900
                            Width =885
                            Height =240
                            FontSize =8
                            Name ="Label25"
                            Caption ="POA Sent"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =12135
                            LayoutCachedTop =900
                            LayoutCachedWidth =13020
                            LayoutCachedHeight =1140
                            RowStart =1
                            RowEnd =1
                            ColumnStart =16
                            ColumnEnd =16
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5235
                    Top =2475
                    Width =1080
                    Height =255
                    ColumnOrder =30
                    FontSize =10
                    TabIndex =34
                    BackColor =15523798
                    Name ="UnitPDStartDate"
                    ControlSource ="UnitPDStartDate"
                    Format ="Short Date"
                    StatusBarText ="Unit per diem start date"
                    GroupTable =106

                    LayoutCachedLeft =5235
                    LayoutCachedTop =2475
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =2730
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4470
                            Top =2475
                            Width =705
                            Height =255
                            FontSize =8
                            Name ="Label35"
                            Caption ="Start"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =4470
                            LayoutCachedTop =2475
                            LayoutCachedWidth =5175
                            LayoutCachedHeight =2730
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =7
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5235
                    Top =2790
                    Width =1080
                    Height =255
                    ColumnOrder =31
                    FontSize =10
                    TabIndex =38
                    BackColor =15523798
                    Name ="UNITPDRENT"
                    ControlSource ="UNITPDRENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="For FSL Lease (FSL pmt + rent+Tax)/30; all others Rent/30"
                    ValidationRule =">=0"
                    ValidationText ="PD Rent must be greater than of = to 0"
                    DefaultValue ="0"
                    ControlTipText ="For FSL Lease (FSL pmt + rent+Tax)/30; all others Rent/30"
                    GroupTable =106

                    LayoutCachedLeft =5235
                    LayoutCachedTop =2790
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =3045
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4470
                            Top =2790
                            Width =705
                            Height =255
                            FontSize =8
                            Name ="Label36"
                            Caption ="Amount"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =4470
                            LayoutCachedTop =2790
                            LayoutCachedWidth =5175
                            LayoutCachedHeight =3045
                            RowStart =1
                            RowEnd =1
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =7
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7425
                    Top =2475
                    Width =1065
                    Height =255
                    ColumnOrder =32
                    FontSize =10
                    TabIndex =35
                    BackColor =15523798
                    Name ="LORent"
                    ControlSource ="UNITSRENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="MONTLHY STRIP RENT"
                    ValidationRule =">=0"
                    ValidationText ="Strip rent must be greater than of = to 0"
                    DefaultValue ="0"
                    GroupTable =106

                    LayoutCachedLeft =7425
                    LayoutCachedTop =2475
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =2730
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =6450
                            Top =2475
                            Width =915
                            Height =255
                            FontSize =8
                            Name ="Label79"
                            Caption ="Schedule"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =6450
                            LayoutCachedTop =2475
                            LayoutCachedWidth =7365
                            LayoutCachedHeight =2730
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =7
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7425
                    Top =3105
                    Width =1065
                    Height =255
                    ColumnOrder =33
                    FontSize =10
                    TabIndex =43
                    BackColor =15523798
                    Name ="SyndRENT"
                    ControlSource ="UNITRENT"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="MONTHLY RENT"
                    ValidationRule =">=0"
                    ValidationText ="Syn Rent must be greater than of = to 0"
                    DefaultValue ="0"
                    GroupTable =106

                    LayoutCachedLeft =7425
                    LayoutCachedTop =3105
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =3360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =6450
                            Top =3105
                            Width =915
                            Height =255
                            FontSize =8
                            Name ="Label80"
                            Caption ="Syndicated"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =6450
                            LayoutCachedTop =3105
                            LayoutCachedWidth =7365
                            LayoutCachedHeight =3360
                            RowStart =2
                            RowEnd =2
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =7
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9960
                    Top =2790
                    Width =1215
                    Height =255
                    ColumnOrder =38
                    FontSize =10
                    TabIndex =40
                    Name ="UnitSynResid"
                    ControlSource ="UnitSynResid"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="residual set at funding"
                    ValidationRule =">=0"
                    ValidationText ="Syn Resid must be >=0"
                    DefaultValue ="0"
                    GroupTable =106

                    LayoutCachedLeft =9960
                    LayoutCachedTop =2790
                    LayoutCachedWidth =11175
                    LayoutCachedHeight =3045
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =7
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =8655
                            Top =2790
                            Width =1245
                            Height =255
                            FontSize =8
                            Name ="Label177"
                            Caption ="Syndication"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =8655
                            LayoutCachedTop =2790
                            LayoutCachedWidth =9900
                            LayoutCachedHeight =3045
                            RowStart =1
                            RowEnd =1
                            ColumnStart =9
                            ColumnEnd =9
                            LayoutGroup =7
                            BorderThemeColorIndex =2
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19560
                    Top =6705
                    Width =750
                    FontSize =8
                    TabIndex =62
                    BorderColor =0
                    ForeColor =5026082
                    Name ="SGrpID"
                    ControlSource ="SGrpID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19560
                    LayoutCachedTop =6705
                    LayoutCachedWidth =20310
                    LayoutCachedHeight =6945
                    RowStart =4
                    RowEnd =4
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9960
                    Top =2475
                    Width =1215
                    Height =255
                    FontSize =10
                    TabIndex =36
                    BackColor =15523798
                    Name ="UnitSalesResid"
                    ControlSource ="UnitSalesResid"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    ValidationText ="Sales Resid must be >=0"
                    DefaultValue ="0"
                    Tag ="Cloned"
                    GroupTable =106

                    LayoutCachedLeft =9960
                    LayoutCachedTop =2475
                    LayoutCachedWidth =11175
                    LayoutCachedHeight =2730
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =7
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8655
                            Top =2475
                            Width =1245
                            Height =255
                            FontSize =8
                            Name ="Label335"
                            Caption ="Schedule"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =8655
                            LayoutCachedTop =2475
                            LayoutCachedWidth =9900
                            LayoutCachedHeight =2730
                            ColumnStart =9
                            ColumnEnd =9
                            LayoutGroup =7
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13050
                    Top =1200
                    Width =1140
                    FontSize =10
                    TabIndex =18
                    Name ="UnitTitleRecd"
                    ControlSource ="UnitTitleRecd"
                    Format ="Short Date"
                    ControlTipText ="Accountable Mail: entered when title checked-in"
                    GroupTable =66

                    LayoutCachedLeft =13050
                    LayoutCachedTop =1200
                    LayoutCachedWidth =14190
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =12135
                            Top =1200
                            Width =885
                            Height =240
                            FontSize =8
                            Name ="Label337"
                            Caption ="Title Recvd"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =12135
                            LayoutCachedTop =1200
                            LayoutCachedWidth =13020
                            LayoutCachedHeight =1440
                            RowStart =2
                            RowEnd =2
                            ColumnStart =16
                            ColumnEnd =16
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9690
                    Top =1200
                    Width =1123
                    FontSize =10
                    TabIndex =17
                    Name ="UnitVendorPytDate"
                    ControlSource ="UnitVendorPytDate"
                    Format ="Short Date"
                    GroupTable =66

                    LayoutCachedLeft =9690
                    LayoutCachedTop =1200
                    LayoutCachedWidth =10813
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8640
                            Top =1200
                            Width =1027
                            Height =240
                            FontSize =8
                            Name ="Label5040"
                            Caption ="Inv Paid"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =8640
                            LayoutCachedTop =1200
                            LayoutCachedWidth =9667
                            LayoutCachedHeight =1440
                            RowStart =2
                            RowEnd =2
                            ColumnStart =12
                            ColumnEnd =12
                            LayoutGroup =3
                            BackThemeColorIndex =9
                            BackTint =20.0
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9690
                    Top =1800
                    Width =1123
                    FontSize =10
                    TabIndex =25
                    Name ="UnitDep"
                    ControlSource ="UnitDep"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =66

                    LayoutCachedLeft =9690
                    LayoutCachedTop =1800
                    LayoutCachedWidth =10813
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8640
                            Top =1800
                            Width =1027
                            Height =240
                            FontSize =8
                            Name ="Label412"
                            Caption ="Dep / Paid"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =8640
                            LayoutCachedTop =1800
                            LayoutCachedWidth =9667
                            LayoutCachedHeight =2040
                            RowStart =4
                            RowEnd =4
                            ColumnStart =12
                            ColumnEnd =12
                            LayoutGroup =3
                            BackThemeColorIndex =9
                            BackTint =20.0
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10875
                    Top =1800
                    Width =1063
                    FontSize =10
                    TabIndex =26
                    Name ="UnitDepPdDT"
                    ControlSource ="UnitDepPdDT"
                    Format ="Short Date"
                    GroupTable =66

                    LayoutCachedLeft =10875
                    LayoutCachedTop =1800
                    LayoutCachedWidth =11938
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9690
                    Top =900
                    Width =1123
                    FontSize =10
                    TabIndex =11
                    Name ="Text421"
                    ControlSource ="UnitInvoiceDate"
                    Format ="Short Date"
                    GroupTable =66

                    LayoutCachedLeft =9690
                    LayoutCachedTop =900
                    LayoutCachedWidth =10813
                    LayoutCachedHeight =1140
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8640
                            Top =900
                            Width =1027
                            Height =240
                            FontSize =8
                            Name ="Label422"
                            Caption ="Inv Recv"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =8640
                            LayoutCachedTop =900
                            LayoutCachedWidth =9667
                            LayoutCachedHeight =1140
                            RowStart =1
                            RowEnd =1
                            ColumnStart =12
                            ColumnEnd =12
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19560
                    Top =6225
                    Width =750
                    FontSize =8
                    TabIndex =60
                    BorderColor =0
                    ForeColor =5026082
                    Name ="ClientID"
                    ControlSource ="Clients.clientID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19560
                    LayoutCachedTop =6225
                    LayoutCachedWidth =20310
                    LayoutCachedHeight =6465
                    RowStart =2
                    RowEnd =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19560
                    Top =5985
                    Width =750
                    FontSize =8
                    TabIndex =59
                    BorderColor =0
                    ForeColor =5026082
                    Name ="Text373"
                    ControlSource ="UnitID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19560
                    LayoutCachedTop =5985
                    LayoutCachedWidth =20310
                    LayoutCachedHeight =6225
                    RowStart =1
                    RowEnd =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19560
                    Top =7185
                    Width =750
                    Height =315
                    FontSize =8
                    TabIndex =64
                    BorderColor =0
                    ForeColor =5026082
                    Name ="MLLessor"
                    ControlSource ="MLLessor"
                    StatusBarText ="If other name of lessor"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19560
                    LayoutCachedTop =7185
                    LayoutCachedWidth =20310
                    LayoutCachedHeight =7500
                    RowStart =6
                    RowEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13050
                    Top =1500
                    Width =1140
                    ColumnWidth =1875
                    FontSize =10
                    TabIndex =23
                    Name ="UnitTitleNumber"
                    ControlSource ="UnitTitleNumber"
                    OnGotFocus ="[Event Procedure]"
                    GroupTable =66

                    LayoutCachedLeft =13050
                    LayoutCachedTop =1500
                    LayoutCachedWidth =14190
                    LayoutCachedHeight =1740
                    RowStart =3
                    RowEnd =3
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =12135
                            Top =1500
                            Width =885
                            Height =240
                            FontSize =8
                            Name ="Label444"
                            Caption ="Title #"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =12135
                            LayoutCachedTop =1500
                            LayoutCachedWidth =13020
                            LayoutCachedHeight =1740
                            RowStart =3
                            RowEnd =3
                            ColumnStart =16
                            ColumnEnd =16
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2385
                    Top =2790
                    Width =1935
                    Height =255
                    FontSize =10
                    TabIndex =37
                    Name ="AcqMileage"
                    ControlSource ="AcqMileage"
                    Format ="Standard"
                    ControlTipText ="Miles when the unit was acquired"
                    GroupTable =106

                    LayoutCachedLeft =2385
                    LayoutCachedTop =2790
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =3045
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =180
                            Top =2790
                            Width =2145
                            Height =255
                            FontSize =8
                            Name ="Label466"
                            Caption ="Odometer  @ acquisition"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =180
                            LayoutCachedTop =2790
                            LayoutCachedWidth =2325
                            LayoutCachedHeight =3045
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =7
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5070
                    Top =570
                    Width =1140
                    Height =270
                    FontSize =10
                    TabIndex =2
                    Name ="UnitCOASent"
                    ControlSource ="UnitCOASent"
                    Format ="Short Date"
                    ControlTipText ="Date we sent a request"
                    GroupTable =66

                    LayoutCachedLeft =5070
                    LayoutCachedTop =570
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =840
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4485
                            Top =570
                            Width =555
                            Height =270
                            FontSize =8
                            Name ="Label19"
                            Caption ="Sent"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =4485
                            LayoutCachedTop =570
                            LayoutCachedWidth =5040
                            LayoutCachedHeight =840
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5070
                    Top =900
                    Width =1140
                    FontSize =10
                    TabIndex =9
                    Name ="UnitCofACopy"
                    ControlSource ="UnitCofACopy"
                    Format ="Short Date"
                    StatusBarText ="Date we recv'd a copy of the COA"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =66

                    LayoutCachedLeft =5070
                    LayoutCachedTop =900
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =1140
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4485
                            Top =900
                            Width =555
                            Height =240
                            FontSize =8
                            Name ="Label237"
                            Caption ="Copy"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =4485
                            LayoutCachedTop =900
                            LayoutCachedWidth =5040
                            LayoutCachedHeight =1140
                            RowStart =1
                            RowEnd =1
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7275
                    Top =570
                    Width =1155
                    Height =270
                    ColumnOrder =18
                    FontSize =10
                    TabIndex =3
                    Name ="UnitAcceptDate"
                    ControlSource ="UnitAcceptDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =66

                    LayoutCachedLeft =7275
                    LayoutCachedTop =570
                    LayoutCachedWidth =8430
                    LayoutCachedHeight =840
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =6405
                            Top =570
                            Width =840
                            Height =270
                            FontSize =8
                            Name ="Label18"
                            Caption ="Accepted"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =6405
                            LayoutCachedTop =570
                            LayoutCachedWidth =7245
                            LayoutCachedHeight =840
                            ColumnStart =9
                            ColumnEnd =9
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1155
                    Top =570
                    Width =1170
                    Height =270
                    ColumnOrder =27
                    FontSize =10
                    BackColor =15523798
                    Name ="UnitBuildDate"
                    ControlSource ="[UnitBuildDate]"
                    Format ="Short Date"
                    GroupTable =66

                    LayoutCachedLeft =1155
                    LayoutCachedTop =570
                    LayoutCachedWidth =2325
                    LayoutCachedHeight =840
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =150
                            Top =570
                            Width =975
                            Height =270
                            FontSize =8
                            Name ="Label81"
                            Caption ="Build Date"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =150
                            LayoutCachedTop =570
                            LayoutCachedWidth =1125
                            LayoutCachedHeight =840
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1155
                    Top =900
                    Width =1170
                    FontSize =10
                    TabIndex =7
                    BackColor =15523798
                    Name ="UnitEstDelDate"
                    ControlSource ="UnitEstDelDate"
                    Format ="Short Date"
                    GroupTable =66

                    LayoutCachedLeft =1155
                    LayoutCachedTop =900
                    LayoutCachedWidth =2325
                    LayoutCachedHeight =1140
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =150
                            Top =900
                            Width =975
                            Height =240
                            FontSize =8
                            Name ="Label83"
                            Caption ="Estimated"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =150
                            LayoutCachedTop =900
                            LayoutCachedWidth =1125
                            LayoutCachedHeight =1140
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1155
                    Top =1200
                    Width =1170
                    ColumnOrder =29
                    FontSize =10
                    TabIndex =14
                    Name ="UnitActualDelDate"
                    ControlSource ="UnitActualDelDate"
                    Format ="Short Date"
                    GroupTable =66

                    LayoutCachedLeft =1155
                    LayoutCachedTop =1200
                    LayoutCachedWidth =2325
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =150
                            Top =1200
                            Width =975
                            Height =240
                            FontSize =8
                            Name ="Label84"
                            Caption ="Actual"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =150
                            LayoutCachedTop =1200
                            LayoutCachedWidth =1125
                            LayoutCachedHeight =1440
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5070
                    Top =1200
                    Width =1140
                    FontSize =10
                    TabIndex =16
                    Name ="UnitCOABack"
                    ControlSource ="UnitCOABack"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Orginal COA. This field may only be entered/modified from Accountable Mail via C"
                        "OA Check In"
                    GroupTable =66

                    LayoutCachedLeft =5070
                    LayoutCachedTop =1200
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =9
                    BorderShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4485
                            Top =1200
                            Width =555
                            Height =240
                            FontSize =8
                            Name ="Label20"
                            Caption ="Orig"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =4485
                            LayoutCachedTop =1200
                            LayoutCachedWidth =5040
                            LayoutCachedHeight =1440
                            RowStart =2
                            RowEnd =2
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7275
                    Top =900
                    Width =1155
                    ColumnOrder =42
                    FontSize =10
                    TabIndex =10
                    Name ="UnitInServ"
                    ControlSource ="UnitInServ"
                    Format ="Short Date"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =66

                    LayoutCachedLeft =7275
                    LayoutCachedTop =900
                    LayoutCachedWidth =8430
                    LayoutCachedHeight =1140
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =6405
                            Top =900
                            Width =840
                            Height =240
                            FontSize =8
                            Name ="Label229"
                            Caption ="In Service"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =6405
                            LayoutCachedTop =900
                            LayoutCachedWidth =7245
                            LayoutCachedHeight =1140
                            RowStart =1
                            RowEnd =1
                            ColumnStart =9
                            ColumnEnd =9
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3000
                    Top =570
                    Width =1320
                    Height =270
                    FontSize =10
                    TabIndex =1
                    Name ="UnitMSOCopy"
                    ControlSource ="UnitMSOCopy"
                    Format ="Short Date"
                    StatusBarText ="Can only be updated through Accountable Mail"
                    ControlTipText ="Can only be updated through Accountable Mail"
                    GroupTable =66
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =3000
                    LayoutCachedTop =570
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =840
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2520
                            Top =570
                            Width =480
                            Height =270
                            FontSize =8
                            Name ="Label334"
                            Caption ="Copy"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2520
                            LayoutCachedTop =570
                            LayoutCachedWidth =3000
                            LayoutCachedHeight =840
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3000
                    Top =900
                    Width =1320
                    ColumnOrder =20
                    FontSize =10
                    TabIndex =8
                    Name ="UnitMSORecd"
                    ControlSource ="UnitMSORecd"
                    Format ="Short Date"
                    StatusBarText ="Can only be updated through Accountable Mail"
                    ControlTipText ="Can only be updated through Accountable Mail"
                    GroupTable =66
                    LeftPadding =0
                    RightPadding =0

                    LayoutCachedLeft =3000
                    LayoutCachedTop =900
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1140
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =2520
                            Top =900
                            Width =480
                            Height =240
                            FontSize =8
                            Name ="Label10"
                            Caption ="Orig"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =2520
                            LayoutCachedTop =900
                            LayoutCachedWidth =3000
                            LayoutCachedHeight =1140
                            RowStart =1
                            RowEnd =1
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin EmptyCell
                    Left =2385
                    Top =570
                    Width =105
                    Height =570
                    Name ="EmptyCell1463"
                    GroupTable =66
                    LayoutCachedLeft =2385
                    LayoutCachedTop =570
                    LayoutCachedWidth =2490
                    LayoutCachedHeight =1140
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =6270
                    Top =570
                    Width =105
                    Height =570
                    Name ="EmptyCell1492"
                    GroupTable =66
                    LayoutCachedLeft =6270
                    LayoutCachedTop =570
                    LayoutCachedWidth =6375
                    LayoutCachedHeight =1140
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =8490
                    Top =570
                    Width =120
                    Height =570
                    Name ="EmptyCell1507"
                    GroupTable =66
                    LayoutCachedLeft =8490
                    LayoutCachedTop =570
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =1140
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =12000
                    Top =570
                    Width =105
                    Height =1170
                    Name ="EmptyCell1537"
                    GroupTable =66
                    LayoutCachedLeft =12000
                    LayoutCachedTop =570
                    LayoutCachedWidth =12105
                    LayoutCachedHeight =1740
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =6
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7425
                    Top =2790
                    Width =1065
                    Height =255
                    ColumnOrder =1
                    FontSize =9
                    FontWeight =700
                    TabIndex =39
                    Name ="LRFLo"
                    Format ="Percent"
                    GroupTable =106

                    LayoutCachedLeft =7425
                    LayoutCachedTop =2790
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =3045
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =7
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =6450
                            Top =2790
                            Width =915
                            Height =255
                            FontSize =8
                            Name ="Label2284"
                            Caption ="LRF (Sch)"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =6450
                            LayoutCachedTop =2790
                            LayoutCachedWidth =7365
                            LayoutCachedHeight =3045
                            RowStart =1
                            RowEnd =1
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =7
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =6
                    SpecialEffect =3
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7425
                    Top =3420
                    Width =1065
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =46
                    Name ="LRFSyn"
                    Format ="Percent"
                    GroupTable =106

                    LayoutCachedLeft =7425
                    LayoutCachedTop =3420
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =3675
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =7
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =6450
                            Top =3420
                            Width =915
                            Height =255
                            FontSize =8
                            Name ="Label2294"
                            Caption ="LRF (Synd)"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =6450
                            LayoutCachedTop =3420
                            LayoutCachedWidth =7365
                            LayoutCachedHeight =3675
                            RowStart =3
                            RowEnd =3
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =7
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin EmptyCell
                    Left =6345
                    Top =2475
                    Width =75
                    Height =255
                    Name ="EmptyCell2672"
                    GroupTable =106
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =6345
                    LayoutCachedTop =2475
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =2730
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =6345
                    Top =2790
                    Width =75
                    Height =255
                    Name ="EmptyCell2673"
                    GroupTable =106
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =6345
                    LayoutCachedTop =2790
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =3045
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =6345
                    Top =3105
                    Width =75
                    Height =255
                    Name ="EmptyCell2674"
                    GroupTable =106
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =6345
                    LayoutCachedTop =3105
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =3360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =6345
                    Top =3420
                    Width =75
                    Height =255
                    Name ="EmptyCell2675"
                    GroupTable =106
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =6345
                    LayoutCachedTop =3420
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =3675
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =8625
                    Top =2145
                    Width =2550
                    Height =240
                    FontSize =8
                    BackColor =12349952
                    Name ="Label1755"
                    Caption =" Residual Values"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8625
                    LayoutCachedTop =2145
                    LayoutCachedWidth =11175
                    LayoutCachedHeight =2385
                    ColumnEnd =1
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =6420
                    Top =2145
                    Width =2085
                    Height =240
                    FontSize =8
                    BackColor =10040879
                    Name ="Label976"
                    Caption =" Rent"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6420
                    LayoutCachedTop =2145
                    LayoutCachedWidth =8505
                    LayoutCachedHeight =2385
                    ColumnEnd =1
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19560
                    Top =6945
                    Width =750
                    FontSize =8
                    TabIndex =63
                    BorderColor =0
                    ForeColor =5026082
                    Name ="MLOrig"
                    ControlSource ="MLOrig"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19560
                    LayoutCachedTop =6945
                    LayoutCachedWidth =20310
                    LayoutCachedHeight =7185
                    RowStart =5
                    RowEnd =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =19320
                    Top =3900
                    Width =1590
                    Height =246
                    FontSize =10
                    FontWeight =700
                    TabIndex =54
                    Name ="AddUnits"
                    Caption ="Add Unit(s)"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =19320
                    LayoutCachedTop =3900
                    LayoutCachedWidth =20910
                    LayoutCachedHeight =4146
                    ColumnStart =7
                    ColumnEnd =7
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =8210719
                    ThemeFontIndex =1
                    HoverThemeColorIndex =9
                    HoverShade =75.0
                    PressedColor =1643706
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19320
                    Top =4500
                    Width =1590
                    Height =375
                    ColumnOrder =16
                    FontSize =14
                    TabIndex =55
                    BorderColor =0
                    Name ="NoUnits"
                    Format ="General Number"

                    LayoutCachedLeft =19320
                    LayoutCachedTop =4500
                    LayoutCachedWidth =20910
                    LayoutCachedHeight =4875
                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =247
                            TextAlign =2
                            Left =19320
                            Top =4260
                            Width =1590
                            Height =225
                            FontSize =9
                            FontWeight =700
                            BackColor =0
                            Name ="Label1848"
                            Caption ="Units to Add"
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =19320
                            LayoutCachedTop =4260
                            LayoutCachedWidth =20910
                            LayoutCachedHeight =4485
                            RowStart =1
                            RowEnd =1
                            ForeThemeColorIndex =1
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    Left =20040
                    Top =4980
                    Width =975
                    Height =435
                    FontSize =9
                    Name ="Label339"
                    Caption ="= required \015\012= replicated "
                    LayoutCachedLeft =20040
                    LayoutCachedTop =4980
                    LayoutCachedWidth =21015
                    LayoutCachedHeight =5415
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    Left =19335
                    Top =4980
                    Width =675
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =56
                    BackColor =11271167
                    Name ="RequiredLabel"
                    ControlSource ="=\"Yellow\""
                    OnMouseMove ="[Event Procedure]"

                    LayoutCachedLeft =19335
                    LayoutCachedTop =4980
                    LayoutCachedWidth =20010
                    LayoutCachedHeight =5220
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =19335
                    Top =5205
                    Width =675
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =57
                    BackColor =15523798
                    Name ="ClonedLabel"
                    ControlSource ="=\"Lt Blue\""
                    OnMouseMove ="[Event Procedure]"

                    LayoutCachedLeft =19335
                    LayoutCachedTop =5205
                    LayoutCachedWidth =20010
                    LayoutCachedHeight =5445
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =4470
                    Top =2145
                    Width =1845
                    Height =240
                    FontSize =8
                    BackColor =10040879
                    Name ="Label3393"
                    Caption =" Per Diem"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4470
                    LayoutCachedTop =2145
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =2385
                    ColumnEnd =1
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =150
                    Top =240
                    Width =2190
                    Height =240
                    FontSize =8
                    BackColor =12349952
                    Name ="Label3397"
                    Caption =" Delivery"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =150
                    LayoutCachedTop =240
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =480
                    ColumnEnd =1
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =4485
                    Top =240
                    Width =1726
                    Height =240
                    FontSize =8
                    Name ="Label3401"
                    Caption =" COA"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4485
                    LayoutCachedTop =240
                    LayoutCachedWidth =6211
                    LayoutCachedHeight =480
                    ColumnEnd =1
                    BackThemeColorIndex =9
                    BackShade =75.0
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =6405
                    Top =240
                    Width =2025
                    Height =240
                    FontSize =8
                    Name ="Label3404"
                    Caption =" Acceptance"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6405
                    LayoutCachedTop =240
                    LayoutCachedWidth =8430
                    LayoutCachedHeight =480
                    ColumnEnd =1
                    BackThemeColorIndex =9
                    BackShade =75.0
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =8625
                    Top =240
                    Width =3333
                    Height =240
                    FontSize =8
                    BackColor =12349952
                    Name ="Label3407"
                    Caption =" Vendor Invoice"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8625
                    LayoutCachedTop =240
                    LayoutCachedWidth =11958
                    LayoutCachedHeight =480
                    ColumnEnd =1
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =2520
                    Top =240
                    Width =1815
                    Height =240
                    FontSize =8
                    BackColor =12349952
                    Name ="Label3428"
                    Caption =" MSO"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2520
                    LayoutCachedTop =240
                    LayoutCachedWidth =4335
                    LayoutCachedHeight =480
                    ColumnEnd =1
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =12180
                    Top =240
                    Width =2010
                    Height =240
                    FontSize =8
                    BackColor =12349952
                    Name ="Label3491"
                    Caption =" Title Info"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12180
                    LayoutCachedTop =240
                    LayoutCachedWidth =14190
                    LayoutCachedHeight =480
                    ColumnEnd =1
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15330
                    Top =570
                    Width =1335
                    Height =270
                    FontSize =10
                    TabIndex =6
                    Name ="UnitIFDate"
                    ControlSource ="UnitIFDate"
                    Format ="Short Date"
                    StatusBarText ="date unit iterim funded"
                    GroupTable =66

                    LayoutCachedLeft =15330
                    LayoutCachedTop =570
                    LayoutCachedWidth =16665
                    LayoutCachedHeight =840
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14385
                            Top =570
                            Width =885
                            Height =270
                            FontSize =8
                            Name ="Label356"
                            Caption ="Fund Date"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LayoutCachedLeft =14385
                            LayoutCachedTop =570
                            LayoutCachedWidth =15270
                            LayoutCachedHeight =840
                            ColumnStart =19
                            ColumnEnd =19
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =6840
                    Left =15330
                    Top =900
                    Width =1335
                    FontSize =10
                    TabIndex =13
                    Name ="UnitIFBank"
                    ControlSource ="UnitIFBank"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Clients.ClientShNm, Clients.clientCompanyName FROM Clients WHERE (((Clien"
                        "ts.clientType)=\"bank\")) ORDER BY Clients.ClientShNm;"
                    ColumnWidths ="1800;5040"
                    StatusBarText ="bank used"
                    GroupTable =66
                    AllowValueListEdits =0

                    LayoutCachedLeft =15330
                    LayoutCachedTop =900
                    LayoutCachedWidth =16665
                    LayoutCachedHeight =1140
                    RowStart =1
                    RowEnd =1
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14385
                            Top =900
                            Width =885
                            Height =240
                            FontSize =8
                            Name ="Label357"
                            Caption ="Loan Bank"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LayoutCachedLeft =14385
                            LayoutCachedTop =900
                            LayoutCachedWidth =15270
                            LayoutCachedHeight =1140
                            RowStart =1
                            RowEnd =1
                            ColumnStart =19
                            ColumnEnd =19
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =15330
                    Top =1200
                    Width =1335
                    FontSize =10
                    TabIndex =19
                    Name ="UnitIFAmt"
                    ControlSource ="UnitIFAmt"
                    Format ="$#,##0.00;($#,##0.00)"
                    ValidationRule =">=0"
                    ValidationText ="Interim Fund amount must be >=0"
                    DefaultValue ="0"
                    GroupTable =66

                    LayoutCachedLeft =15330
                    LayoutCachedTop =1200
                    LayoutCachedWidth =16665
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14385
                            Top =1200
                            Width =885
                            Height =240
                            FontSize =8
                            Name ="Label358"
                            Caption ="Amount "
                            FontName ="Segoe UI"
                            GroupTable =66
                            LayoutCachedLeft =14385
                            LayoutCachedTop =1200
                            LayoutCachedWidth =15270
                            LayoutCachedHeight =1440
                            RowStart =2
                            RowEnd =2
                            ColumnStart =19
                            ColumnEnd =19
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =16770
                    Top =1155
                    Width =735
                    Height =270
                    FontSize =10
                    TabIndex =27
                    Name ="Text361"
                    ControlSource ="=IIf(IsNull([unitpayoff]),Null,DateDiff(\"d\",[unitifdate],[unitpayoff]))"
                    Format ="Fixed"

                    LayoutCachedLeft =16770
                    LayoutCachedTop =1155
                    LayoutCachedWidth =17505
                    LayoutCachedHeight =1425
                    RowStart =3
                    RowEnd =3
                    ColumnStart =20
                    ColumnEnd =20
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextAlign =1
                            Left =16770
                            Top =900
                            Width =735
                            Height =270
                            FontSize =8
                            Name ="Label362"
                            Caption ="Duration "
                            FontName ="Segoe UI"
                            LayoutCachedLeft =16770
                            LayoutCachedTop =900
                            LayoutCachedWidth =17505
                            LayoutCachedHeight =1170
                            RowStart =3
                            RowEnd =3
                            ColumnStart =19
                            ColumnEnd =19
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =15330
                    Top =1500
                    Width =1335
                    FontSize =10
                    TabIndex =24
                    Name ="Text359"
                    ControlSource ="UnitPayOff"
                    Format ="Short Date"
                    GroupTable =66

                    LayoutCachedLeft =15330
                    LayoutCachedTop =1500
                    LayoutCachedWidth =16665
                    LayoutCachedHeight =1740
                    RowStart =3
                    RowEnd =3
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =14385
                            Top =1500
                            Width =885
                            Height =240
                            FontSize =8
                            Name ="Label360"
                            Caption ="Paid Off "
                            FontName ="Segoe UI"
                            GroupTable =66
                            LayoutCachedLeft =14385
                            LayoutCachedTop =1500
                            LayoutCachedWidth =15270
                            LayoutCachedHeight =1740
                            RowStart =3
                            RowEnd =3
                            ColumnStart =19
                            ColumnEnd =19
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin EmptyCell
                    Left =14250
                    Top =570
                    Width =75
                    Height =270
                    Name ="EmptyCell3610"
                    GroupTable =66
                    LayoutCachedLeft =14250
                    LayoutCachedTop =570
                    LayoutCachedWidth =14325
                    LayoutCachedHeight =840
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =14250
                    Top =900
                    Width =75
                    Name ="EmptyCell3611"
                    GroupTable =66
                    LayoutCachedLeft =14250
                    LayoutCachedTop =900
                    LayoutCachedWidth =14325
                    LayoutCachedHeight =1140
                    RowStart =1
                    RowEnd =1
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =14250
                    Top =1200
                    Width =75
                    Name ="EmptyCell3612"
                    GroupTable =66
                    LayoutCachedLeft =14250
                    LayoutCachedTop =1200
                    LayoutCachedWidth =14325
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =14250
                    Top =1500
                    Width =75
                    Name ="EmptyCell3613"
                    GroupTable =66
                    LayoutCachedLeft =14250
                    LayoutCachedTop =1500
                    LayoutCachedWidth =14325
                    LayoutCachedHeight =1740
                    RowStart =3
                    RowEnd =3
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =6
                    ListWidth =5760
                    Left =5070
                    Top =1500
                    Width =3360
                    ColumnOrder =19
                    FontSize =9
                    TabIndex =21
                    BackColor =15592953
                    Name ="CofaSign"
                    ControlSource ="CofASigner"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Contacts.contactID, Contacts.contactFullName, Contacts.contactTitle, Cont"
                        "acts.contactInc, Contacts.contactCofA FROM Contacts WHERE (((Contacts.clientRef)"
                        "=[forms]![UnitsForm]![clientid]) AND ((Contacts.ContactObsolete)=No));"
                    ColumnWidths ="0;1728;2880;576;576"
                    OnGotFocus ="[Event Procedure]"
                    GroupTable =66

                    LayoutCachedLeft =5070
                    LayoutCachedTop =1500
                    LayoutCachedWidth =8430
                    LayoutCachedHeight =1740
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =10
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =4485
                            Top =1500
                            Width =555
                            Height =240
                            FontSize =8
                            FontWeight =700
                            Name ="Label215"
                            Caption ="Signer "
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =4485
                            LayoutCachedTop =1500
                            LayoutCachedWidth =5040
                            LayoutCachedHeight =1740
                            RowStart =3
                            RowEnd =3
                            ColumnStart =6
                            ColumnEnd =6
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =14370
                    Top =240
                    Width =2310
                    Height =240
                    FontSize =10
                    BackColor =32768
                    Name ="Label3623"
                    Caption =" Interim Funding"
                    LayoutCachedLeft =14370
                    LayoutCachedTop =240
                    LayoutCachedWidth =16680
                    LayoutCachedHeight =480
                    BorderThemeColorIndex =6
                    BorderShade =75.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin EmptyCell
                    Left =4320
                    Top =570
                    Width =165
                    Height =870
                    Name ="EmptyCell3679"
                    GroupTable =66
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4320
                    LayoutCachedTop =570
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =1440
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =4320
                    Top =1500
                    Width =165
                    Name ="EmptyCell3680"
                    GroupTable =66
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4320
                    LayoutCachedTop =1500
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =1740
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =4380
                    Top =2475
                    Width =30
                    Height =255
                    Name ="EmptyCell3837"
                    GroupTable =106
                    LayoutCachedLeft =4380
                    LayoutCachedTop =2475
                    LayoutCachedWidth =4410
                    LayoutCachedHeight =2730
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =4380
                    Top =2790
                    Width =30
                    Height =255
                    Name ="EmptyCell3838"
                    GroupTable =106
                    LayoutCachedLeft =4380
                    LayoutCachedTop =2790
                    LayoutCachedWidth =4410
                    LayoutCachedHeight =3045
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =4380
                    Top =3105
                    Width =30
                    Height =255
                    Name ="EmptyCell3839"
                    GroupTable =106
                    LayoutCachedLeft =4380
                    LayoutCachedTop =3105
                    LayoutCachedWidth =4410
                    LayoutCachedHeight =3360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =4380
                    Top =3420
                    Width =30
                    Height =255
                    Name ="EmptyCell3840"
                    GroupTable =106
                    LayoutCachedLeft =4380
                    LayoutCachedTop =3420
                    LayoutCachedWidth =4410
                    LayoutCachedHeight =3675
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =8520
                    Top =2475
                    Width =105
                    Height =255
                    Name ="EmptyCell3842"
                    GroupTable =106
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8520
                    LayoutCachedTop =2475
                    LayoutCachedWidth =8625
                    LayoutCachedHeight =2730
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =8520
                    Top =2790
                    Width =105
                    Height =255
                    Name ="EmptyCell3843"
                    GroupTable =106
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8520
                    LayoutCachedTop =2790
                    LayoutCachedWidth =8625
                    LayoutCachedHeight =3045
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =8520
                    Top =3105
                    Width =105
                    Height =255
                    Name ="EmptyCell3844"
                    GroupTable =106
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8520
                    LayoutCachedTop =3105
                    LayoutCachedWidth =8625
                    LayoutCachedHeight =3360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =8520
                    Top =3420
                    Width =105
                    Height =255
                    Name ="EmptyCell3845"
                    GroupTable =106
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =8520
                    LayoutCachedTop =3420
                    LayoutCachedWidth =8625
                    LayoutCachedHeight =3675
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =8490
                    Top =1200
                    Width =120
                    Name ="EmptyCell3855"
                    GroupTable =66
                    LayoutCachedLeft =8490
                    LayoutCachedTop =1200
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =2385
                    Top =1200
                    Width =105
                    Name ="EmptyCell3858"
                    GroupTable =66
                    LayoutCachedLeft =2385
                    LayoutCachedTop =1200
                    LayoutCachedWidth =2490
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =255
                    Left =19695
                    Top =10200
                    Width =1485
                    Height =285
                    ColumnOrder =8
                    TabIndex =74
                    Name ="EXPortal"
                    ControlSource ="UnitExckPortRpt"
                    Tag ="ITTOnly"
                    GroupTable =114
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19695
                    LayoutCachedTop =10200
                    LayoutCachedWidth =21180
                    LayoutCachedHeight =10485
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =11
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =114
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =18330
                            Top =10200
                            Width =1335
                            Height =285
                            FontSize =9
                            FontWeight =700
                            ForeColor =8210719
                            Name ="Label470"
                            Caption ="No Portal"
                            Tag ="ITTOnly"
                            GroupTable =114
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18330
                            LayoutCachedTop =10200
                            LayoutCachedWidth =19665
                            LayoutCachedHeight =10485
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =11
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =114
                        End
                    End
                End
                Begin EmptyCell
                    Left =8490
                    Top =1500
                    Width =120
                    Name ="EmptyCell3971"
                    GroupTable =66
                    LayoutCachedLeft =8490
                    LayoutCachedTop =1500
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =1740
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =4320
                    Top =1800
                    Width =165
                    Name ="EmptyCell4024"
                    GroupTable =66
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4320
                    LayoutCachedTop =1800
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =6270
                    Top =1800
                    Width =105
                    Name ="EmptyCell4027"
                    GroupTable =66
                    LayoutCachedLeft =6270
                    LayoutCachedTop =1800
                    LayoutCachedWidth =6375
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =6405
                    Top =1800
                    Width =840
                    Name ="EmptyCell4028"
                    GroupTable =66
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =6405
                    LayoutCachedTop =1800
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =7275
                    Top =1800
                    Width =1155
                    Name ="EmptyCell4029"
                    GroupTable =66
                    LayoutCachedLeft =7275
                    LayoutCachedTop =1800
                    LayoutCachedWidth =8430
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =8490
                    Top =1800
                    Width =120
                    Name ="EmptyCell4030"
                    GroupTable =66
                    LayoutCachedLeft =8490
                    LayoutCachedTop =1800
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =6270
                    Top =1200
                    Width =105
                    Name ="EmptyCell4039"
                    GroupTable =66
                    LayoutCachedLeft =6270
                    LayoutCachedTop =1200
                    LayoutCachedWidth =6375
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =9960
                    Top =3420
                    Width =1215
                    Height =255
                    Name ="EmptyCell4121"
                    GroupTable =106
                    LayoutCachedLeft =9960
                    LayoutCachedTop =3420
                    LayoutCachedWidth =11175
                    LayoutCachedHeight =3675
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =8655
                    Top =3420
                    Width =1245
                    Height =255
                    Name ="EmptyCell4122"
                    GroupTable =106
                    LayoutCachedLeft =8655
                    LayoutCachedTop =3420
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =3675
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =4485
                    Top =1800
                    Width =555
                    Name ="EmptyCell4136"
                    GroupTable =66
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =4485
                    LayoutCachedTop =1800
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =5070
                    Top =1800
                    Width =1140
                    Name ="EmptyCell4137"
                    GroupTable =66
                    LayoutCachedLeft =5070
                    LayoutCachedTop =1800
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =6405
                    Top =1200
                    Width =840
                    Name ="EmptyCell4183"
                    GroupTable =66
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =6405
                    LayoutCachedTop =1200
                    LayoutCachedWidth =7245
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =7275
                    Top =1200
                    Width =1155
                    Name ="EmptyCell4184"
                    GroupTable =66
                    LayoutCachedLeft =7275
                    LayoutCachedTop =1200
                    LayoutCachedWidth =8430
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19695
                    Top =8910
                    Width =1485
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =69
                    Name ="ObcUnitNumber"
                    ControlSource ="ObcUnitNumber"
                    OnDblClick ="[Event Procedure]"
                    Tag ="ITTOnly"
                    GroupTable =114
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19695
                    LayoutCachedTop =8910
                    LayoutCachedWidth =21180
                    LayoutCachedHeight =9180
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =11
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =114
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =18330
                            Top =8910
                            Width =1335
                            Height =270
                            FontSize =8
                            Name ="Label4287"
                            Caption ="OBCNumber "
                            FontName ="Segoe UI"
                            Tag ="ITTOnly"
                            GroupTable =114
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18330
                            LayoutCachedTop =8910
                            LayoutCachedWidth =19665
                            LayoutCachedHeight =9180
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =11
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =114
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14985
                    Top =3090
                    Width =1125
                    FontSize =10
                    TabIndex =52
                    Name ="unitoffleasedt"
                    ControlSource ="unitoffleasedt"
                    Format ="Short Date"
                    StatusBarText ="off lease date (rent paid throuogh)"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =112

                    LayoutCachedLeft =14985
                    LayoutCachedTop =3090
                    LayoutCachedWidth =16110
                    LayoutCachedHeight =3330
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =9
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =112
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =13695
                            Top =3090
                            Width =1230
                            Height =240
                            FontSize =8
                            Name ="Label175"
                            Caption ="Actual OffLease "
                            FontName ="Segoe UI"
                            GroupTable =112
                            LayoutCachedLeft =13695
                            LayoutCachedTop =3090
                            LayoutCachedWidth =14925
                            LayoutCachedHeight =3330
                            RowStart =2
                            RowEnd =2
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =9
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =112
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12600
                    Top =2790
                    Width =1035
                    FontSize =10
                    TabIndex =49
                    BackColor =15523798
                    Name ="MTMRent"
                    ControlSource ="MTMRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="MTM- Month to Month extension unit rent"
                    ValidationRule =">=0"
                    ValidationText ="Rent fields must be 0 or greater"
                    DefaultValue ="0"
                    ControlTipText ="MTM- Rent  Month to Month extension unit rent"
                    GroupTable =112

                    LayoutCachedLeft =12600
                    LayoutCachedTop =2790
                    LayoutCachedWidth =13635
                    LayoutCachedHeight =3030
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =112
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =11340
                            Top =2790
                            Width =1200
                            Height =240
                            FontSize =8
                            Name ="M-T-MRentlbl"
                            Caption ="MTM Rent"
                            FontName ="Segoe UI"
                            EventProcPrefix ="M_T_MRentlbl"
                            GroupTable =112
                            LayoutCachedLeft =11340
                            LayoutCachedTop =2790
                            LayoutCachedWidth =12540
                            LayoutCachedHeight =3030
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =9
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =112
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12600
                    Top =2475
                    Width =1035
                    Height =255
                    FontSize =10
                    TabIndex =47
                    BackColor =15523798
                    Name ="FirstExtRent"
                    ControlSource ="FirstExtRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="1st Extension Rent - Extension rent for fixed period; not M-T-M see schedule gro"
                        "up record."
                    ValidationRule =">=0"
                    ValidationText ="Rent fields must be 0 or greater"
                    DefaultValue ="0"
                    ControlTipText ="1st Extension Rent - Extension rent for fixed period; not M-T-M see schedule gro"
                        "up record."
                    GroupTable =112

                    LayoutCachedLeft =12600
                    LayoutCachedTop =2475
                    LayoutCachedWidth =13635
                    LayoutCachedHeight =2730
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =112
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =11340
                            Top =2475
                            Width =1200
                            Height =255
                            FontSize =8
                            Name ="FirstExtRentlbl"
                            Caption ="Extension Rent"
                            FontName ="Segoe UI"
                            GroupTable =112
                            LayoutCachedLeft =11340
                            LayoutCachedTop =2475
                            LayoutCachedWidth =12540
                            LayoutCachedHeight =2730
                            LayoutGroup =9
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =112
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =11340
                    Top =2145
                    Width =5310
                    Height =240
                    FontSize =8
                    BackColor =12349952
                    Name ="Label4358"
                    Caption =" Lease End Options"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =11340
                    LayoutCachedTop =2145
                    LayoutCachedWidth =16650
                    LayoutCachedHeight =2385
                    ColumnEnd =1
                    BorderThemeColorIndex =0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1584
                    Left =12600
                    Top =3090
                    Width =1035
                    FontSize =10
                    TabIndex =51
                    BoundColumn =1
                    Name ="OffLeaseInd"
                    ControlSource ="UnitOLInd"
                    RowSourceType ="Value List"
                    RowSource ="Projected;P;Actual;A"
                    ColumnWidths ="1080;360"
                    ValidationRule ="=\"P\" Or \"A\""
                    ValidationText ="must be P or A"
                    OnGotFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    GroupTable =112
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =12600
                    LayoutCachedTop =3090
                    LayoutCachedWidth =13635
                    LayoutCachedHeight =3330
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =9
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =112
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            TextAlign =1
                            Left =11340
                            Top =3090
                            Width =1200
                            Height =240
                            FontSize =8
                            Name ="Label4375"
                            Caption ="Off Lease Ind  "
                            FontName ="Segoe UI"
                            GroupTable =112
                            LayoutCachedLeft =11340
                            LayoutCachedTop =3090
                            LayoutCachedWidth =12540
                            LayoutCachedHeight =3330
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =9
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =112
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19695
                    Top =9690
                    Width =1485
                    FontSize =10
                    FontWeight =700
                    TabIndex =72
                    Name ="ReportEndDate"
                    ControlSource ="ReportEndDate"
                    Format ="Short Date"
                    StatusBarText ="off lease date (rent paid throuogh)"
                    Tag ="ITTOnly"
                    GroupTable =114
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19695
                    LayoutCachedTop =9690
                    LayoutCachedWidth =21180
                    LayoutCachedHeight =9930
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =11
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =114
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =18330
                            Top =9690
                            Width =1335
                            Height =240
                            FontSize =8
                            Name ="Label4381"
                            Caption ="Report End"
                            FontName ="Segoe UI"
                            Tag ="ITTOnly"
                            GroupTable =114
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18330
                            LayoutCachedTop =9690
                            LayoutCachedWidth =19665
                            LayoutCachedHeight =9930
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =11
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =114
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19695
                    Top =9450
                    Width =1485
                    FontSize =10
                    FontWeight =700
                    TabIndex =71
                    Name ="ReportStartDate"
                    ControlSource ="ReportStartDate"
                    Format ="Short Date"
                    StatusBarText ="off lease date (rent paid throuogh)"
                    Tag ="ITTOnly"
                    GroupTable =114
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19695
                    LayoutCachedTop =9450
                    LayoutCachedWidth =21180
                    LayoutCachedHeight =9690
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =11
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =114
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =18330
                            Top =9450
                            Width =1335
                            Height =240
                            FontSize =8
                            Name ="Label4383"
                            Caption ="Report Start"
                            FontName ="Segoe UI"
                            Tag ="ITTOnly"
                            GroupTable =114
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18330
                            LayoutCachedTop =9450
                            LayoutCachedWidth =19665
                            LayoutCachedHeight =9690
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =11
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =114
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19560
                    Top =6465
                    Width =750
                    ColumnOrder =10
                    FontSize =8
                    TabIndex =61
                    BorderColor =0
                    ForeColor =5026082
                    Name ="SchID"
                    ControlSource ="SchID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19560
                    LayoutCachedTop =6465
                    LayoutCachedWidth =20310
                    LayoutCachedHeight =6705
                    RowStart =3
                    RowEnd =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =18720
                    Top =2640
                    Width =2415
                    Height =210
                    FontSize =8
                    FontWeight =700
                    TabIndex =32
                    Name ="ClearFilters"
                    Caption ="Clear Filter"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    GroupTable =12
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =18720
                    LayoutCachedTop =2640
                    LayoutCachedWidth =21135
                    LayoutCachedHeight =2850
                    PictureCaptionArrangement =5
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =4
                    ForeShade =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =10855845
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =1
                    PressedShade =75.0
                    HoverForeThemeColorIndex =4
                    HoverForeShade =75.0
                    PressedForeThemeColorIndex =4
                    PressedForeShade =75.0
                    GroupTable =12
                    QuickStyle =22
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =18720
                    Top =2340
                    Width =435
                    Height =246
                    FontSize =8
                    FontWeight =700
                    TabIndex =28
                    Name ="FirstRecord"
                    Caption ="FirstRecord"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddddddd44dddddcdddddddd44ddddccdddddddd44dddcccdddd ,
                        0xdddd44ddccccdddddddd44dcccccdddddddd44ddccccdddddddd44dddcccdddd ,
                        0xdddd44ddddccdddddddd44dddddcdddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddd000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="First Record"
                    GroupTable =12
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="GoToRecord"
                            Argument ="-1"
                            Argument =""
                            Argument ="2"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"FirstRecord\" xmlns=\"http://schemas.microsoft.com/office/ac"
                                "cessservices/2009/11/application\"><Statements><Action Name=\"GoToRecord\"><Argu"
                                "ment Name=\"Record\">First</Argum"
                        End
                        Begin
                            Comment ="_AXL:ent></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =18720
                    LayoutCachedTop =2340
                    LayoutCachedWidth =19155
                    LayoutCachedHeight =2586
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =10855845
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =1
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =12
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =19215
                    Top =2340
                    Width =435
                    Height =246
                    FontSize =8
                    FontWeight =700
                    TabIndex =29
                    Name ="PreviousRecord"
                    Caption ="PreviousRecord"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddddddddddddcddddddddddddddccdddddddddddddcccdddddd ,
                        0xddddddccccdddddddddddcccccddddddddddddccccdddddddddddddcccdddddd ,
                        0xddddddddccdddddddddddddddcdddddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddd000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Previous Record"
                    GroupTable =12
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Action ="OnError"
                            Argument ="0"
                        End
                        Begin
                            Action ="GoToRecord"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Condition ="[MacroError]<>0"
                            Action ="MsgBox"
                            Argument ="=[MacroError].[Description]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"PreviousRecord\" xmlns=\"http://schemas.microsoft.com/office"
                                "/accessservices/2009/11/application\"><Statements><Action Name=\"OnError\"/><Act"
                                "ion Name=\"GoToRecord\"><Argument"
                        End
                        Begin
                            Comment ="_AXL: Name=\"Record\">Previous</Argument></Action><ConditionalBlock><If><Conditi"
                                "on>[MacroError]&lt;&gt;0</Condition><Statements><Action Name=\"MessageBox\"><Arg"
                                "ument Name=\"Message\">=[MacroError].[Description]</Argument></Action></Statemen"
                                "ts></If></Condition"
                        End
                        Begin
                            Comment ="_AXL:alBlock></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =19215
                    LayoutCachedTop =2340
                    LayoutCachedWidth =19650
                    LayoutCachedHeight =2586
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =10855845
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =1
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =12
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =20205
                    Top =2340
                    Width =435
                    Height =246
                    FontSize =8
                    FontWeight =700
                    TabIndex =30
                    Name ="NextRecord"
                    Caption ="NextRecord"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xddddddddddddddddddddddcdddddddddddddddccddddddddddddddcccddddddd ,
                        0xddddddccccddddddddddddcccccdddddddddddccccddddddddddddcccddddddd ,
                        0xddddddccddddddddddddddcddddddddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddd000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Next Record"
                    GroupTable =12
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Action ="OnError"
                            Argument ="0"
                        End
                        Begin
                            Action ="GoToRecord"
                            Argument ="-1"
                            Argument =""
                            Argument ="1"
                        End
                        Begin
                            Condition ="[MacroError]<>0"
                            Action ="MsgBox"
                            Argument ="=[MacroError].[Description]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"NextRecord\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"OnError\"/><Action "
                                "Name=\"GoToRecord\"/><Conditional"
                        End
                        Begin
                            Comment ="_AXL:Block><If><Condition>[MacroError]&lt;&gt;0</Condition><Statements><Action N"
                                "ame=\"MessageBox\"><Argument Name=\"Message\">=[MacroError].[Description]</Argum"
                                "ent></Action></Statements></If></ConditionalBlock></Statements></UserInterfaceMa"
                                "cro>"
                        End
                    End

                    LayoutCachedLeft =20205
                    LayoutCachedTop =2340
                    LayoutCachedWidth =20640
                    LayoutCachedHeight =2586
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =10855845
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =1
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =12
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =20700
                    Top =2340
                    Width =435
                    Height =246
                    FontSize =8
                    FontWeight =700
                    TabIndex =31
                    Name ="LastRecord"
                    Caption ="LastRecord"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddddddddcddddd44ddddddddccdddd44ddddddddcccddd44ddd ,
                        0xdddddccccdd44ddddddddcccccd44ddddddddccccdd44ddddddddcccddd44ddd ,
                        0xdddddccdddd44ddddddddcddddd44ddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddd000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Last Record"
                    GroupTable =12
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="GoToRecord"
                            Argument ="-1"
                            Argument =""
                            Argument ="3"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"LastRecord\" xmlns=\"http://schemas.microsoft.com/office/acc"
                                "essservices/2009/11/application\"><Statements><Action Name=\"GoToRecord\"><Argum"
                                "ent Name=\"Record\">Last</Argumen"
                        End
                        Begin
                            Comment ="_AXL:t></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =20700
                    LayoutCachedTop =2340
                    LayoutCachedWidth =21135
                    LayoutCachedHeight =2586
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =10855845
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =1
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =12
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin EmptyCell
                    Left =19710
                    Top =2340
                    Width =435
                    Height =246
                    Name ="EmptyCell4460"
                    GroupTable =12
                    LayoutCachedLeft =19710
                    LayoutCachedTop =2340
                    LayoutCachedWidth =20145
                    LayoutCachedHeight =2586
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =12
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19560
                    Top =7500
                    Width =750
                    FontSize =8
                    TabIndex =65
                    BorderColor =0
                    ForeColor =5026082
                    Name ="UnitsClientID"
                    ControlSource ="Units.clientID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19560
                    LayoutCachedTop =7500
                    LayoutCachedWidth =20310
                    LayoutCachedHeight =7740
                    RowStart =7
                    RowEnd =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2385
                    Top =3105
                    Width =1935
                    Height =255
                    FontSize =10
                    TabIndex =41
                    Name ="AcqHours"
                    ControlSource ="AcqHours"
                    Format ="Standard"
                    ControlTipText ="Hours when the unit was acquired"
                    GroupTable =106

                    LayoutCachedLeft =2385
                    LayoutCachedTop =3105
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =3360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =180
                            Top =3105
                            Width =2145
                            Height =255
                            FontSize =8
                            Name ="Label4625"
                            Caption ="Hours @ acquisition"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =180
                            LayoutCachedTop =3105
                            LayoutCachedWidth =2325
                            LayoutCachedHeight =3360
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =7
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =10
                    ListWidth =2448
                    Left =18435
                    Top =11190
                    Width =2460
                    Height =300
                    FontSize =10
                    TabIndex =75
                    BackColor =15523798
                    Name ="DivisionOwnerID"
                    ControlSource ="DivisionOwnerID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ClientDivisions.DivisionID, ClientDivisions.DivisionName, ClientDivisions"
                        ".ClientID FROM ClientDivisions WHERE (((ClientDivisions.ClientID)=[Forms]![Units"
                        "Form].[clientID]));"
                    ColumnWidths ="0;2304"
                    Tag ="Cloned"
                    ControlTipText ="The Division or Owership Group This Unit belongs to."

                    LayoutCachedLeft =18435
                    LayoutCachedTop =11190
                    LayoutCachedWidth =20895
                    LayoutCachedHeight =11490
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =18435
                            Top =10920
                            Width =1725
                            Height =240
                            FontSize =8
                            Name ="Label4682"
                            Caption ="Division"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =18435
                            LayoutCachedTop =10920
                            LayoutCachedWidth =20160
                            LayoutCachedHeight =11160
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19560
                    Top =7740
                    Width =750
                    FontSize =8
                    TabIndex =66
                    BorderColor =0
                    ForeColor =5026082
                    Name ="ATFUnitSpecID"
                    ControlSource ="ATFUnitSpecID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19560
                    LayoutCachedTop =7740
                    LayoutCachedWidth =20310
                    LayoutCachedHeight =7980
                    RowStart =8
                    RowEnd =8
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19560
                    Top =5745
                    Width =750
                    FontSize =8
                    TabIndex =58
                    BorderColor =0
                    ForeColor =5026082
                    Name ="ATFGroupSpecID"
                    ControlSource ="ATFGroupSpecID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19560
                    LayoutCachedTop =5745
                    LayoutCachedWidth =20310
                    LayoutCachedHeight =5985
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19695
                    Top =9930
                    Width =1485
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =73
                    Name ="UnitPlateState"
                    ControlSource ="UnitPlateState"
                    Tag ="ITTOnly"
                    GroupTable =114
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19695
                    LayoutCachedTop =9930
                    LayoutCachedWidth =21180
                    LayoutCachedHeight =10200
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =11
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =114
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =18330
                            Top =9930
                            Width =1335
                            Height =270
                            FontSize =8
                            Name ="Label4686"
                            Caption ="PlateState "
                            FontName ="Segoe UI"
                            Tag ="ITTOnly"
                            GroupTable =114
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18330
                            LayoutCachedTop =9930
                            LayoutCachedWidth =19665
                            LayoutCachedHeight =10200
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =11
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =114
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19695
                    Top =8370
                    Width =1485
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =67
                    Name ="UnitClientFleetInfo"
                    ControlSource ="UnitClientFleetInfo"
                    Tag ="ITTOnly"
                    GroupTable =114
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19695
                    LayoutCachedTop =8370
                    LayoutCachedWidth =21180
                    LayoutCachedHeight =8640
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =11
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =114
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =18330
                            Top =8370
                            Width =1335
                            Height =270
                            FontSize =8
                            Name ="Label4692"
                            Caption ="Client Fleet Info"
                            FontName ="Segoe UI"
                            Tag ="ITTOnly"
                            GroupTable =114
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18330
                            LayoutCachedTop =8370
                            LayoutCachedWidth =19665
                            LayoutCachedHeight =8640
                            LayoutGroup =11
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =114
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19695
                    Top =8640
                    Width =1485
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =68
                    Name ="TUnitLessorOwner"
                    ControlSource ="TUnitLessorOwner"
                    Tag ="ITTOnly"
                    GroupTable =114
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19695
                    LayoutCachedTop =8640
                    LayoutCachedWidth =21180
                    LayoutCachedHeight =8910
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =11
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =114
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =18330
                            Top =8640
                            Width =1335
                            Height =270
                            FontSize =8
                            Name ="Label4698"
                            Caption ="Lessor Owner"
                            FontName ="Segoe UI"
                            Tag ="ITTOnly"
                            GroupTable =114
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18330
                            LayoutCachedTop =8640
                            LayoutCachedWidth =19665
                            LayoutCachedHeight =8910
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =11
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =114
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2385
                    Top =3420
                    Width =1935
                    Height =255
                    FontSize =10
                    TabIndex =45
                    Name ="PLBReadDate"
                    ControlSource ="PLBReadDate"
                    Format ="Short Date"
                    ControlTipText ="DAte the HMiles or Hours were read"
                    GroupTable =106

                    LayoutCachedLeft =2385
                    LayoutCachedTop =3420
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =3675
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =180
                            Top =3420
                            Width =2145
                            Height =255
                            FontSize =8
                            Name ="Label4740"
                            Caption ="  Read Date"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =180
                            LayoutCachedTop =3420
                            LayoutCachedWidth =2325
                            LayoutCachedHeight =3675
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =7
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2385
                    Top =2475
                    Width =1935
                    Height =255
                    FontSize =10
                    TabIndex =33
                    Name ="ClientISD"
                    ControlSource ="ClientISD"
                    Format ="Short Date"
                    ControlTipText ="NOT FA's In Service Date, it is the date the Client put the unit in service"
                    GroupTable =106

                    LayoutCachedLeft =2385
                    LayoutCachedTop =2475
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =2730
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =7
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =180
                            Top =2475
                            Width =2145
                            Height =255
                            FontSize =8
                            Name ="Label4762"
                            Caption ="Client's  In Service Date"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =180
                            LayoutCachedTop =2475
                            LayoutCachedWidth =2325
                            LayoutCachedHeight =2730
                            LayoutGroup =7
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14985
                    Top =2790
                    Width =1125
                    FontSize =10
                    TabIndex =50
                    Name ="ProjOffLease"
                    ControlSource ="ProjOffLease"
                    Format ="Short Date"
                    StatusBarText ="off lease date (rent paid throuogh)"
                    GroupTable =112

                    LayoutCachedLeft =14985
                    LayoutCachedTop =2790
                    LayoutCachedWidth =16110
                    LayoutCachedHeight =3030
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =112
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =13695
                            Top =2790
                            Width =1230
                            Height =240
                            FontSize =8
                            Name ="Label4791"
                            Caption ="Proj OffLease "
                            FontName ="Segoe UI"
                            GroupTable =112
                            LayoutCachedLeft =13695
                            LayoutCachedTop =2790
                            LayoutCachedWidth =14925
                            LayoutCachedHeight =3030
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =9
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =112
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14985
                    Top =2475
                    Width =1125
                    Height =255
                    FontSize =10
                    TabIndex =48
                    Name ="Text4796"
                    ControlSource ="LeaseExp"
                    Format ="Short Date"
                    StatusBarText ="off lease date (rent paid throuogh)"
                    GroupTable =112

                    LayoutCachedLeft =14985
                    LayoutCachedTop =2475
                    LayoutCachedWidth =16110
                    LayoutCachedHeight =2730
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =112
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =13695
                            Top =2475
                            Width =1230
                            Height =255
                            FontSize =8
                            Name ="Label4797"
                            Caption ="Lease Exp "
                            FontName ="Segoe UI"
                            GroupTable =112
                            LayoutCachedLeft =13695
                            LayoutCachedTop =2475
                            LayoutCachedWidth =14925
                            LayoutCachedHeight =2730
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =9
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =112
                        End
                    End
                End
                Begin EmptyCell
                    Left =11235
                    Top =2475
                    Height =255
                    Name ="EmptyCell4802"
                    GroupTable =106
                    LayoutCachedLeft =11235
                    LayoutCachedTop =2475
                    LayoutCachedWidth =12675
                    LayoutCachedHeight =2730
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =11235
                    Top =2790
                    Height =255
                    Name ="EmptyCell4803"
                    GroupTable =106
                    LayoutCachedLeft =11235
                    LayoutCachedTop =2790
                    LayoutCachedWidth =12675
                    LayoutCachedHeight =3045
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =11235
                    Top =3420
                    Height =255
                    Name ="EmptyCell4805"
                    GroupTable =106
                    LayoutCachedLeft =11235
                    LayoutCachedTop =3420
                    LayoutCachedWidth =12675
                    LayoutCachedHeight =3675
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =11235
                    Top =3105
                    Height =255
                    Name ="EmptyCell4811"
                    GroupTable =106
                    LayoutCachedLeft =11235
                    LayoutCachedTop =3105
                    LayoutCachedWidth =12675
                    LayoutCachedHeight =3360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1155
                    Top =1500
                    Width =1170
                    FontSize =10
                    FontWeight =700
                    TabIndex =20
                    BackColor =15523798
                    Name ="UnitDel"
                    ControlSource ="UnitDel"
                    RowSourceType ="Value List"
                    RowSource ="Road;Rail;LoadOut"
                    GroupTable =66
                    AllowValueListEdits =0

                    LayoutCachedLeft =1155
                    LayoutCachedTop =1500
                    LayoutCachedWidth =2325
                    LayoutCachedHeight =1740
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =3
                    BorderThemeColorIndex =4
                    BorderTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =150
                            Top =1500
                            Width =975
                            Height =240
                            FontSize =8
                            Name ="Label281"
                            Caption ="Method"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =150
                            LayoutCachedTop =1500
                            LayoutCachedWidth =1125
                            LayoutCachedHeight =1740
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =3
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin EmptyCell
                    Left =2520
                    Top =1500
                    Width =480
                    Name ="EmptyCell4828"
                    GroupTable =66
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =2520
                    LayoutCachedTop =1500
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =1740
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =2385
                    Top =1500
                    Width =105
                    Name ="EmptyCell4829"
                    GroupTable =66
                    LayoutCachedLeft =2385
                    LayoutCachedTop =1500
                    LayoutCachedWidth =2490
                    LayoutCachedHeight =1740
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =3000
                    Top =1500
                    Width =1320
                    Name ="EmptyCell4830"
                    GroupTable =66
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =3000
                    LayoutCachedTop =1500
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1740
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =150
                    Top =1800
                    Width =4170
                    Height =240
                    FontSize =8
                    Name ="Label4831"
                    Caption ="     Delivery Comments moved to Notes Tab"
                    FontName ="Segoe UI"
                    GroupTable =66
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =150
                    LayoutCachedTop =1800
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =4
                    LayoutGroup =3
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19695
                    Top =9180
                    Width =1485
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =70
                    Name ="Text4834"
                    ControlSource ="ClientIDOBC"
                    Tag ="ITTOnly"
                    GroupTable =114
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =19695
                    LayoutCachedTop =9180
                    LayoutCachedWidth =21180
                    LayoutCachedHeight =9450
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =11
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =114
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =18330
                            Top =9180
                            Width =1335
                            Height =270
                            FontSize =8
                            Name ="Label4835"
                            Caption ="ClientIDObc"
                            FontName ="Segoe UI"
                            Tag ="ITTOnly"
                            GroupTable =114
                            TopPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =18330
                            LayoutCachedTop =9180
                            LayoutCachedWidth =19665
                            LayoutCachedHeight =9450
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =11
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =114
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9960
                    Top =3105
                    Width =1215
                    Height =255
                    FontSize =10
                    TabIndex =44
                    BackColor =11193849
                    Name ="RVGap"
                    ControlSource ="RVGuar"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="GAP-Guarantee"
                    ValidationRule =">=0"
                    ValidationText ="Syn Resid must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Sets the RV GAP-Guarantee"
                    GroupTable =106

                    LayoutCachedLeft =9960
                    LayoutCachedTop =3105
                    LayoutCachedWidth =11175
                    LayoutCachedHeight =3360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =7
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =8655
                            Top =3105
                            Width =1245
                            Height =255
                            FontSize =8
                            BackColor =13952764
                            Name ="Label4895"
                            Caption ="RV Guarantee"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =8655
                            LayoutCachedTop =3105
                            LayoutCachedWidth =9900
                            LayoutCachedHeight =3360
                            RowStart =2
                            RowEnd =2
                            ColumnStart =9
                            ColumnEnd =9
                            LayoutGroup =7
                            BorderThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9690
                    Top =1500
                    Width =2248
                    FontSize =10
                    TabIndex =22
                    Name ="Text4918"
                    ControlSource ="VendorCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =66

                    LayoutCachedLeft =9690
                    LayoutCachedTop =1500
                    LayoutCachedWidth =11938
                    LayoutCachedHeight =1740
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =14
                    LayoutGroup =3
                    BackThemeColorIndex =4
                    BackTint =40.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =66
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =2
                            Left =8640
                            Top =1500
                            Width =1027
                            Height =240
                            FontSize =8
                            Name ="Label4919"
                            Caption ="Vendor Cost"
                            FontName ="Segoe UI"
                            GroupTable =66
                            LeftPadding =0
                            RightPadding =0
                            LayoutCachedLeft =8640
                            LayoutCachedTop =1500
                            LayoutCachedWidth =9667
                            LayoutCachedHeight =1740
                            RowStart =3
                            RowEnd =3
                            ColumnStart =12
                            ColumnEnd =12
                            LayoutGroup =3
                            BackThemeColorIndex =4
                            BackTint =20.0
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =66
                        End
                    End
                End
                Begin EmptyCell
                    Left =12000
                    Top =1800
                    Width =105
                    Name ="EmptyCell4927"
                    GroupTable =66
                    LayoutCachedLeft =12000
                    LayoutCachedTop =1800
                    LayoutCachedWidth =12105
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =14250
                    Top =1800
                    Width =75
                    Name ="EmptyCell4929"
                    GroupTable =66
                    LayoutCachedLeft =14250
                    LayoutCachedTop =1800
                    LayoutCachedWidth =14325
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =14385
                    Top =1800
                    Width =885
                    Name ="EmptyCell4931"
                    GroupTable =66
                    LayoutCachedLeft =14385
                    LayoutCachedTop =1800
                    LayoutCachedWidth =15270
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =19
                    ColumnEnd =19
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =15330
                    Top =1800
                    Width =1335
                    Name ="EmptyCell4939"
                    GroupTable =66
                    LayoutCachedLeft =15330
                    LayoutCachedTop =1800
                    LayoutCachedWidth =16665
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =20
                    ColumnEnd =20
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =12135
                    Top =1800
                    Width =885
                    Name ="EmptyCell4954"
                    GroupTable =66
                    LeftPadding =0
                    RightPadding =0
                    LayoutCachedLeft =12135
                    LayoutCachedTop =1800
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =13050
                    Top =1800
                    Width =1140
                    Name ="EmptyCell4955"
                    GroupTable =66
                    LayoutCachedLeft =13050
                    LayoutCachedTop =1800
                    LayoutCachedWidth =14190
                    LayoutCachedHeight =2040
                    RowStart =4
                    RowEnd =4
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =10875
                    Top =1200
                    Width =1063
                    Name ="EmptyCell4982"
                    GroupTable =66
                    LayoutCachedLeft =10875
                    LayoutCachedTop =1200
                    LayoutCachedWidth =11938
                    LayoutCachedHeight =1440
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin EmptyCell
                    Left =10875
                    Top =900
                    Width =1063
                    Name ="EmptyCell5045"
                    GroupTable =66
                    LayoutCachedLeft =10875
                    LayoutCachedTop =900
                    LayoutCachedWidth =11938
                    LayoutCachedHeight =1140
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =3
                    GroupTable =66
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5235
                    Top =3105
                    Width =1080
                    Height =255
                    FontSize =10
                    TabIndex =42
                    BackColor =11193849
                    Name ="PerDiemSurrender"
                    ControlSource ="PerDiemSurrender"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Surrender Per Diem may be Sch Rent / 30"
                    ValidationRule =">=0"
                    ValidationText ="PD Rent must be greater than of = to 0"
                    DefaultValue ="0"
                    ControlTipText ="Surrender Per Diem may be Sch Rent / 30"
                    GroupTable =106

                    LayoutCachedLeft =5235
                    LayoutCachedTop =3105
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =3360
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =7
                    BorderThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =106
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =4470
                            Top =3105
                            Width =705
                            Height =255
                            FontSize =8
                            BackColor =13952764
                            Name ="Label5050"
                            Caption ="Surrendr"
                            FontName ="Segoe UI"
                            GroupTable =106
                            LayoutCachedLeft =4470
                            LayoutCachedTop =3105
                            LayoutCachedWidth =5175
                            LayoutCachedHeight =3360
                            RowStart =2
                            RowEnd =2
                            ColumnStart =3
                            ColumnEnd =3
                            LayoutGroup =7
                            BorderThemeColorIndex =2
                            ForeThemeColorIndex =0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =106
                        End
                    End
                End
                Begin EmptyCell
                    Left =4470
                    Top =3420
                    Width =705
                    Height =255
                    Name ="EmptyCell5053"
                    GroupTable =106
                    LayoutCachedLeft =4470
                    LayoutCachedTop =3420
                    LayoutCachedWidth =5175
                    LayoutCachedHeight =3675
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =7
                    GroupTable =106
                End
                Begin EmptyCell
                    Left =5235
                    Top =3420
                    Width =1080
                    Height =255
                    Name ="EmptyCell5054"
                    GroupTable =106
                    LayoutCachedLeft =5235
                    LayoutCachedTop =3420
                    LayoutCachedWidth =6315
                    LayoutCachedHeight =3675
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =7
                    GroupTable =106
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =15711498
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "UnitsFormBackup.cls"
