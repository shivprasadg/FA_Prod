Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =24396
    DatasheetFontHeight =11
    ItemSuffix =616
    Left =25920
    Top =16695
    Right =-16981
    Bottom =29400
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xe7112cfd6f87e640
    End
    RecordSource ="Select UnitUnitNum, UnitVIN, * From ChildParts as CP Inner Join Bulk_Update_Work"
        "Table_Child As CWT On CWT.ChildPartID= CP.ChildPartID;"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    OnDataSetChange ="[Event Procedure]"
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
            Height =1455
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =10380
                    Top =1050
                    Width =1470
                    Height =285
                    FontSize =9
                    Name ="Label600"
                    Caption ="Desc"
                    RightPadding =38
                    LayoutCachedLeft =10380
                    LayoutCachedTop =1050
                    LayoutCachedWidth =11850
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =12525
                    Width =765
                    Height =210
                    ColumnOrder =5
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedLeft =12525
                    LayoutCachedWidth =13290
                    LayoutCachedHeight =210
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =238
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =840
                    Top =60
                    Width =3888
                    Height =1185
                    ColumnOrder =4
                    FontSize =8
                    TabIndex =1
                    BoundColumn =3
                    ForeColor =-2147483617
                    BorderColor =14211288
                    Name ="ChildList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Bulk_Update_SelectedUnits_Child.ChildMake AS Make, Bulk_Update_SelectedUn"
                        "its_Child.ChildClass AS Class, Bulk_Update_SelectedUnits_Child.ChildType AS Type"
                        ", [ChildType] & [ChildVendorREF] & [ChildMake] & [ChildClass] AS UID FROM Bulk_U"
                        "pdate_SelectedUnits_Child GROUP BY Bulk_Update_SelectedUnits_Child.ChildMake, Bu"
                        "lk_Update_SelectedUnits_Child.ChildClass, Bulk_Update_SelectedUnits_Child.ChildT"
                        "ype, [ChildType] & [ChildVendorREF] & [ChildMake] & [ChildClass] ORDER BY Bulk_U"
                        "pdate_SelectedUnits_Child.ChildMake, Bulk_Update_SelectedUnits_Child.ChildType;"
                    ColumnWidths ="979;1440;1440;0"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =840
                    LayoutCachedTop =60
                    LayoutCachedWidth =4728
                    LayoutCachedHeight =1245
                    BackThemeColorIndex =1
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =6240
                    Top =60
                    Height =255
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="MakeExcelExport"
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

                    LayoutCachedLeft =6240
                    LayoutCachedTop =60
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =315
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
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =18870
                    Top =180
                    Width =3495
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="Label520"
                    Caption ="┌────── Interim Funding ────────────────┐"
                    LayoutCachedLeft =18870
                    LayoutCachedTop =180
                    LayoutCachedWidth =22365
                    LayoutCachedHeight =450
                    BackThemeColorIndex =8
                    BackShade =75.0
                    ForeThemeColorIndex =1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =22710
                    Top =180
                    Width =1350
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label511"
                    Caption ="┌ Show/Hide┐"
                    LayoutCachedLeft =22710
                    LayoutCachedTop =180
                    LayoutCachedWidth =24060
                    LayoutCachedHeight =480
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =13965
                    Top =195
                    Width =3255
                    Height =270
                    FontSize =9
                    FontWeight =700
                    ForeColor =8210719
                    Name ="Label985"
                    Caption ="┌─────────  Vendor Invoice  ──────────┐"
                    LayoutCachedLeft =13965
                    LayoutCachedTop =195
                    LayoutCachedWidth =17220
                    LayoutCachedHeight =465
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Top =60
                    Width =780
                    Height =1035
                    FontSize =16
                    FontWeight =700
                    Name ="Label114"
                    Caption ="Child\015\012Parts"
                    LayoutCachedTop =60
                    LayoutCachedWidth =780
                    LayoutCachedHeight =1095
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    Left =8070
                    Top =240
                    Width =1035
                    Height =285
                    FontSize =8
                    Name ="Label246"
                    Caption ="┌  Linked  ┐"
                    LayoutCachedLeft =8070
                    LayoutCachedTop =240
                    LayoutCachedWidth =9105
                    LayoutCachedHeight =525
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =5220
                    Top =60
                    Width =840
                    Height =270
                    FontSize =9
                    TabIndex =2
                    ForeColor =16777215
                    Name ="SYNCChild2Units"
                    Caption =" UNITS"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe MDL2 Assets"
                    ControlTipText ="Filter Unit Selection to these Child Parts"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5220
                    LayoutCachedTop =60
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =330
                    ColumnStart =5
                    ColumnEnd =5
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =5676533
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =4
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =31
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =5205
                    Top =420
                    Width =1545
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="MakeSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5205
                    LayoutCachedTop =420
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =660
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
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =8670
                    Top =420
                    Width =1650
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    Name ="TypeSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8670
                    LayoutCachedTop =420
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =660
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
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =6810
                    Top =420
                    Width =1800
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="ClassSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6810
                    LayoutCachedTop =420
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =660
                    ColumnStart =8
                    ColumnEnd =8
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
                    TextFontCharSet =238
                    Left =11880
                    Top =420
                    Width =1920
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="VendorSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11880
                    LayoutCachedTop =420
                    LayoutCachedWidth =13800
                    LayoutCachedHeight =660
                    ColumnStart =9
                    ColumnEnd =9
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
                    TextFontCharSet =238
                    Left =13860
                    Top =420
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="InvoiceNumSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =13860
                    LayoutCachedTop =420
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =660
                    ColumnStart =10
                    ColumnEnd =10
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
                    TextFontCharSet =238
                    Left =15360
                    Top =420
                    Width =1065
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="InvoiceDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =15360
                    LayoutCachedTop =420
                    LayoutCachedWidth =16425
                    LayoutCachedHeight =660
                    ColumnStart =11
                    ColumnEnd =11
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
                    TextFontCharSet =238
                    Left =16530
                    Top =420
                    Width =1065
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    Name ="VendorPayDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =16530
                    LayoutCachedTop =420
                    LayoutCachedWidth =17595
                    LayoutCachedHeight =660
                    ColumnStart =13
                    ColumnEnd =13
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
                    TextFontCharSet =238
                    Left =17655
                    Top =420
                    Width =990
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    Name ="CostSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =17655
                    LayoutCachedTop =420
                    LayoutCachedWidth =18645
                    LayoutCachedHeight =660
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
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =21990
                    Top =420
                    Width =660
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    Name ="DrawDownSubmit"
                    Caption ="▼"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =21990
                    LayoutCachedTop =420
                    LayoutCachedWidth =22650
                    LayoutCachedHeight =660
                    ColumnStart =12
                    ColumnEnd =12
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
                    TextFontCharSet =238
                    Left =18690
                    Top =420
                    Width =990
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    Name ="cIfAmountSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =18690
                    LayoutCachedTop =420
                    LayoutCachedWidth =19680
                    LayoutCachedHeight =660
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
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =22710
                    Top =420
                    Width =663
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =16
                    Name ="ExcludeCostDocsReportsSubmit"
                    Caption ="▼"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =22710
                    LayoutCachedTop =420
                    LayoutCachedWidth =23373
                    LayoutCachedHeight =660
                    ColumnStart =15
                    ColumnEnd =15
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
                    TextFontCharSet =238
                    Left =23433
                    Top =420
                    Width =648
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =17
                    Name ="EXCofaSubmit"
                    Caption ="▼"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =23433
                    LayoutCachedTop =420
                    LayoutCachedWidth =24081
                    LayoutCachedHeight =660
                    ColumnStart =15
                    ColumnEnd =15
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
                    ListRows =20
                    Left =5205
                    Top =750
                    Width =1545
                    ColumnOrder =7
                    FontSize =9
                    TabIndex =18
                    BoundColumn =1
                    Name ="MakeCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_ChildPartMake.ChildPartMakeID, LookUp_ChildPartMake.ChildPartMake "
                        "FROM LookUp_ChildPartMake ORDER BY LookUp_ChildPartMake.ChildPartMake;"
                    ColumnWidths ="0;1440"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =5205
                    LayoutCachedTop =750
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    Left =8670
                    Top =750
                    Width =1650
                    ColumnOrder =8
                    FontSize =9
                    TabIndex =20
                    Name ="TypeCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Lookup_PartTypes.cPartType, Lookup_PartTypes.[cPartClassID] FROM Lookup_P"
                        "artTypes WHERE (((Lookup_PartTypes.[cPartClassID])=50)) ORDER BY Lookup_PartType"
                        "s.cPartType;"
                    ColumnWidths ="1800;0"
                    OnGotFocus ="[Event Procedure]"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =8670
                    LayoutCachedTop =750
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =11880
                    Top =750
                    Width =1920
                    ColumnOrder =10
                    FontSize =9
                    TabIndex =21
                    Name ="ChildVendorCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName FROM Vendors WHERE (((Vendors.Vendor"
                        "Obsolete)=No)) ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;3600"
                    StatusBarText ="lookup from Vendor table"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Vendor is required"
                    Tag ="Medit"
                    RightPadding =38
                    OnDblClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenReport"
                            Argument ="VendorPmtsgrpChildInstRpt"
                            Argument ="2"
                            Argument =""
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"ChildVendorCbo\" xmlns=\"http://schemas.microsoft.com/office"
                                "/accessservices/2009/11/application\"><Statements><Action Name=\"OpenReport\"><A"
                                "rgument Name=\"ReportName\">Vendo"
                        End
                        Begin
                            Comment ="_AXL:rPmtsgrpChildInstRpt</Argument><Argument Name=\"View\">Print Preview</Argum"
                                "ent></Action></Statements></UserInterfaceMacro>"
                        End
                    End
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =11880
                    LayoutCachedTop =750
                    LayoutCachedWidth =13800
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =9
                    ColumnEnd =9
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =13860
                    Top =750
                    ColumnOrder =11
                    FontSize =9
                    TabIndex =22
                    Name ="InvoiceNum"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =13860
                    LayoutCachedTop =750
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =10
                    ColumnEnd =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15360
                    Top =750
                    Width =1065
                    ColumnOrder =12
                    FontSize =9
                    TabIndex =23
                    Name ="InvoiceDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =15360
                    LayoutCachedTop =750
                    LayoutCachedWidth =16425
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =16530
                    Top =750
                    Width =1065
                    ColumnOrder =14
                    FontSize =9
                    TabIndex =24
                    Name ="VendorPayDate"
                    Format ="Short Date"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =16530
                    LayoutCachedTop =750
                    LayoutCachedWidth =17595
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =13
                    ColumnEnd =13
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =17655
                    Top =750
                    Width =990
                    ColumnOrder =15
                    FontSize =9
                    TabIndex =25
                    BackColor =15858167
                    Name ="CostCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =17655
                    LayoutCachedTop =750
                    LayoutCachedWidth =18645
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =14
                    ColumnEnd =14
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =4320
                    Left =21990
                    Top =750
                    Width =660
                    ColumnOrder =13
                    FontSize =9
                    TabIndex =28
                    BackColor =15590879
                    Name ="DrawDownCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DrawDowns.DrawDownId, DrawDowns.DrawNumber AS Draw, DrawDowns.DrawName, B"
                        "anks.BankShortName AS Bank, Banks.BankID FROM DrawDowns INNER JOIN Banks ON Draw"
                        "Downs.DrawBankId = Banks.bankID WHERE (((DrawDowns.GroupID)=[Forms]![BulkChange_"
                        "Form]![GrpID]));"
                    ColumnWidths ="0;432;1296;864;0"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =21990
                    LayoutCachedTop =750
                    LayoutCachedWidth =22650
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
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
                    Left =22710
                    Top =750
                    Width =663
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =29
                    Name ="ExcludeCostDocsReportsCbo"
                    RowSourceType ="Value List"
                    RowSource ="-1;Hide;0;Show"
                    ColumnWidths ="0;634"
                    Tag ="Medit"
                    ControlTipText ="Show or Hide Part on Costing / Schedule Reports"
                    RightPadding =38

                    LayoutCachedLeft =22710
                    LayoutCachedTop =750
                    LayoutCachedWidth =23373
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
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
                    Left =23433
                    Top =750
                    Width =648
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =30
                    Name ="EXCofaCbox"
                    RowSourceType ="Value List"
                    RowSource ="-1;Hide;0;Show"
                    ColumnWidths ="0;634"
                    Tag ="Medit"
                    ControlTipText ="Show or Hide Part on COA Reports"
                    RightPadding =38

                    LayoutCachedLeft =23433
                    LayoutCachedTop =750
                    LayoutCachedWidth =24081
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =15
                    ColumnEnd =15
                    BackThemeColorIndex =9
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =4740
                    Top =420
                    Width =420
                    Height =225
                    FontSize =8
                    FontWeight =700
                    Name ="Label247"
                    Caption ="◄Sel"
                    LayoutCachedLeft =4740
                    LayoutCachedTop =420
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =645
                    ForeThemeColorIndex =2
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =18690
                    Top =1050
                    Width =990
                    Height =285
                    FontSize =9
                    ForeColor =16777215
                    Name ="Label522"
                    Caption ="IF Amount"
                    RightPadding =38
                    LayoutCachedLeft =18690
                    LayoutCachedTop =1050
                    LayoutCachedWidth =19680
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    BackThemeColorIndex =8
                    BackShade =75.0
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =238
                    Left =30
                    Top =1050
                    Width =705
                    Height =285
                    FontSize =9
                    Name ="Label180"
                    Caption ="AssetID"
                    RightPadding =38
                    LayoutCachedLeft =30
                    LayoutCachedTop =1050
                    LayoutCachedWidth =735
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =810
                    Top =1050
                    Width =45
                    Height =285
                    FontSize =9
                    Name ="Label182"
                    Caption ="Child ID"
                    RightPadding =38
                    LayoutCachedLeft =810
                    LayoutCachedTop =1050
                    LayoutCachedWidth =855
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =3
                    Left =3855
                    Top =1050
                    Width =1245
                    Height =270
                    FontSize =9
                    Name ="Label181"
                    Caption ="    :             Ser."
                    RightPadding =38
                    LayoutCachedLeft =3855
                    LayoutCachedTop =1050
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =1320
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =5205
                    Top =1050
                    Width =1545
                    Height =285
                    FontSize =9
                    Name ="Label229"
                    Caption ="Make"
                    RightPadding =38
                    LayoutCachedLeft =5205
                    LayoutCachedTop =1050
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =6
                    ColumnEnd =6
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =6810
                    Top =1050
                    Width =1800
                    Height =285
                    FontSize =9
                    Name ="Label230"
                    Caption ="Class"
                    RightPadding =38
                    LayoutCachedLeft =6810
                    LayoutCachedTop =1050
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =7
                    ColumnEnd =7
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =8670
                    Top =1050
                    Width =1650
                    Height =285
                    FontSize =9
                    Name ="Label231"
                    Caption ="Type"
                    RightPadding =38
                    LayoutCachedLeft =8670
                    LayoutCachedTop =1050
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =8
                    ColumnEnd =8
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =11880
                    Top =1050
                    Width =1920
                    Height =285
                    FontSize =9
                    Name ="Label14"
                    Caption ="Child Vendor:"
                    RightPadding =38
                    LayoutCachedLeft =11880
                    LayoutCachedTop =1050
                    LayoutCachedWidth =13800
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =9
                    ColumnEnd =9
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =13860
                    Top =1050
                    Width =1440
                    Height =285
                    FontSize =9
                    Name ="Label236"
                    Caption ="Invoice Num"
                    RightPadding =38
                    LayoutCachedLeft =13860
                    LayoutCachedTop =1050
                    LayoutCachedWidth =15300
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =10
                    ColumnEnd =10
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =15360
                    Top =1050
                    Width =1065
                    Height =285
                    FontSize =9
                    Name ="Label239"
                    Caption ="Invoice Date"
                    RightPadding =38
                    LayoutCachedLeft =15360
                    LayoutCachedTop =1050
                    LayoutCachedWidth =16425
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =11
                    ColumnEnd =11
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =16530
                    Top =1050
                    Width =1065
                    Height =285
                    FontSize =9
                    Name ="Label126"
                    Caption ="Paymt Date:"
                    RightPadding =38
                    LayoutCachedLeft =16530
                    LayoutCachedTop =1050
                    LayoutCachedWidth =17595
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =13
                    ColumnEnd =13
                End
                Begin Label
                    OverlapFlags =93
                    Left =17655
                    Top =1050
                    Width =990
                    Height =285
                    FontSize =9
                    Name ="Label335"
                    Caption ="Part Cost"
                    RightPadding =38
                    LayoutCachedLeft =17655
                    LayoutCachedTop =1050
                    LayoutCachedWidth =18645
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =14
                    ColumnEnd =14
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    Left =21990
                    Top =1050
                    Width =660
                    Height =285
                    FontSize =9
                    ForeColor =16777215
                    Name ="Label203"
                    Caption ="DRAW"
                    RightPadding =38
                    LayoutCachedLeft =21990
                    LayoutCachedTop =1050
                    LayoutCachedWidth =22650
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    BackThemeColorIndex =8
                    BackShade =75.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =22710
                    Top =1050
                    Width =450
                    Height =270
                    FontSize =9
                    Name ="Label416"
                    Caption ="RPTS"
                    RightPadding =38
                    LayoutCachedLeft =22710
                    LayoutCachedTop =1050
                    LayoutCachedWidth =23160
                    LayoutCachedHeight =1320
                    RowStart =3
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =23433
                    Top =1050
                    Width =648
                    Height =210
                    FontSize =9
                    FontWeight =700
                    Name ="Label507"
                    Caption ="COA"
                    RightPadding =38
                    LayoutCachedLeft =23433
                    LayoutCachedTop =1050
                    LayoutCachedWidth =24081
                    LayoutCachedHeight =1260
                    RowStart =3
                    RowEnd =3
                    ColumnStart =15
                    ColumnEnd =15
                    BackThemeColorIndex =9
                    BackShade =50.0
                    ForeThemeColorIndex =9
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =20865
                    Top =420
                    Width =1080
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =14
                    Name ="cIfBankSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =20865
                    LayoutCachedTop =420
                    LayoutCachedWidth =21945
                    LayoutCachedHeight =660
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
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =20865
                    Top =750
                    Width =1080
                    FontSize =9
                    TabIndex =27
                    BackColor =15590879
                    Name ="cIFBankCBO"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankID, Banks.BankShortName, Banks.BankName FROM Banks ORDER BY Ban"
                        "ks.BankShortName;"
                    ColumnWidths ="0;864;2736"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =20865
                    LayoutCachedTop =750
                    LayoutCachedWidth =21945
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =12
                    ColumnEnd =12
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =20865
                    Top =1050
                    Width =1080
                    Height =285
                    FontSize =9
                    Name ="Label529"
                    Caption ="IF Bank"
                    RightPadding =38
                    LayoutCachedLeft =20865
                    LayoutCachedTop =1050
                    LayoutCachedWidth =21945
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    BackThemeColorIndex =8
                    BackTint =60.0
                    ForeThemeColorIndex =0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =19752
                    Top =756
                    Width =1068
                    FontSize =9
                    TabIndex =26
                    BackColor =15788753
                    Name ="cIFDateCBO"
                    Format ="Short Date"
                    GroupTable =27

                    LayoutCachedLeft =19752
                    LayoutCachedTop =756
                    LayoutCachedWidth =20820
                    LayoutCachedHeight =996
                    LayoutGroup =3
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =27
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =19755
                    Top =1050
                    Width =1065
                    Height =270
                    FontSize =9
                    BackColor =15788753
                    Name ="Label590"
                    Caption ="Fund Date"
                    RightPadding =38
                    LayoutCachedLeft =19755
                    LayoutCachedTop =1050
                    LayoutCachedWidth =20820
                    LayoutCachedHeight =1320
                    RowStart =3
                    RowEnd =3
                    ColumnStart =12
                    ColumnEnd =12
                    ForeThemeColorIndex =0
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Left =60
                    Top =1260
                    Width =24048
                    Height =195
                    Name ="Image454"
                    Picture ="LtGrey2_Shadow"
                    GroupTable =6
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedLeft =60
                    LayoutCachedTop =1260
                    LayoutCachedWidth =24108
                    LayoutCachedHeight =1455
                    TabIndex =33
                    LayoutGroup =1
                    GroupTable =6
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =19755
                    Top =420
                    Width =1080
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    Name ="cIFDateSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =19755
                    LayoutCachedTop =420
                    LayoutCachedWidth =20835
                    LayoutCachedHeight =660
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
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =18690
                    Top =750
                    Width =990
                    Height =210
                    FontSize =9
                    Name ="Label306"
                    Caption ="Interim All"
                    RightPadding =38
                    LayoutCachedLeft =18690
                    LayoutCachedTop =750
                    LayoutCachedWidth =19680
                    LayoutCachedHeight =960
                    RowStart =3
                    RowEnd =3
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =2
                    BackTint =60.0
                    ForeThemeColorIndex =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =4320
                    Left =6810
                    Top =750
                    Width =1800
                    ColumnOrder =9
                    FontSize =9
                    TabIndex =19
                    Name ="ClassCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT     Lookup_PartClasses.cPartClass,     Lookup_PartClasses.cPartClassID FR"
                        "OM     Lookup_PartClasses ORDER BY     Lookup_PartClasses.cPartClass;"
                    ColumnWidths ="4320;0"
                    StatusBarText ="UNIT TYPE"
                    OnClick ="[Event Procedure]"
                    Tag ="Medit"
                    RightPadding =38
                    AllowValueListEdits =0

                    LayoutCachedLeft =6810
                    LayoutCachedTop =750
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =10380
                    Top =420
                    Width =1470
                    Height =240
                    FontSize =8
                    FontWeight =700
                    TabIndex =31
                    Name ="cDescSubmit"
                    Caption ="▼Submit"
                    OnClick ="[Event Procedure]"
                    RightPadding =38
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10380
                    LayoutCachedTop =420
                    LayoutCachedWidth =11850
                    LayoutCachedHeight =660
                    ColumnStart =11
                    ColumnEnd =11
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
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10380
                    Top =750
                    Width =1470
                    FontSize =9
                    TabIndex =32
                    Name ="cDescCbo"
                    Format ="Short Date"
                    Tag ="Medit"
                    RightPadding =38

                    LayoutCachedLeft =10380
                    LayoutCachedTop =750
                    LayoutCachedWidth =11850
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =11
                    ColumnEnd =11
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
        Begin Section
            Height =417
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =23700
                    Top =36
                    Width =660
                    Height =312
                    FontSize =9
                    TabIndex =18
                    Name ="ChildEXCofa"
                    ControlSource ="ChildEXCofa"
                    RowSourceType ="Value List"
                    RowSource ="-1;Hide;0;Show"
                    ColumnWidths ="0;634"
                    ControlTipText ="Show or Hide Part on COA Reports"
                    Format ="True/False"
                    GroupTable =7

                    LayoutCachedLeft =23700
                    LayoutCachedTop =36
                    LayoutCachedWidth =24360
                    LayoutCachedHeight =348
                    ColumnStart =18
                    ColumnEnd =18
                    LayoutGroup =2
                    BackThemeColorIndex =9
                    BackTint =20.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =22968
                    Top =36
                    Width =660
                    Height =312
                    FontSize =9
                    TabIndex =17
                    Name ="ExcludeCostDocsReports"
                    ControlSource ="ExcludeCostDocsReports"
                    RowSourceType ="Value List"
                    RowSource ="-1;Hide;0;Show"
                    ColumnWidths ="0;634"
                    ControlTipText ="Show or Hide Part on Costing / Schedule Reports"
                    Format ="True/False"
                    GroupTable =7

                    LayoutCachedLeft =22968
                    LayoutCachedTop =36
                    LayoutCachedWidth =23628
                    LayoutCachedHeight =348
                    ColumnStart =17
                    ColumnEnd =17
                    LayoutGroup =2
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =36
                    Top =36
                    Width =708
                    Height =312
                    FontSize =9
                    Name ="UnitID"
                    ControlSource ="UnitRef"
                    GroupTable =7

                    LayoutCachedLeft =36
                    LayoutCachedTop =36
                    LayoutCachedWidth =744
                    LayoutCachedHeight =348
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3828
                    Top =36
                    Width =1356
                    Height =312
                    FontSize =9
                    TabIndex =3
                    Name ="ChildVIN"
                    ControlSource ="ChildVIN"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =7

                    LayoutCachedLeft =3828
                    LayoutCachedTop =36
                    LayoutCachedWidth =5184
                    LayoutCachedHeight =348
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =810
                    Top =45
                    Width =45
                    Height =315
                    FontSize =9
                    TabIndex =19
                    Name ="ChildPartID"
                    ControlSource ="ChildPartID"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =810
                    LayoutCachedTop =45
                    LayoutCachedWidth =855
                    LayoutCachedHeight =360
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =12024
                    Top =36
                    Width =1920
                    Height =312
                    FontSize =9
                    TabIndex =8
                    Name ="ChildVendorREF"
                    ControlSource ="ChildVendorREF"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorName FROM Vendors WHERE (((Vendors.Vendor"
                        "Obsolete)=No)) ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;3600"
                    StatusBarText ="lookup from Vendor table"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Vendor is required"
                    GroupTable =7
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =12024
                    LayoutCachedTop =36
                    LayoutCachedWidth =13944
                    LayoutCachedHeight =348
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5256
                    Top =36
                    Width =1548
                    Height =312
                    FontSize =9
                    TabIndex =4
                    Name ="ChildMake"
                    ControlSource ="ChildMake"
                    StatusBarText ="IRP Acct Number"
                    GroupTable =7

                    LayoutCachedLeft =5256
                    LayoutCachedTop =36
                    LayoutCachedWidth =6804
                    LayoutCachedHeight =348
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =16668
                    Top =36
                    Width =1068
                    Height =312
                    FontSize =9
                    TabIndex =11
                    Name ="ChildVendorPytDate"
                    ControlSource ="ChildVendorPytDate"
                    Format ="Short Date"
                    GroupTable =7

                    LayoutCachedLeft =16668
                    LayoutCachedTop =36
                    LayoutCachedWidth =17736
                    LayoutCachedHeight =348
                    ColumnStart =11
                    ColumnEnd =11
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =14016
                    Top =36
                    Height =312
                    FontSize =9
                    TabIndex =9
                    Name ="ChildInvoiceNum"
                    ControlSource ="ChildInvoiceNum"
                    GroupTable =7

                    LayoutCachedLeft =14016
                    LayoutCachedTop =36
                    LayoutCachedWidth =15456
                    LayoutCachedHeight =348
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =15528
                    Top =36
                    Width =1068
                    Height =312
                    FontSize =9
                    TabIndex =10
                    Name ="ChildInvdt"
                    ControlSource ="ChildInvdt"
                    Format ="Short Date"
                    GroupTable =7

                    LayoutCachedLeft =15528
                    LayoutCachedTop =36
                    LayoutCachedWidth =16596
                    LayoutCachedHeight =348
                    ColumnStart =10
                    ColumnEnd =10
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =17808
                    Top =36
                    Width =996
                    Height =312
                    FontSize =9
                    TabIndex =12
                    Name ="childFinalCost"
                    ControlSource ="childFinalCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =7

                    LayoutCachedLeft =17808
                    LayoutCachedTop =36
                    LayoutCachedWidth =18804
                    LayoutCachedHeight =348
                    ColumnStart =12
                    ColumnEnd =12
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1896
                    Top =36
                    Width =1860
                    Height =312
                    FontSize =9
                    TabIndex =2
                    Name ="unitvin"
                    ControlSource ="unitvin"
                    GroupTable =7

                    LayoutCachedLeft =1896
                    LayoutCachedTop =36
                    LayoutCachedWidth =3756
                    LayoutCachedHeight =348
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =4320
                    Left =22236
                    Top =36
                    Width =660
                    Height =312
                    FontSize =9
                    TabIndex =16
                    BackColor =15590879
                    Name ="DrawDownId"
                    ControlSource ="DrawDownId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DrawDowns.DrawDownId, DrawDowns.DrawNumber AS Draw, DrawDowns.DrawName, B"
                        "anks.BankShortName AS Bank, Banks.BankID FROM DrawDowns INNER JOIN Banks ON Draw"
                        "Downs.DrawBankId = Banks.bankID WHERE (((DrawDowns.GroupID)=[Forms]![BulkChange_"
                        "Form]![GrpID]));"
                    ColumnWidths ="0;432;1296;864;0"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    AfterUpdate ="[Event Procedure]"
                    Format ="Standard"
                    GroupTable =7

                    LayoutCachedLeft =22236
                    LayoutCachedTop =36
                    LayoutCachedWidth =22896
                    LayoutCachedHeight =348
                    ColumnStart =16
                    ColumnEnd =16
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Left =816
                    Top =36
                    Width =1008
                    Height =312
                    FontSize =9
                    TabIndex =1
                    Name ="unitunitnum"
                    ControlSource ="unitunitnum"
                    GroupTable =7

                    LayoutCachedLeft =816
                    LayoutCachedTop =36
                    LayoutCachedWidth =1824
                    LayoutCachedHeight =348
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =18876
                    Top =36
                    Width =996
                    Height =312
                    FontSize =9
                    TabIndex =13
                    BackColor =15788753
                    Name ="cIFAmount"
                    ControlSource ="cIFAmount"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =7

                    LayoutCachedLeft =18876
                    LayoutCachedTop =36
                    LayoutCachedWidth =19872
                    LayoutCachedHeight =348
                    ColumnStart =13
                    ColumnEnd =13
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =60
                    Top =345
                    Width =24048
                    Height =72
                    TabIndex =20
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

                    LayoutCachedLeft =60
                    LayoutCachedTop =345
                    LayoutCachedWidth =24108
                    LayoutCachedHeight =417
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff001f497d00150000005b00 ,
                        0x55006e0069007400490044005d003d005b007400780074004300750072007200 ,
                        0x65006e0074005d00000000000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =119
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =21084
                    Top =36
                    Width =1080
                    Height =312
                    FontSize =9
                    TabIndex =15
                    BackColor =16381933
                    Name ="cIFBankId"
                    ControlSource ="cIFBankId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Banks.BankID, Banks.BankShortName, Banks.BankName FROM Banks ORDER BY Ban"
                        "ks.BankShortName;"
                    ColumnWidths ="0;864;2736"
                    StatusBarText ="DAILY PER DIEM RENT AMOUNT"
                    AfterUpdate ="[Event Procedure]"
                    Format ="Standard"
                    GroupTable =7

                    LayoutCachedLeft =21084
                    LayoutCachedTop =36
                    LayoutCachedWidth =22164
                    LayoutCachedHeight =348
                    ColumnStart =15
                    ColumnEnd =15
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    OverlapFlags =119
                    TextAlign =1
                    IMESentenceMode =3
                    Left =19944
                    Top =36
                    Width =1068
                    Height =312
                    FontSize =9
                    TabIndex =14
                    Name ="cIFDate"
                    ControlSource ="cIFDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    GroupTable =7

                    LayoutCachedLeft =19944
                    LayoutCachedTop =36
                    LayoutCachedWidth =21012
                    LayoutCachedHeight =348
                    ColumnStart =14
                    ColumnEnd =14
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    OverlapFlags =119
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10476
                    Top =36
                    Width =1476
                    Height =312
                    FontSize =9
                    TabIndex =7
                    Name ="cDesc"
                    ControlSource ="cDesc"
                    StatusBarText ="VIN number must be unique in table"
                    GroupTable =7

                    LayoutCachedLeft =10476
                    LayoutCachedTop =36
                    LayoutCachedWidth =11952
                    LayoutCachedHeight =348
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =119
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =4320
                    Left =6876
                    Top =36
                    Width =1800
                    Height =312
                    FontSize =9
                    TabIndex =5
                    Name ="ChildClass"
                    ControlSource ="ChildClass"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT     Lookup_PartClasses.cPartClass,     Lookup_PartClasses.cPartClassID FR"
                        "OM     Lookup_PartClasses ORDER BY     Lookup_PartClasses.cPartClass;"
                    ColumnWidths ="4320;0"
                    StatusBarText ="UNIT TYPE"
                    OnClick ="[Event Procedure]"
                    GroupTable =7
                    AllowValueListEdits =0

                    LayoutCachedLeft =6876
                    LayoutCachedTop =36
                    LayoutCachedWidth =8676
                    LayoutCachedHeight =348
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =119
                    TextAlign =1
                    IMESentenceMode =3
                    ListRows =20
                    Left =8748
                    Top =36
                    Width =1656
                    Height =312
                    FontSize =9
                    TabIndex =6
                    Name ="ChildType"
                    ControlSource ="ChildType"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="1800"
                    OnGotFocus ="[Event Procedure]"
                    GroupTable =7

                    LayoutCachedLeft =8748
                    LayoutCachedTop =36
                    LayoutCachedWidth =10404
                    LayoutCachedHeight =348
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
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
' See "BulkChange_ChildParts.cls"
