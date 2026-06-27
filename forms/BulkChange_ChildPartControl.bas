Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =19635
    DatasheetFontHeight =11
    ItemSuffix =316
    Left =16140
    Top =6660
    Right =-26791
    Bottom =19365
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x7b0c69c77188e640
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Bulk_Update_SelectedUnits"
    DatasheetFontName ="Calibri"
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
            Height =998
            BackColor =14211288
            Name ="FormHeader"
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Top =675
                    Width =19635
                    Height =274
                    BackColor =1643706
                    BorderColor =10855845
                    Name ="Box22"
                    HorizontalAnchor =2
                    LayoutCachedTop =675
                    LayoutCachedWidth =19635
                    LayoutCachedHeight =949
                End
                Begin Label
                    OverlapFlags =223
                    Left =115
                    Top =230
                    Width =7830
                    Height =660
                    FontSize =22
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label115"
                    Caption ="Insert / Delete Child Parts"
                    LayoutCachedLeft =115
                    LayoutCachedTop =230
                    LayoutCachedWidth =7945
                    LayoutCachedHeight =890
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =16530
                    Top =105
                    Width =765
                    Height =210
                    ColumnOrder =0
                    FontSize =8
                    Name ="txtCurrent"

                    LayoutCachedLeft =16530
                    LayoutCachedTop =105
                    LayoutCachedWidth =17295
                    LayoutCachedHeight =315
                End
                Begin Image
                    PictureTiling = NotDefault
                    PictureAlignment =0
                    PictureType =2
                    Top =840
                    Width =19635
                    Height =158
                    Name ="Image454"
                    Picture ="Red_Dark_Shadow"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    HorizontalAnchor =2

                    LayoutCachedTop =840
                    LayoutCachedWidth =19635
                    LayoutCachedHeight =998
                    TabIndex =1
                    LayoutGroup =2
                    GroupTable =7
                End
            End
        End
        Begin Section
            Height =11220
            Name ="Detail"
            AlternateBackColor =-2147483606
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =1320
                    Top =1320
                    Width =840
                    Height =795
                    BackColor =6697472
                    ForeColor =16777215
                    Name ="UnitCountLabel"
                    Caption ="Units"
                    LayoutCachedLeft =1320
                    LayoutCachedTop =1320
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =2115
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =6
                    ListWidth =9216
                    Left =4995
                    Top =360
                    Width =3195
                    Height =300
                    FontSize =9
                    Name ="ChildVendorCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Vendors.VendorID, Vendors.VendorSNm, Vendors.VendorName, Vendors.VendorAd"
                        "dress, Vendors.VendorCity, Vendors.VendorState FROM Vendors WHERE (((Vendors.Ven"
                        "dorObsolete)=No)) ORDER BY Vendors.VendorName;"
                    ColumnWidths ="0;1728;2880;2160;1440;432"
                    StatusBarText ="lookup from Vendor table"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Vendor is required"
                    Tag ="Medit"
                    GroupTable =6
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =4995
                    LayoutCachedTop =360
                    LayoutCachedWidth =8190
                    LayoutCachedHeight =660
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =6
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =1
                            Left =3345
                            Top =360
                            Width =1590
                            Height =300
                            ForeColor =16777215
                            Name ="Label221"
                            Caption ="Vendor:"
                            GroupTable =6
                            LayoutCachedLeft =3345
                            LayoutCachedTop =360
                            LayoutCachedWidth =4935
                            LayoutCachedHeight =660
                            LayoutGroup =1
                            BackThemeColorIndex =2
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4995
                    Top =1800
                    Width =3195
                    Height =300
                    FontSize =9
                    TabIndex =4
                    BackColor =15858167
                    Name ="CostCurr"
                    Format ="$#,##0.00;($#,##0.00)"
                    Tag ="Medit"
                    GroupTable =6

                    LayoutCachedLeft =4995
                    LayoutCachedTop =1800
                    LayoutCachedWidth =8190
                    LayoutCachedHeight =2100
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =6
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =1
                            Left =3345
                            Top =1800
                            Width =1590
                            Height =300
                            ForeColor =16777215
                            Name ="Label166"
                            Caption ="Cost:"
                            GroupTable =6
                            LayoutCachedLeft =3345
                            LayoutCachedTop =1800
                            LayoutCachedWidth =4935
                            LayoutCachedHeight =2100
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            BackThemeColorIndex =2
                            GroupTable =6
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =8700
                    Top =960
                    Width =2766
                    Height =321
                    TabIndex =6
                    Name ="InsertChildParts"
                    Caption ="◄ Add Child Part to Units"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add Record"
                    LeftPadding =90
                    RightPadding =105
                    BottomPadding =165

                    LayoutCachedLeft =8700
                    LayoutCachedTop =960
                    LayoutCachedWidth =11466
                    LayoutCachedHeight =1281
                    PictureCaptionArrangement =2
                    Alignment =1
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackThemeColorIndex =2
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =2
                    PressedThemeColorIndex =2
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =2
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =161
                    Left =3135
                    Top =1080
                    Width =240
                    Height =525
                    FontSize =12
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label247"
                    Caption ="┌\015\012└"
                    FontName ="Arial"
                    LayoutCachedLeft =3135
                    LayoutCachedTop =1080
                    LayoutCachedWidth =3375
                    LayoutCachedHeight =1605
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =238
                    Left =705
                    Top =360
                    Width =2550
                    Height =585
                    FontWeight =700
                    Name ="Label250"
                    Caption ="Select the items that 'Describe' the Child Part►"
                    LayoutCachedLeft =705
                    LayoutCachedTop =360
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =945
                    ForeThemeColorIndex =2
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1350
                    Top =1590
                    Width =780
                    Height =456
                    FontSize =20
                    FontWeight =700
                    TabIndex =7
                    Name ="UnitCount"
                    ControlSource ="=Count([UnitID])"

                    LayoutCachedLeft =1350
                    LayoutCachedTop =1590
                    LayoutCachedWidth =2130
                    LayoutCachedHeight =2046
                End
                Begin Label
                    OverlapFlags =87
                    TextFontCharSet =238
                    Left =240
                    Top =360
                    Width =465
                    Height =585
                    FontSize =22
                    FontWeight =700
                    Name ="Label252"
                    Caption ="1:"
                    LayoutCachedLeft =240
                    LayoutCachedTop =360
                    LayoutCachedWidth =705
                    LayoutCachedHeight =945
                End
                Begin Label
                    OverlapFlags =223
                    TextFontCharSet =238
                    Left =8700
                    Top =360
                    Width =2805
                    Height =585
                    FontWeight =700
                    Name ="Label253"
                    Caption ="Press the \015\012▼ 'Insert Child Parts' button"
                    LayoutCachedLeft =8700
                    LayoutCachedTop =360
                    LayoutCachedWidth =11505
                    LayoutCachedHeight =945
                    ForeThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    Left =8235
                    Top =360
                    Width =465
                    Height =585
                    FontSize =22
                    FontWeight =700
                    Name ="Label254"
                    Caption ="2:"
                    LayoutCachedLeft =8235
                    LayoutCachedTop =360
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =945
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =238
                    TextAlign =2
                    Left =780
                    Top =900
                    Width =2220
                    Height =285
                    FontSize =8
                    FontWeight =700
                    ForeColor =-2147483617
                    Name ="Label256"
                    Caption ="(all fields are required)"
                    LayoutCachedLeft =780
                    LayoutCachedTop =900
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =1185
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8700
                    Top =4200
                    Width =2766
                    Height =321
                    TabIndex =9
                    Name ="DeleteChildPart"
                    Caption ="◄ Delete Selected Part"
                    OnClick ="[Event Procedure]"
                    LeftPadding =90
                    RightPadding =105
                    BottomPadding =165
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8700
                    LayoutCachedTop =4200
                    LayoutCachedWidth =11466
                    LayoutCachedHeight =4521
                    Alignment =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =25
                    BackThemeColorIndex =5
                    BorderThemeColorIndex =1
                    HoverThemeColorIndex =5
                    HoverTint =80.0
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =2
                    QuickStyle =31
                    QuickStyleMask =-69
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =3345
                    Top =3285
                    Width =4815
                    Height =7815
                    FontSize =8
                    TabIndex =8
                    BoundColumn =3
                    ForeColor =-2147483617
                    Name ="ChildList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Bulk_Update_SelectedUnits_Child.ChildMake AS Make, Bulk_Update_SelectedUn"
                        "its_Child.ChildClass AS Class, Bulk_Update_SelectedUnits_Child.ChildType AS Type"
                        ", [ChildType] & [ChildVendorREF] & [ChildMake] & [ChildClass] AS UID FROM Bulk_U"
                        "pdate_SelectedUnits_Child GROUP BY Bulk_Update_SelectedUnits_Child.ChildMake, Bu"
                        "lk_Update_SelectedUnits_Child.ChildClass, Bulk_Update_SelectedUnits_Child.ChildT"
                        "ype, [ChildType] & [ChildVendorREF] & [ChildMake] & [ChildClass] ORDER BY Bulk_U"
                        "pdate_SelectedUnits_Child.ChildMake, Bulk_Update_SelectedUnits_Child.ChildType;"
                    ColumnWidths ="1584;1584;1584;0"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3345
                    LayoutCachedTop =3285
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =11100
                    BackThemeColorIndex =1
                End
                Begin Label
                    OverlapFlags =93
                    Left =705
                    Top =3315
                    Width =2550
                    Height =585
                    FontWeight =700
                    ForeColor =2366701
                    Name ="Label257"
                    Caption ="Select A Child Part to 'Delete' ►"
                    LayoutCachedLeft =705
                    LayoutCachedTop =3315
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =3900
                End
                Begin Label
                    OverlapFlags =95
                    Left =240
                    Top =3315
                    Width =465
                    Height =585
                    FontSize =22
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label258"
                    Caption ="X:"
                    LayoutCachedLeft =240
                    LayoutCachedTop =3315
                    LayoutCachedWidth =705
                    LayoutCachedHeight =3900
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =705
                    Top =3855
                    Width =2550
                    Height =420
                    FontSize =8
                    FontWeight =700
                    ForeColor =-2147483617
                    Name ="Label259"
                    Caption ="(this will delete that part from all selected units on this form)"
                    LayoutCachedLeft =705
                    LayoutCachedTop =3855
                    LayoutCachedWidth =3255
                    LayoutCachedHeight =4275
                End
                Begin Label
                    OverlapFlags =93
                    Left =8685
                    Top =3480
                    Width =2805
                    Height =585
                    FontWeight =700
                    ForeColor =2366701
                    Name ="Label260"
                    Caption ="Press the \015\012▼ 'Delete Selected' button"
                    LayoutCachedLeft =8685
                    LayoutCachedTop =3480
                    LayoutCachedWidth =11490
                    LayoutCachedHeight =4065
                End
                Begin Label
                    OverlapFlags =87
                    Left =8220
                    Top =3480
                    Width =465
                    Height =585
                    FontSize =22
                    FontWeight =700
                    ForeColor =2366701
                    Name ="Label261"
                    Caption ="Z:"
                    LayoutCachedLeft =8220
                    LayoutCachedTop =3480
                    LayoutCachedWidth =8685
                    LayoutCachedHeight =4065
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    Left =4995
                    Top =720
                    Width =3195
                    Height =300
                    FontSize =9
                    TabIndex =1
                    BoundColumn =1
                    Name ="MakeCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LookUp_ChildPartMake.ChildPartMakeID, LookUp_ChildPartMake.ChildPartMake "
                        "FROM LookUp_ChildPartMake ORDER BY LookUp_ChildPartMake.ChildPartMake;"
                    ColumnWidths ="0;1440"
                    Tag ="Medit"
                    GroupTable =6

                    LayoutCachedLeft =4995
                    LayoutCachedTop =720
                    LayoutCachedWidth =8190
                    LayoutCachedHeight =1020
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =6
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =3345
                            Top =720
                            Width =1590
                            Height =300
                            ForeColor =16777215
                            Name ="Label272"
                            Caption ="Make:"
                            GroupTable =6
                            LayoutCachedLeft =3345
                            LayoutCachedTop =720
                            LayoutCachedWidth =4935
                            LayoutCachedHeight =1020
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            BackThemeColorIndex =2
                            GroupTable =6
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    ListWidth =4320
                    Left =4995
                    Top =1080
                    Width =3195
                    Height =300
                    FontSize =9
                    TabIndex =2
                    Name ="ClassCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT     Lookup_PartClasses.cPartClass,     Lookup_PartClasses.cPartClassID FR"
                        "OM     Lookup_PartClasses ORDER BY     Lookup_PartClasses.cPartClass;"
                    ColumnWidths ="4320;0"
                    StatusBarText ="UNIT TYPE"
                    OnClick ="[Event Procedure]"
                    Tag ="Medit"
                    GroupTable =6
                    AllowValueListEdits =0

                    LayoutCachedLeft =4995
                    LayoutCachedTop =1080
                    LayoutCachedWidth =8190
                    LayoutCachedHeight =1380
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =6
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =1
                            Left =3345
                            Top =1080
                            Width =1590
                            Height =300
                            ForeColor =16777215
                            Name ="Label281"
                            Caption ="Class:"
                            GroupTable =6
                            LayoutCachedLeft =3345
                            LayoutCachedTop =1080
                            LayoutCachedWidth =4935
                            LayoutCachedHeight =1380
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            BackThemeColorIndex =2
                            GroupTable =6
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =20
                    Left =4995
                    Top =1440
                    Width =3195
                    Height =300
                    FontSize =9
                    TabIndex =3
                    Name ="TypeCbo"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT cPartType, cPartClassID FROM Lookup_PartTypes WHERE cPartClassID=100 Orde"
                        "r By cPartType;"
                    ColumnWidths ="1800;0"
                    OnGotFocus ="[Event Procedure]"
                    Tag ="Medit"
                    GroupTable =6

                    LayoutCachedLeft =4995
                    LayoutCachedTop =1440
                    LayoutCachedWidth =8190
                    LayoutCachedHeight =1740
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =6
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =1
                            Left =3345
                            Top =1440
                            Width =1590
                            Height =300
                            ForeColor =16777215
                            Name ="Label290"
                            Caption ="Type:"
                            GroupTable =6
                            LayoutCachedLeft =3345
                            LayoutCachedTop =1440
                            LayoutCachedWidth =4935
                            LayoutCachedHeight =1740
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            BackThemeColorIndex =2
                            GroupTable =6
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4995
                    Top =2160
                    Width =3195
                    Height =300
                    FontSize =9
                    TabIndex =5
                    Name ="cDescCbo"
                    Tag ="Medit"
                    GroupTable =6

                    LayoutCachedLeft =4995
                    LayoutCachedTop =2160
                    LayoutCachedWidth =8190
                    LayoutCachedHeight =2460
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =6
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =1
                            Left =3345
                            Top =2160
                            Width =1590
                            Height =300
                            ForeColor =16777215
                            Name ="Label302"
                            Caption ="Desc: (optional)"
                            GroupTable =6
                            LayoutCachedLeft =3345
                            LayoutCachedTop =2160
                            LayoutCachedWidth =4935
                            LayoutCachedHeight =2460
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            BackThemeColorIndex =8
                            BackShade =75.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =6
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =11640
                    Top =780
                    Width =7860
                    Height =10260
                    FontSize =8
                    TabIndex =10
                    BoundColumn =3
                    ForeColor =-2147483617
                    Name ="SisterParts"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vw_Sch_Grp_Units_ChildParts.UnitGroup AS [Group], vw_Sch_Grp_Units_ChildP"
                        "arts.ChildMake AS Make, vw_Sch_Grp_Units_ChildParts.ChildClass AS Class, vw_Sch_"
                        "Grp_Units_ChildParts.ChildType AS Type, vw_Sch_Grp_Units_ChildParts.ChildCost AS"
                        " Cost FROM vw_Sch_Grp_Units_ChildParts WHERE (((vw_Sch_Grp_Units_ChildParts.SchI"
                        "d)=[Forms]![BulkChange_Form]![ScheduleList])) GROUP BY vw_Sch_Grp_Units_ChildPar"
                        "ts.UnitGroup, vw_Sch_Grp_Units_ChildParts.ChildMake, vw_Sch_Grp_Units_ChildParts"
                        ".ChildClass, vw_Sch_Grp_Units_ChildParts.ChildType, vw_Sch_Grp_Units_ChildParts."
                        "ChildCost ORDER BY vw_Sch_Grp_Units_ChildParts.UnitGroup, vw_Sch_Grp_Units_Child"
                        "Parts.ChildClass;"
                    ColumnWidths ="720;1872;1987;1872;1008"

                    LayoutCachedLeft =11640
                    LayoutCachedTop =780
                    LayoutCachedWidth =19500
                    LayoutCachedHeight =11040
                    BackThemeColorIndex =1
                End
                Begin Label
                    OverlapFlags =93
                    Left =11640
                    Top =420
                    Width =5100
                    Height =480
                    FontSize =12
                    FontWeight =700
                    Name ="Label312"
                    Caption ="▼  --  Child Parts That are in this Schedule -- ▼ "
                    LayoutCachedLeft =11640
                    LayoutCachedTop =420
                    LayoutCachedWidth =16740
                    LayoutCachedHeight =900
                    ForeThemeColorIndex =2
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =8700
                    Top =2100
                    Width =2766
                    Height =606
                    TabIndex =11
                    Name ="InsertSchChildParts"
                    Caption ="◄  Add Child Part to ALL  Units in SCHEDULE"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add Record"
                    LeftPadding =90
                    RightPadding =105
                    BottomPadding =165

                    LayoutCachedLeft =8700
                    LayoutCachedTop =2100
                    LayoutCachedWidth =11466
                    LayoutCachedHeight =2706
                    PictureCaptionArrangement =2
                    Alignment =1
                    ForeThemeColorIndex =1
                    UseTheme =1
                    BackColor =39423
                    BorderThemeColorIndex =1
                    HoverColor =39423
                    PressedColor =39423
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =2
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    Left =8700
                    Top =1500
                    Width =2805
                    Height =585
                    FontWeight =700
                    Name ="Label315"
                    Caption ="OR - Press the \015\012▼ 'Insert  Parts' to Schedule"
                    LayoutCachedLeft =8700
                    LayoutCachedTop =1500
                    LayoutCachedWidth =11505
                    LayoutCachedHeight =2085
                    ForeThemeColorIndex =2
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
' See "BulkChange_ChildPartControl.cls"
