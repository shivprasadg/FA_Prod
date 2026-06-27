Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11630
    DatasheetFontHeight =11
    ItemSuffix =324
    Left =1830
    Top =1845
    RecSrcDt = Begin
        0x648207436cf9e440
    End
    RecordSource ="ATF_Report_EquipmentActiveQ"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
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
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
            ShowDatePicker =0
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
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="UnitGroup"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="UnitDesc"
        End
        Begin BreakLevel
            ControlSource ="TotalCost"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="SchRent"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =540
            OnPrint ="[Event Procedure]"
            Name ="GroupHeader2"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5505
                    Top =330
                    Width =288
                    Height =150
                    ColumnOrder =2
                    FontSize =9
                    BorderColor =14277081
                    ForeColor =0
                    Name ="GroupsIn"
                    ControlSource ="GroupsIn"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5505
                    LayoutCachedTop =330
                    LayoutCachedWidth =5793
                    LayoutCachedHeight =480
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5925
                    Top =330
                    Width =288
                    Height =150
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =1
                    BorderColor =14277081
                    ForeColor =0
                    Name ="GroupId"
                    ControlSource ="GroupRef"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5925
                    LayoutCachedTop =330
                    LayoutCachedWidth =6213
                    LayoutCachedHeight =480
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    TextAlign =1
                    Left =120
                    Top =330
                    Width =420
                    Height =194
                    FontSize =8
                    BorderColor =16777215
                    Name ="UnitsLabel"
                    Caption ="Units"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedTop =330
                    LayoutCachedWidth =540
                    LayoutCachedHeight =524
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =6210
                    Top =330
                    Width =1260
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label284"
                    Caption ="Vendor Cost"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6210
                    LayoutCachedTop =330
                    LayoutCachedWidth =7470
                    LayoutCachedHeight =540
                    ColumnStart =10
                    ColumnEnd =10
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =8679
                    Top =330
                    Width =945
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label282"
                    Caption ="Sch/Ext Rent"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8679
                    LayoutCachedTop =330
                    LayoutCachedWidth =9624
                    LayoutCachedHeight =540
                    ColumnStart =9
                    ColumnEnd =9
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    TextAlign =1
                    Left =616
                    Top =330
                    Width =4320
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="COmmLBL"
                    Caption ="Description / Vendor"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =616
                    LayoutCachedTop =330
                    LayoutCachedWidth =4936
                    LayoutCachedHeight =540
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =10815
                    Top =330
                    Width =806
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label283"
                    Caption ="UTA "
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10815
                    LayoutCachedTop =330
                    LayoutCachedWidth =11621
                    LayoutCachedHeight =540
                    ColumnStart =10
                    ColumnEnd =10
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =9591
                    Top =330
                    Width =1125
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label288"
                    Caption ="Sch RV"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9591
                    LayoutCachedTop =330
                    LayoutCachedWidth =10716
                    LayoutCachedHeight =540
                    ColumnStart =9
                    ColumnEnd =9
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Width =11630
                    Height =330
                    ColumnOrder =0
                    FontWeight =700
                    TabIndex =2
                    BackColor =9605778
                    Name ="Text245"
                    ControlSource ="=\"Group: \" & [UnitGroup] & \"  \" & [GrpDesc]"

                    LayoutCachedWidth =11630
                    LayoutCachedHeight =330
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =7575
                    Top =330
                    Width =1095
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label319"
                    Caption ="Sch Cost"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7575
                    LayoutCachedTop =330
                    LayoutCachedWidth =8670
                    LayoutCachedHeight =540
                    ColumnStart =10
                    ColumnEnd =10
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            BreakLevel =1
            OnPrint ="[Event Procedure]"
            Name ="GroupHeader1"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =753
            BreakLevel =3
            OnPrint ="[Event Procedure]"
            Name ="GroupHeader0"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    Left =119
                    Top =51
                    Width =420
                    Height =314
                    BorderColor =9605778
                    Name ="Box313"
                    LayoutCachedLeft =119
                    LayoutCachedTop =51
                    LayoutCachedWidth =539
                    LayoutCachedHeight =365
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6090
                    Width =1365
                    Height =270
                    FontSize =9
                    TabIndex =2
                    BorderColor =12566463
                    ForeColor =0
                    Name ="VendorCostAvg"
                    ControlSource ="=([VendorCost])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =6090
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =270
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8715
                    Width =909
                    Height =270
                    FontSize =9
                    TabIndex =1
                    BorderColor =12566463
                    ForeColor =0
                    Name ="EstimatedDeliveryDate"
                    ControlSource ="SchRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8715
                    LayoutCachedWidth =9624
                    LayoutCachedHeight =270
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =69
                    Width =405
                    Height =270
                    FontSize =9
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =0
                    Name ="UnitCnt"
                    ControlSource ="=Count([UnitID])"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =120
                    LayoutCachedTop =69
                    LayoutCachedWidth =525
                    LayoutCachedHeight =339
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =631
                    Width =5646
                    Height =270
                    FontSize =9
                    TabIndex =3
                    BorderColor =12566463
                    ForeColor =0
                    Name ="UnitDescLine"
                    ControlSource ="[UnitDesc]"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =631
                    LayoutCachedWidth =6277
                    LayoutCachedHeight =270
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10560
                    Width =1061
                    Height =270
                    FontSize =9
                    TabIndex =4
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text282"
                    ControlSource ="UTA"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =10560
                    LayoutCachedWidth =11621
                    LayoutCachedHeight =270
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin Subform
                    Locked = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =135
                    Top =465
                    Width =11478
                    Height =288
                    TabIndex =6
                    Name ="ATF_Report_ChildPartsActive"
                    SourceObject ="Report.ATF_Report_ChildPartsActive"
                    LinkChildFields ="SpecProfile"
                    LinkMasterFields ="SpecProfile"

                    LayoutCachedLeft =135
                    LayoutCachedTop =465
                    LayoutCachedWidth =11613
                    LayoutCachedHeight =753
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9591
                    Width =1125
                    Height =270
                    FontSize =9
                    TabIndex =5
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text289"
                    ControlSource ="ResidualSch"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =9591
                    LayoutCachedWidth =10716
                    LayoutCachedHeight =270
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4710
                    Top =225
                    Width =749
                    Height =225
                    FontSize =8
                    TabIndex =7
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text292"
                    ControlSource ="EstimatedDeliveryDate"
                    Format ="mm/dd/yy"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =4710
                    LayoutCachedTop =225
                    LayoutCachedWidth =5459
                    LayoutCachedHeight =450
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =631
                    Top =225
                    Width =3345
                    Height =225
                    FontSize =8
                    TabIndex =8
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text310"
                    ControlSource ="VendorSNm"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =631
                    LayoutCachedTop =225
                    LayoutCachedWidth =3976
                    LayoutCachedHeight =450
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8715
                    Top =225
                    Width =909
                    Height =270
                    FontSize =9
                    TabIndex =9
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text312"
                    ControlSource ="FirstExtRent"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8715
                    LayoutCachedTop =225
                    LayoutCachedWidth =9624
                    LayoutCachedHeight =495
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =3930
                    Top =240
                    Width =765
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label317"
                    Caption ="Est. Delv"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3930
                    LayoutCachedTop =240
                    LayoutCachedWidth =4695
                    LayoutCachedHeight =450
                    ColumnStart =9
                    ColumnEnd =9
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7305
                    Width =1365
                    Height =270
                    FontSize =9
                    TabIndex =10
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Cost"
                    ControlSource ="[UnitCost]"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =7305
                    LayoutCachedWidth =8670
                    LayoutCachedHeight =270
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7476
                    Top =225
                    Width =1194
                    Height =270
                    FontSize =9
                    TabIndex =11
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text322"
                    ControlSource ="=[Partcost]"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =7476
                    LayoutCachedTop =225
                    LayoutCachedWidth =8670
                    LayoutCachedHeight =495
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6261
                    Top =225
                    Width =1194
                    Height =270
                    FontSize =9
                    TabIndex =12
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text323"
                    ControlSource ="=[Partcost]"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =6261
                    LayoutCachedTop =225
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =495
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="Detail"
            AlternateBackColor =16777215
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Width =840
                    Height =0
                    FontSize =8
                    BorderColor =12566463
                    ForeColor =0
                    Name ="SpecProfile"
                    ControlSource ="SpecProfile"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedWidth =840
                    RowStart =1
                    RowEnd =1
                    ColumnStart =10
                    ColumnEnd =10
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =30
            BreakLevel =3
            OnPrint ="[Event Procedure]"
            Name ="GroupFooter0"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7305
                    Width =1364
                    Height =30
                    FontSize =9
                    TabIndex =1
                    BorderColor =8355711
                    ForeColor =0
                    Name ="TotalOEC"
                    ControlSource ="=IIf([PartCost]>0,[UnitCost]+[PartCost],Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =7305
                    LayoutCachedWidth =8669
                    LayoutCachedHeight =30
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    Visible = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6375
                    Width =1275
                    Height =14
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label285"
                    ControlSource ="=IIf([TotalCost]-[UnitCost]<1,Null,\"Total Unit Cost:\")"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6375
                    LayoutCachedWidth =7650
                    LayoutCachedHeight =14
                    ColumnStart =10
                    ColumnEnd =10
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6091
                    Width =1364
                    Height =30
                    FontSize =9
                    TabIndex =2
                    BorderColor =8355711
                    ForeColor =0
                    Name ="VendorOec"
                    ControlSource ="=IIf([PartCost]>0,[VendorCost]+[PartCost],Null)"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =6091
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =30
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =300
            OnPrint ="[Event Procedure]"
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8535
                    Top =29
                    Width =1089
                    Height =270
                    FontSize =9
                    BorderColor =12566463
                    ForeColor =0
                    Name ="GroupRent"
                    ControlSource ="=IIf([Groupsin]<1,Null,Sum([SchRent]))"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8535
                    LayoutCachedTop =29
                    LayoutCachedWidth =9624
                    LayoutCachedHeight =299
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =120
                    Width =420
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    BackColor =8355711
                    BorderColor =8355711
                    Name ="Units"
                    ControlSource ="=IIf([Groupsin]<2,Null,Count([UnitID]))"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =120
                    LayoutCachedWidth =540
                    LayoutCachedHeight =270
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9501
                    Top =29
                    Width =1215
                    Height =270
                    FontSize =9
                    TabIndex =2
                    BorderColor =12566463
                    ForeColor =0
                    Name ="GroupRV"
                    ControlSource ="=IIf([Groupsin]<1,Null,Sum([ResidualSch]))"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =9501
                    LayoutCachedTop =29
                    LayoutCachedWidth =10716
                    LayoutCachedHeight =299
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7305
                    Top =30
                    Width =1350
                    Height =270
                    FontSize =9
                    TabIndex =3
                    BorderColor =12566463
                    ForeColor =0
                    Name ="GroupOec"
                    ControlSource ="=IIf([Groupsin]<1,Null,Sum([UnitCost]+[PartCost]))"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =7305
                    LayoutCachedTop =30
                    LayoutCachedWidth =8655
                    LayoutCachedHeight =300
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =645
                    Top =30
                    Width =3111
                    Height =270
                    FontSize =8
                    TabIndex =4
                    BorderColor =16777215
                    Name ="Label304"
                    ControlSource ="=IIf([Groupsin]<1,Null,\"Group:<b> \" & [UnitGroup] & \"</b> Sub Total\")"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =645
                    LayoutCachedTop =30
                    LayoutCachedWidth =3756
                    LayoutCachedHeight =300
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10569
                    Top =29
                    Width =1061
                    Height =270
                    FontSize =9
                    TabIndex =5
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text305"
                    ControlSource ="=IIf([Groupsin]<1,Null,Sum([UTA]))"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =10569
                    LayoutCachedTop =29
                    LayoutCachedWidth =11630
                    LayoutCachedHeight =299
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3930
                    Top =30
                    Width =1086
                    Height =270
                    FontSize =8
                    TabIndex =6
                    BorderColor =16777215
                    Name ="LRF"
                    ControlSource ="=IIf([GroupsIn]<1,Null,\"Group LRF: \")"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =3930
                    LayoutCachedTop =30
                    LayoutCachedWidth =5016
                    LayoutCachedHeight =300
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =6
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5025
                    Top =30
                    Width =1035
                    Height =270
                    FontSize =9
                    TabIndex =7
                    BorderColor =14277081
                    ForeColor =0
                    Name ="LRFGroup"
                    ControlSource ="=IIf([GroupsIn]<1,Null,([GroupRent]/[GroupOEC]))"
                    Format ="Percent"
                    ValidationRule =">=0"
                    ValidationText ="Rates must be >=0"
                    DefaultValue ="0"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =5025
                    LayoutCachedTop =30
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =300
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5985
                    Top =30
                    Width =1470
                    Height =270
                    FontSize =9
                    TabIndex =8
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text320"
                    ControlSource ="=IIf([Groupsin]<1,Null,Sum([VendorCost]+[PartCost]))"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =5985
                    LayoutCachedTop =30
                    LayoutCachedWidth =7455
                    LayoutCachedHeight =300
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
            End
        End
    End
End
CodeBehindForm
' See "ATf_Report_EquipmentActive.cls"
