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
    Width =11085
    DatasheetFontHeight =11
    ItemSuffix =269
    Left =2190
    Top =3990
    RecSrcDt = Begin
        0x6c5409b314f6e440
    End
    RecordSource ="ATF_Report_AllowancesActiveQ"
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
            ControlSource ="ATFGroupSpecID"
        End
        Begin BreakLevel
            ControlSource ="ATFGroupSpecID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =285
            OnPrint ="[Event Procedure]"
            Name ="GroupHeader0"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    SpecialEffect =5
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Width =6240
                    Height =285
                    FontWeight =700
                    BorderColor =7500402
                    ForeColor =0
                    Name ="Text262"
                    ControlSource ="=\"Fees\""

                    LayoutCachedWidth =6240
                    LayoutCachedHeight =285
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    SpecialEffect =5
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5265
                    Width =5820
                    Height =285
                    FontWeight =700
                    TabIndex =1
                    BorderColor =7500402
                    ForeColor =0
                    Name ="Text264"
                    ControlSource ="=\"Allowances\""

                    LayoutCachedLeft =5265
                    LayoutCachedWidth =11085
                    LayoutCachedHeight =285
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =568
            OnPrint ="[Event Procedure]"
            Name ="Detail"
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =8469
                    Top =210
                    Width =864
                    FontSize =9
                    BorderColor =14277081
                    ForeColor =0
                    Name ="CPMileThreshold"
                    ControlSource ="CPMileThreshold"
                    Format ="Standard"
                    StatusBarText ="charge exceed allowance 1"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    ShowDatePicker =1

                    LayoutCachedLeft =8469
                    LayoutCachedTop =210
                    LayoutCachedWidth =9333
                    LayoutCachedHeight =450
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
                    DecimalPlaces =4
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9318
                    Top =525
                    Width =765
                    Height =43
                    FontSize =9
                    TabIndex =1
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text517"
                    ControlSource ="=IIf(Nz([HoursAllowed],0)=0,Null,[CPHourAfter])"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Per Hour Charge After Threshold Allowance"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Per Hour Charge After Threshold Allowance"
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =9318
                    LayoutCachedTop =525
                    LayoutCachedWidth =10083
                    LayoutCachedHeight =568
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin Label
                    TextAlign =2
                    Left =9333
                    Width =765
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label515"
                    Caption ="Chrg After"
                    LayoutCachedLeft =9333
                    LayoutCachedWidth =10098
                    LayoutCachedHeight =216
                    ColumnStart =5
                    ColumnEnd =5
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =4
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9333
                    Top =210
                    Width =765
                    FontSize =9
                    TabIndex =2
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text516"
                    ControlSource ="CPMileAfter"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Per Mile Charge After Threshold Allowance"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Per Mile Charge After Threshold Allowance"
                    ShowDatePicker =1

                    LayoutCachedLeft =9333
                    LayoutCachedTop =210
                    LayoutCachedWidth =10098
                    LayoutCachedHeight =450
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin Label
                    TextAlign =2
                    Left =8469
                    Width =864
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label506"
                    Caption ="Threshold"
                    LayoutCachedLeft =8469
                    LayoutCachedWidth =9333
                    LayoutCachedHeight =216
                    ColumnStart =4
                    ColumnEnd =4
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =8454
                    Top =525
                    Width =864
                    Height =43
                    FontSize =9
                    TabIndex =3
                    BorderColor =14277081
                    ForeColor =0
                    Name ="HourThreshold"
                    ControlSource ="=IIf(Nz([HoursAllowed],0)=0,Null,[CPHourThreshold])"
                    Format ="Standard"
                    StatusBarText ="Per Hour Charge After Threshold Allowance"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8454
                    LayoutCachedTop =525
                    LayoutCachedWidth =9318
                    LayoutCachedHeight =568
                    RowStart =2
                    RowEnd =2
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =6870
                    Top =525
                    Width =804
                    Height =43
                    FontSize =9
                    TabIndex =4
                    BorderColor =14277081
                    ForeColor =0
                    Name ="HoursAllowed"
                    ControlSource ="=IIf(Nz([HPYAllowance],0)=0,Null,[HPYAllowance])"
                    Format ="Standard"
                    StatusBarText ="Hours Allowed per Method (Term/Year)"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Hours Allowed per Method (Term/Year)"
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =6870
                    LayoutCachedTop =525
                    LayoutCachedWidth =7674
                    LayoutCachedHeight =568
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =4
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7674
                    Top =525
                    Width =780
                    Height =43
                    FontSize =9
                    TabIndex =5
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text496"
                    ControlSource ="=IIf(Nz([HoursAllowed],0)=0,Null,[CPHour])"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Per Hour Charge Before Threshold Allowance"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Per Hour Charge Before Threshold Allowance"
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =7674
                    LayoutCachedTop =525
                    LayoutCachedWidth =8454
                    LayoutCachedHeight =568
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5325
                    Top =216
                    Width =945
                    FontSize =9
                    TabIndex =6
                    BorderColor =14277081
                    ForeColor =0
                    Name ="AllowedMilesPer"
                    ControlSource ="AllowedMilesPer"
                    RowSourceType ="Value List"
                    RowSource ="P/YEAR;P/TERM"

                    LayoutCachedLeft =5325
                    LayoutCachedTop =216
                    LayoutCachedWidth =6270
                    LayoutCachedHeight =456
                    RowStart =1
                    RowEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =5325
                    Width =945
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label95"
                    Caption ="Allowed Per"
                    LayoutCachedLeft =5325
                    LayoutCachedWidth =6270
                    LayoutCachedHeight =216
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =7689
                    Width =780
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label390"
                    Caption ="Chrg Per"
                    LayoutCachedLeft =7689
                    LayoutCachedWidth =8469
                    LayoutCachedHeight =216
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =3
                    Left =6285
                    Top =210
                    Width =585
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label391"
                    Caption ="Miles:"
                    LayoutCachedLeft =6285
                    LayoutCachedTop =210
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =450
                    ColumnStart =2
                    ColumnEnd =2
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =4
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7689
                    Top =210
                    Width =780
                    FontSize =9
                    TabIndex =7
                    BorderColor =14277081
                    ForeColor =0
                    Name ="SGprAllow1"
                    ControlSource ="CPMile"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Per Mile Charge Before Threshold Allowance"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Per Mile Charge Before Threshold Allowance"
                    ShowDatePicker =1

                    LayoutCachedLeft =7689
                    LayoutCachedTop =210
                    LayoutCachedWidth =8469
                    LayoutCachedHeight =450
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =6885
                    Top =210
                    Width =804
                    FontSize =9
                    TabIndex =8
                    BorderColor =14277081
                    ForeColor =0
                    Name ="MPYAllowance"
                    ControlSource ="MPYAllowance"
                    Format ="Standard"
                    ShowDatePicker =1

                    LayoutCachedLeft =6885
                    LayoutCachedTop =210
                    LayoutCachedWidth =7689
                    LayoutCachedHeight =450
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =6825
                    Width =864
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label245"
                    Caption ="Allowed"
                    LayoutCachedLeft =6825
                    LayoutCachedWidth =7689
                    LayoutCachedHeight =210
                    ColumnStart =2
                    ColumnEnd =2
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6285
                    Top =525
                    Width =585
                    Height =43
                    FontSize =8
                    TabIndex =9
                    BorderColor =16777215
                    Name ="Label688"
                    ControlSource ="=IIf(Nz([HoursAllowed],0)=0,Null,\"Hours: \")"

                    LayoutCachedLeft =6285
                    LayoutCachedTop =525
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =568
                    ColumnStart =2
                    ColumnEnd =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =2235
                    Width =432
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label642"
                    Caption ="Start"
                    LayoutCachedLeft =2235
                    LayoutCachedWidth =2667
                    LayoutCachedHeight =216
                    ColumnStart =2
                    ColumnEnd =2
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =1395
                    Width =720
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label643"
                    Caption ="Amount"
                    LayoutCachedLeft =1395
                    LayoutCachedWidth =2115
                    LayoutCachedHeight =210
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =4
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2235
                    Top =216
                    Width =432
                    FontSize =9
                    TabIndex =10
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text644"
                    ControlSource ="=IIf(Nz([SGrpFSLPmt],0)=0,Null,[FMSFeeStartMonth])"
                    Format ="General Number"
                    StatusBarText ="Start Month of the Servce Fee (Zero indicates full term of lease)"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Start Month of the Servce Fee (Zero indicates full term of lease)"
                    ShowDatePicker =1

                    LayoutCachedLeft =2235
                    LayoutCachedTop =216
                    LayoutCachedWidth =2667
                    LayoutCachedHeight =456
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =4
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2235
                    Top =480
                    Width =432
                    Height =43
                    FontSize =9
                    TabIndex =11
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text647"
                    ControlSource ="=IIf(Nz([Fee2Amount],0)=0,Null,[Fee2StartMonth])"
                    Format ="General Number"
                    StatusBarText ="Start Month of the 2nd Fee (Zero indicates full term of lease)"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Start Month of the 2nd Fee (Zero indicates full term of lease)"
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2235
                    LayoutCachedTop =480
                    LayoutCachedWidth =2667
                    LayoutCachedHeight =523
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1395
                    Top =480
                    Width =720
                    Height =43
                    FontSize =9
                    TabIndex =12
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text648"
                    ControlSource ="=IIf(Nz([Fee2Amount],0)=0,Null,[Fee2Amount])"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Amount of the 2nd Fee (if any)"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Amount of the 2nd Fee (if any)"
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1395
                    LayoutCachedTop =480
                    LayoutCachedWidth =2115
                    LayoutCachedHeight =523
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin Label
                    TextAlign =1
                    Left =2775
                    Width =432
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label649"
                    Caption ="Term"
                    LayoutCachedLeft =2775
                    LayoutCachedWidth =3207
                    LayoutCachedHeight =216
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2775
                    Top =480
                    Width =432
                    Height =43
                    FontSize =9
                    TabIndex =13
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text650"
                    ControlSource ="=IIf(Nz([Fee2Amount],0)=0,Null,[Fee2TermMonths])"
                    Format ="General Number"
                    StatusBarText ="Term in Months of the 2nd Fee (Zero indicates full term of lease)"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Term in Months of the 2nd Fee (Zero indicates full term of lease)"
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2775
                    LayoutCachedTop =480
                    LayoutCachedWidth =3207
                    LayoutCachedHeight =523
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2775
                    Top =216
                    Width =432
                    FontSize =9
                    TabIndex =14
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text651"
                    ControlSource ="=IIf(Nz([SGrpFSLPmt],0)=0,Null,[ScheduleGroupPaymentTerm])"
                    Format ="General Number"
                    StatusBarText ="Term in Months of the Servce Fee (if any)"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Term in Months of the Servce Fee (if any)"
                    ShowDatePicker =1

                    LayoutCachedLeft =2775
                    LayoutCachedTop =216
                    LayoutCachedWidth =3207
                    LayoutCachedHeight =456
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1395
                    Top =216
                    Width =720
                    FontSize =9
                    TabIndex =15
                    BorderColor =12566463
                    ForeColor =0
                    Name ="Text653"
                    ControlSource ="SGrpFSLPmt"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShortcutMenuBar ="Amount of the Servce Fee (if any)"
                    ControlTipText ="Amount of the Servce Fee (if any)"
                    ShowDatePicker =1

                    LayoutCachedLeft =1395
                    LayoutCachedTop =216
                    LayoutCachedWidth =2115
                    LayoutCachedHeight =456
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =375
                    Top =480
                    Width =990
                    Height =43
                    FontSize =8
                    TabIndex =16
                    BorderColor =16777215
                    Name ="Fee2TypeName"
                    ControlSource ="=IIf(Nz([Fee2Amount],0)=0,Null,[Fee2Type])"
                    Format ="Standard"
                    StatusBarText ="Type/Name of the 2nd Fee (if any)"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Type/Name of the 2nd Fee (if any)"
                    TopPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =375
                    LayoutCachedTop =480
                    LayoutCachedWidth =1365
                    LayoutCachedHeight =523
                    RowStart =2
                    RowEnd =2
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =3
                    Left =375
                    Top =216
                    Width =990
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label248"
                    Caption ="FMSS Fee"
                    LayoutCachedLeft =375
                    LayoutCachedTop =216
                    LayoutCachedWidth =1365
                    LayoutCachedHeight =456
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    Visible = NotDefault
                    Left =3870
                    Width =576
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label246"
                    Caption ="  Residual (Avg)"
                    LayoutCachedLeft =3870
                    LayoutCachedWidth =4446
                    LayoutCachedHeight =216
                    ColumnStart =5
                    ColumnEnd =5
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3870
                    Top =216
                    Width =576
                    FontSize =9
                    TabIndex =17
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text247"
                    ControlSource ="AvgResidual"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Per Mile Charge After Threshold Allowance"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Per Mile Charge After Threshold Allowance"
                    ShowDatePicker =1

                    LayoutCachedLeft =3870
                    LayoutCachedTop =216
                    LayoutCachedWidth =4446
                    LayoutCachedHeight =456
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =3
                    Left =4446
                    Width =576
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label249"
                    Caption ="UTA Rebate"
                    LayoutCachedLeft =4446
                    LayoutCachedWidth =5022
                    LayoutCachedHeight =210
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4446
                    Top =216
                    Width =576
                    FontSize =9
                    TabIndex =18
                    BorderColor =12566463
                    ForeColor =0
                    Name ="UTARebate"
                    ControlSource ="UTARebate"
                    Format ="$#,##0.00;($#,##0.00)"
                    ShortcutMenuBar ="Amount of the Servce Fee (if any)"
                    ControlTipText ="Amount of the Servce Fee (if any)"
                    ShowDatePicker =1

                    LayoutCachedLeft =4446
                    LayoutCachedTop =216
                    LayoutCachedWidth =5022
                    LayoutCachedHeight =456
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =8469
                    Top =450
                    Width =864
                    Height =43
                    FontSize =9
                    TabIndex =19
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text265"
                    ControlSource ="=IIf(Nz([CPMileThreshold2],0)=0,Null,[CPMileThreshold2])"
                    Format ="Standard"
                    StatusBarText ="charge exceed allowance 1"
                    ValidationRule =">=0"
                    ValidationText ="Over use amounts must be >=0"
                    DefaultValue ="0"
                    ShowDatePicker =1

                    LayoutCachedLeft =8469
                    LayoutCachedTop =450
                    LayoutCachedWidth =9333
                    LayoutCachedHeight =493
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
                    DecimalPlaces =4
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9333
                    Top =450
                    Width =765
                    Height =43
                    FontSize =9
                    TabIndex =20
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text266"
                    ControlSource ="=IIf(Nz([CPMileAfter2],0)=0,Null,[CPMileAfter2])"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Per Mile Charge After Threshold Allowance"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Per Mile Charge After Threshold Allowance"
                    ShowDatePicker =1

                    LayoutCachedLeft =9333
                    LayoutCachedTop =450
                    LayoutCachedWidth =10098
                    LayoutCachedHeight =493
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =4
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7689
                    Top =450
                    Width =780
                    Height =43
                    FontSize =8
                    TabIndex =21
                    BorderColor =16777215
                    Name ="MilesTier2"
                    ControlSource ="=IIf(Nz([CPMileAfter2],0)=0,Null,\"Tier 3 >>\")"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Per Mile Charge Before Threshold Allowance"
                    ValidationRule =">=0"
                    ValidationText ="Allowances must be >=0"
                    DefaultValue ="0"
                    ControlTipText ="Per Mile Charge Before Threshold Allowance"
                    ShowDatePicker =1

                    LayoutCachedLeft =7689
                    LayoutCachedTop =450
                    LayoutCachedWidth =8469
                    LayoutCachedHeight =493
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "ATF_Report_AllowancesActive.cls"
