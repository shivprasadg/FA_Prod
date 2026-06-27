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
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =264
    Left =1440
    Top =195
    RecSrcDt = Begin
        0x6362f321b7f9e440
    End
    RecordSource ="ATF_Report_LeaseActiveQ"
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
            KeepTogether =1
            ControlSource ="SGrpID"
        End
        Begin BreakLevel
            ControlSource ="SGrpID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin PageHeader
            Height =30
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1041
            OnPrint ="[Event Procedure]"
            Name ="GroupHeader1"
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =1
                    Left =5040
                    Top =585
                    Width =2553
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="CommBaselbl"
                    Caption ="Neg. BLCD"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5040
                    LayoutCachedTop =585
                    LayoutCachedWidth =7593
                    LayoutCachedHeight =801
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =7593
                    Top =585
                    Width =3570
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label4113"
                    Caption ="Per Diem Starts On"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7593
                    LayoutCachedTop =585
                    LayoutCachedWidth =11163
                    LayoutCachedHeight =801
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5040
                    Top =801
                    Width =2553
                    ColumnOrder =0
                    FontSize =9
                    BorderColor =14277081
                    ForeColor =0
                    Name ="SgrpCommInd"
                    ControlSource ="NegBLCD"
                    StatusBarText ="BLCD source accepted or in-service"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5040
                    LayoutCachedTop =801
                    LayoutCachedWidth =7593
                    LayoutCachedHeight =1041
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7593
                    Top =801
                    Width =3570
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =1
                    BorderColor =12632256
                    ForeColor =0
                    Name ="PerDiemStartsOn"
                    ControlSource ="PerDiemType"
                    StatusBarText ="blcd source accepted or in-service"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7593
                    LayoutCachedTop =801
                    LayoutCachedWidth =11163
                    LayoutCachedHeight =1041
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =1545
                    Top =585
                    Width =1770
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label653"
                    Caption ="Type of Exchange IT"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1545
                    LayoutCachedTop =585
                    LayoutCachedWidth =3315
                    LayoutCachedHeight =801
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1545
                    Top =801
                    Width =1770
                    FontSize =9
                    TabIndex =2
                    BorderColor =12632256
                    ForeColor =0
                    Name ="ExchangeITType"
                    ControlSource ="ExchangeITType"

                    LayoutCachedLeft =1545
                    LayoutCachedTop =801
                    LayoutCachedWidth =3315
                    LayoutCachedHeight =1041
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =360
                    Top =585
                    Width =1050
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label237"
                    Caption ="Lease Type"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =360
                    LayoutCachedTop =585
                    LayoutCachedWidth =1410
                    LayoutCachedHeight =801
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =360
                    Top =801
                    Width =1050
                    FontSize =9
                    TabIndex =3
                    BorderColor =12632256
                    ForeColor =0
                    Name ="Text238"
                    ControlSource ="LeaseType"
                    StatusBarText ="Exchange Month Not notice date"
                    ValidationRule =">=0"
                    ValidationText ="Must be >= 0"
                    DefaultValue ="0"
                    ControlTipText ="Exchange Month Not notice date"
                    ShowDatePicker =1

                    LayoutCachedLeft =360
                    LayoutCachedTop =801
                    LayoutCachedWidth =1410
                    LayoutCachedHeight =1041
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3360
                    Top =801
                    Width =780
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =4
                    BorderColor =12632256
                    ForeColor =0
                    Name ="ExchangeITMonth"
                    ControlSource ="ExchangeITMonth"
                    StatusBarText ="Exchange Month Not notice date"
                    ValidationRule =">=0"
                    ValidationText ="Must be >= 0"
                    DefaultValue ="0"
                    ControlTipText ="Exchange Month Not notice date"
                    ShowDatePicker =1

                    LayoutCachedLeft =3360
                    LayoutCachedTop =801
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =1041
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =3360
                    Top =585
                    Width =780
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label652"
                    Caption ="@ Month"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3360
                    LayoutCachedTop =585
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =801
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Width =11520
                    Height =255
                    TabIndex =5
                    Name ="ATf_Report_EquipmentActive"
                    SourceObject ="Report.ATf_Report_EquipmentActive"
                    LinkChildFields ="GroupRef"
                    LinkMasterFields ="SGrpID"

                    LayoutCachedWidth =11520
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4125
                    Top =801
                    Width =840
                    FontSize =9
                    TabIndex =6
                    BorderColor =12632256
                    ForeColor =0
                    Name ="Text249"
                    ControlSource ="BLCD"
                    StatusBarText ="Exchange Month Not notice date"
                    ValidationRule =">=0"
                    ValidationText ="Must be >= 0"
                    DefaultValue ="0"
                    ControlTipText ="Exchange Month Not notice date"
                    ShowDatePicker =1

                    LayoutCachedLeft =4125
                    LayoutCachedTop =801
                    LayoutCachedWidth =4965
                    LayoutCachedHeight =1041
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =4125
                    Top =585
                    Width =840
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label250"
                    Caption ="BLCD"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4125
                    LayoutCachedTop =585
                    LayoutCachedWidth =4965
                    LayoutCachedHeight =801
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    SpecialEffect =5
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =360
                    Top =300
                    Width =11160
                    Height =285
                    FontWeight =700
                    TabIndex =7
                    BorderColor =7500402
                    ForeColor =0
                    Name ="Text263"
                    ControlSource ="=\"Structure\""

                    LayoutCachedLeft =360
                    LayoutCachedTop =300
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =585
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
            Height =1152
            OnPrint ="[Event Procedure]"
            Name ="Detail"
            AlternateBackColor =16777215
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =3
                    Left =360
                    Top =585
                    Width =1050
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label2525"
                    Caption ="Primary"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =360
                    LayoutCachedTop =585
                    LayoutCachedWidth =1410
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    DecimalPlaces =5
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2250
                    Top =825
                    Width =1065
                    Height =14
                    FontSize =9
                    TabIndex =5
                    BorderColor =14277081
                    ForeColor =7500402
                    Name ="LeaseExtOption1LRF"
                    ControlSource ="=IIf([FirstExtendedRent]=0,Null,Round([LRFExt],7))"
                    Format ="Percent"
                    StatusBarText ="1st option lrf"
                    ValidationRule =">=0"
                    ValidationText ="Rates must be >=0"
                    DefaultValue ="0"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2250
                    LayoutCachedTop =825
                    LayoutCachedWidth =3315
                    LayoutCachedHeight =839
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
                    FontItalic = NotDefault
                    DecimalPlaces =6
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2250
                    Top =839
                    Width =1065
                    Height =14
                    FontSize =9
                    TabIndex =8
                    BorderColor =12632256
                    ForeColor =7500402
                    Name ="MTMRate"
                    ControlSource ="=IIf(Nz([LRFMTM],0)=0,Null,Round([LRFMTM],7))"
                    Format ="Percent"
                    ValidationRule =">=0"
                    ValidationText ="Must be greater than or equal to 0"
                    DefaultValue ="0"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2250
                    LayoutCachedTop =839
                    LayoutCachedWidth =3315
                    LayoutCachedHeight =853
                    RowStart =3
                    RowEnd =3
                    ColumnStart =4
                    ColumnEnd =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1410
                    Top =825
                    Width =840
                    Height =14
                    FontSize =9
                    TabIndex =4
                    BorderColor =12632256
                    ForeColor =0
                    Name ="ExtendITerm"
                    ControlSource ="=IIf(Nz([FirstExtendedRent],0)=0,Null,[ExtendITOptionTerm])"
                    StatusBarText ="1st option in months"
                    ControlTipText ="1st option in months"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1410
                    LayoutCachedTop =825
                    LayoutCachedWidth =2250
                    LayoutCachedHeight =839
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1410
                    Top =585
                    Width =840
                    FontSize =9
                    BorderColor =14277081
                    ForeColor =0
                    Name ="LeaseTermPrimary"
                    ControlSource ="=([PrimaryLeaseTerm]-[PrimaryLeaseNotice]) & \" / \" & [PrimaryLeaseNotice]"
                    StatusBarText ="Primary term in months"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =1410
                    LayoutCachedTop =585
                    LayoutCachedWidth =2250
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =6
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2250
                    Top =585
                    Width =1065
                    FontSize =9
                    TabIndex =1
                    BorderColor =14277081
                    ForeColor =7500402
                    Name ="LRFGroup"
                    ControlSource ="LRF"
                    Format ="Percent"
                    ValidationRule =">=0"
                    ValidationText ="Rates must be >=0"
                    DefaultValue ="0"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2250
                    LayoutCachedTop =585
                    LayoutCachedWidth =3315
                    LayoutCachedHeight =825
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
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3360
                    Top =585
                    Width =705
                    FontSize =9
                    TabIndex =2
                    BorderColor =14277081
                    ForeColor =0
                    Name ="LeasePymtFreq"
                    ControlSource ="LeasePymtFreq"
                    StatusBarText ="Payment Freq."
                    Tag ="Protected"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3360
                    LayoutCachedTop =585
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4125
                    Top =585
                    Width =540
                    FontSize =9
                    TabIndex =3
                    BorderColor =14277081
                    ForeColor =0
                    Name ="LeasePaidAdvArr"
                    ControlSource ="LeasePaidAdvArr"
                    StatusBarText ="Paid Advance or Arrears"
                    Tag ="Protected"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4125
                    LayoutCachedTop =585
                    LayoutCachedWidth =4665
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3360
                    Top =825
                    Width =705
                    Height =14
                    FontSize =9
                    TabIndex =6
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text177"
                    ControlSource ="=IIf([FirstExtendedRent]=0,Null,[LeasePymtFreq])"
                    StatusBarText ="Payment Freq."
                    Tag ="Protected"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3360
                    LayoutCachedTop =825
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =839
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4125
                    Top =825
                    Width =540
                    Height =14
                    FontSize =9
                    TabIndex =7
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text179"
                    ControlSource ="=IIf([FirstExtendedRent]=0,Null,[LeasePaidAdvArr])"
                    StatusBarText ="Paid Advance or Arrears"
                    Tag ="Protected"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4125
                    LayoutCachedTop =825
                    LayoutCachedWidth =4665
                    LayoutCachedHeight =839
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =5040
                    Top =345
                    Width =690
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label619"
                    Caption ="Index"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5040
                    LayoutCachedTop =345
                    LayoutCachedWidth =5730
                    LayoutCachedHeight =585
                    ColumnStart =8
                    ColumnEnd =8
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =5730
                    Top =345
                    Width =1109
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label282"
                    Caption ="Swap Date"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5730
                    LayoutCachedTop =345
                    LayoutCachedWidth =6839
                    LayoutCachedHeight =585
                    ColumnStart =9
                    ColumnEnd =9
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =6839
                    Top =345
                    Width =720
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label284"
                    Caption ="Rate"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6839
                    LayoutCachedTop =345
                    LayoutCachedWidth =7559
                    LayoutCachedHeight =585
                    ColumnStart =10
                    ColumnEnd =10
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =7559
                    Top =345
                    Width =648
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label283"
                    Caption ="Term (yr)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7559
                    LayoutCachedTop =345
                    LayoutCachedWidth =8207
                    LayoutCachedHeight =585
                    ColumnStart =11
                    ColumnEnd =11
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =8207
                    Top =345
                    Width =720
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label286"
                    Caption ="Spread"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8207
                    LayoutCachedTop =345
                    LayoutCachedWidth =8927
                    LayoutCachedHeight =585
                    ColumnStart =12
                    ColumnEnd =12
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =8927
                    Top =345
                    Width =720
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label285"
                    Caption ="Yield"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8927
                    LayoutCachedTop =345
                    LayoutCachedWidth =9647
                    LayoutCachedHeight =585
                    ColumnStart =13
                    ColumnEnd =13
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =9647
                    Top =345
                    Width =660
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label299"
                    Caption ="Bonus"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9647
                    LayoutCachedTop =345
                    LayoutCachedWidth =10307
                    LayoutCachedHeight =585
                    ColumnStart =14
                    ColumnEnd =14
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =10307
                    Top =345
                    Width =825
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label347"
                    Caption ="Adjmnt"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10307
                    LayoutCachedTop =345
                    LayoutCachedWidth =11132
                    LayoutCachedHeight =585
                    ColumnStart =15
                    ColumnEnd =15
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5730
                    Top =585
                    Width =1109
                    FontSize =9
                    TabIndex =9
                    BorderColor =12566463
                    ForeColor =0
                    Name ="SwapDate"
                    ControlSource ="SwapDate"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =5730
                    LayoutCachedTop =585
                    LayoutCachedWidth =6839
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7559
                    Top =585
                    Width =648
                    FontSize =9
                    TabIndex =10
                    BorderColor =12566463
                    ForeColor =0
                    Name ="SwapTermYears"
                    ControlSource ="SwapTermYears"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =7559
                    LayoutCachedTop =585
                    LayoutCachedWidth =8207
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =11
                    ColumnEnd =11
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6839
                    Top =585
                    Width =720
                    FontSize =9
                    TabIndex =11
                    BorderColor =12566463
                    ForeColor =0
                    Name ="SwapRate"
                    ControlSource ="SwapRate"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =6839
                    LayoutCachedTop =585
                    LayoutCachedWidth =7559
                    LayoutCachedHeight =825
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
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8207
                    Top =585
                    Width =720
                    FontSize =9
                    TabIndex =12
                    BorderColor =14277081
                    ForeColor =0
                    Name ="SalesSpread"
                    ControlSource ="=[SwapYield]-[SwapRate]"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000000 ,
                        0x00000000ffc20e00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8207
                    LayoutCachedTop =585
                    LayoutCachedWidth =8927
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =12
                    ColumnEnd =12
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x01000100000000000000050000000100000000000000ffc20e00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8927
                    Top =585
                    Width =720
                    FontSize =9
                    TabIndex =13
                    BorderColor =12566463
                    ForeColor =0
                    Name ="SwapYield"
                    ControlSource ="SwapYield"
                    Format ="Percent"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =8927
                    LayoutCachedTop =585
                    LayoutCachedWidth =9647
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =13
                    ColumnEnd =13
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =5
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10307
                    Top =585
                    Width =806
                    FontSize =8
                    TabIndex =14
                    BorderColor =12566463
                    ForeColor =0
                    Name ="SwapAdjustment"
                    ControlSource ="SwapAdjustment"
                    Format ="Percent"
                    StatusBarText ="downadj"
                    FontName ="Segoe UI"
                    Tag ="TrackUpdate"
                    ControlTipText ="Rate Adjustment - If Downard"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =10307
                    LayoutCachedTop =585
                    LayoutCachedWidth =11113
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =15
                    ColumnEnd =15
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
                    Left =5040
                    Top =585
                    Width =690
                    FontSize =9
                    TabIndex =15
                    BorderColor =12566463
                    ForeColor =0
                    Name ="SwapIndex"
                    ControlSource ="SwapIndex"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5040
                    LayoutCachedTop =585
                    LayoutCachedWidth =5730
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9647
                    Top =585
                    Width =660
                    FontSize =9
                    TabIndex =16
                    BorderColor =12566463
                    ForeColor =0
                    Name ="SwapBonus"
                    ControlSource ="SwapBonus"
                    StatusBarText ="Bonus Depreciation"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9647
                    LayoutCachedTop =585
                    LayoutCachedWidth =10307
                    LayoutCachedHeight =825
                    RowStart =1
                    RowEnd =1
                    ColumnStart =14
                    ColumnEnd =14
                    BackThemeColorIndex =-1
                    BorderShade =75.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =1395
                    Top =345
                    Width =855
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="COmmLBL"
                    Caption ="Term "
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1395
                    LayoutCachedTop =345
                    LayoutCachedWidth =2250
                    LayoutCachedHeight =585
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =2250
                    Top =345
                    Width =1065
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="LRFLBL"
                    Caption ="Actual LRF"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2250
                    LayoutCachedTop =345
                    LayoutCachedWidth =3315
                    LayoutCachedHeight =585
                    ColumnStart =4
                    ColumnEnd =4
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =3360
                    Top =345
                    Width =705
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="PisLBL"
                    Caption ="Freq"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3360
                    LayoutCachedTop =345
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =585
                    ColumnStart =5
                    ColumnEnd =5
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =4125
                    Top =345
                    Width =540
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="MdeinLBL"
                    Caption ="Paid in"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4125
                    LayoutCachedTop =345
                    LayoutCachedWidth =4665
                    LayoutCachedHeight =585
                    ColumnStart =6
                    ColumnEnd =6
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =360
                    Top =839
                    Width =1035
                    Height =14
                    FontSize =8
                    TabIndex =17
                    BorderColor =16777215
                    Name ="MTMLabe"
                    ControlSource ="=IIf([MTMRent]=0,Null,\"MTM / LRF\")"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =360
                    LayoutCachedTop =839
                    LayoutCachedWidth =1395
                    LayoutCachedHeight =853
                    RowStart =3
                    RowEnd =3
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Subform
                    Locked = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =360
                    Top =853
                    Width =11160
                    Height =299
                    TabIndex =18
                    Name ="ATF_Report_AllowancesActive"
                    SourceObject ="Report.ATF_Report_AllowancesActive"
                    LinkChildFields ="sGrpId"
                    LinkMasterFields ="sGrpId"

                    LayoutCachedLeft =360
                    LayoutCachedTop =853
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =1152
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =360
                    Top =825
                    Width =1050
                    Height =14
                    FontSize =8
                    TabIndex =19
                    BorderColor =16777215
                    Name ="ExtendedLabel"
                    ControlSource ="=IIf([FirstExtendedRent]=0,Null,\"Extended\")"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =360
                    LayoutCachedTop =825
                    LayoutCachedWidth =1410
                    LayoutCachedHeight =839
                    RowStart =2
                    RowEnd =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    SpecialEffect =5
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =360
                    Top =60
                    Width =11160
                    Height =285
                    FontWeight =700
                    TabIndex =20
                    BorderColor =7500402
                    ForeColor =0
                    Name ="Text262"
                    ControlSource ="=\"Terms\""

                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =345
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =0
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "ATF_Report_StructureActive.cls"
