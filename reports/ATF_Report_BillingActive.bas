Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11535
    DatasheetFontHeight =11
    ItemSuffix =331
    Left =3300
    Top =2610
    RecSrcDt = Begin
        0x193ba69c15f6e440
    End
    RecordSource ="ATF_Report_BillingActiveQ"
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
            ControlSource ="SchID"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =552
            Name ="GroupHeader0"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    SpecialEffect =5
                    OldBorderStyle =1
                    TextAlign =1
                    Left =360
                    Width =9048
                    Height =330
                    FontWeight =700
                    BorderColor =7500402
                    ForeColor =0
                    Name ="Label15"
                    Caption ="Billing "
                    LayoutCachedLeft =360
                    LayoutCachedWidth =9408
                    LayoutCachedHeight =330
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =420
                    Top =336
                    Width =2445
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label247"
                    Caption ="BillTo"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =420
                    LayoutCachedTop =336
                    LayoutCachedWidth =2865
                    LayoutCachedHeight =552
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    Left =2868
                    Top =336
                    Width =636
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label248"
                    Caption ="Method"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2868
                    LayoutCachedTop =336
                    LayoutCachedWidth =3504
                    LayoutCachedHeight =552
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    Left =3504
                    Top =336
                    Width =648
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label27"
                    Caption ="Servicer"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3504
                    LayoutCachedTop =336
                    LayoutCachedWidth =4152
                    LayoutCachedHeight =552
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =1
                    Left =9252
                    Top =336
                    Width =2232
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label269"
                    Caption ="Tax Instruction"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9252
                    LayoutCachedTop =336
                    LayoutCachedWidth =11484
                    LayoutCachedHeight =552
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =1
                    Left =6240
                    Top =336
                    Width =3012
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label289"
                    Caption ="Billing Instructions"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6240
                    LayoutCachedTop =336
                    LayoutCachedWidth =9252
                    LayoutCachedHeight =552
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    Left =4152
                    Top =336
                    Width =648
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label296"
                    Caption ="Invoicer"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4152
                    LayoutCachedTop =336
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =552
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    Left =4800
                    Top =336
                    Width =708
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label304"
                    Caption ="Collector"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4800
                    LayoutCachedTop =336
                    LayoutCachedWidth =5508
                    LayoutCachedHeight =552
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    Left =5508
                    Top =336
                    Width =732
                    Height =216
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label314"
                    Caption ="3rd Party"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5508
                    LayoutCachedTop =336
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =552
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    Left =3600
                    Top =120
                    Width =2475
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label328"
                    Caption ="┌──────── FA Invoicing ───────┐"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3600
                    LayoutCachedTop =120
                    LayoutCachedWidth =6075
                    LayoutCachedHeight =330
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    SpecialEffect =5
                    OldBorderStyle =1
                    TextAlign =1
                    Left =9408
                    Width =2058
                    Height =330
                    FontWeight =700
                    BorderColor =7500402
                    ForeColor =0
                    Name ="Label330"
                    Caption =" Taxes"
                    LayoutCachedLeft =9408
                    LayoutCachedWidth =11466
                    LayoutCachedHeight =330
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =285
            Name ="Detail"
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =420
                    Width =2445
                    Height =230
                    FontSize =8
                    BorderColor =14211288
                    ForeColor =0
                    Name ="BillingAddress"
                    ControlSource ="BillingAddress"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =420
                    LayoutCachedWidth =2865
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2868
                    Width =636
                    Height =230
                    FontSize =8
                    TabIndex =1
                    BorderColor =14211288
                    ForeColor =0
                    Name ="Method"
                    ControlSource ="PayMethod"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =2868
                    LayoutCachedWidth =3504
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3504
                    Width =648
                    Height =230
                    FontSize =8
                    TabIndex =2
                    BorderColor =14211288
                    ForeColor =0
                    Name ="contactPrimaryPhone"
                    ControlSource ="FAServicer"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =3504
                    LayoutCachedWidth =4152
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6240
                    Width =3012
                    Height =230
                    FontSize =8
                    TabIndex =6
                    BorderColor =14211288
                    ForeColor =0
                    Name ="BillingInstructions"
                    ControlSource ="BillingInstructions"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =6240
                    LayoutCachedWidth =9252
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9252
                    Width =2232
                    Height =230
                    FontSize =8
                    TabIndex =7
                    BorderColor =14211288
                    ForeColor =0
                    Name ="TaxInstructions"
                    ControlSource ="TaxInstructions"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =9252
                    LayoutCachedWidth =11484
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4152
                    Width =648
                    Height =230
                    FontSize =8
                    TabIndex =3
                    BorderColor =14211288
                    ForeColor =0
                    Name ="Text297"
                    ControlSource ="FAInvoicing"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =4152
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4800
                    Width =708
                    Height =230
                    FontSize =8
                    TabIndex =4
                    BorderColor =14211288
                    ForeColor =0
                    Name ="Text305"
                    ControlSource ="FACollecting"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =4800
                    LayoutCachedWidth =5508
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5508
                    Width =732
                    Height =230
                    FontSize =8
                    TabIndex =5
                    BorderColor =14211288
                    ForeColor =0
                    Name ="Text315"
                    ControlSource ="FA3rdParty"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =5508
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =230
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
            End
        End
    End
End
CodeBehindForm
' See "ATF_Report_BillingActive.cls"
