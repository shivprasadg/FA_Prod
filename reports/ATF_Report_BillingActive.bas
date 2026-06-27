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
            Height =540
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
                    Name ="Label15"
                    Caption ="Billing "
                    GridlineColor =10921638
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
                    Top =330
                    Width =2445
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label247"
                    Caption ="BillTo"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =420
                    LayoutCachedTop =330
                    LayoutCachedWidth =2865
                    LayoutCachedHeight =540
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    Left =2865
                    Top =330
                    Width =630
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label248"
                    Caption ="Method"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2865
                    LayoutCachedTop =330
                    LayoutCachedWidth =3495
                    LayoutCachedHeight =540
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
                    Left =3495
                    Top =330
                    Width =645
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label27"
                    Caption ="Servicer"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =3495
                    LayoutCachedTop =330
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =540
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
                    Left =9240
                    Top =330
                    Width =2235
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label269"
                    Caption ="Tax Instruction"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =9240
                    LayoutCachedTop =330
                    LayoutCachedWidth =11475
                    LayoutCachedHeight =540
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
                    Left =6225
                    Top =330
                    Width =3015
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label289"
                    Caption ="Billing Instructions"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =6225
                    LayoutCachedTop =330
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =540
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
                    Left =4140
                    Top =330
                    Width =645
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label296"
                    Caption ="Invoicer"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =4140
                    LayoutCachedTop =330
                    LayoutCachedWidth =4785
                    LayoutCachedHeight =540
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
                    Left =4785
                    Top =330
                    Width =705
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label304"
                    Caption ="Collector"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =4785
                    LayoutCachedTop =330
                    LayoutCachedWidth =5490
                    LayoutCachedHeight =540
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
                    Left =5490
                    Top =330
                    Width =735
                    Height =210
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label314"
                    Caption ="3rd Party"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =5490
                    LayoutCachedTop =330
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =540
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
                    GridlineColor =10921638
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
                    Name ="Label330"
                    Caption =" Taxes"
                    GridlineColor =10921638
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
                    Name ="BillingAddress"
                    ControlSource ="BillingAddress"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
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
                    Left =2865
                    Width =630
                    Height =230
                    FontSize =8
                    TabIndex =1
                    BorderColor =14211288
                    Name ="Method"
                    ControlSource ="PayMethod"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =2865
                    LayoutCachedWidth =3495
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
                    Left =3495
                    Width =645
                    Height =230
                    FontSize =8
                    TabIndex =2
                    BorderColor =14211288
                    Name ="contactPrimaryPhone"
                    ControlSource ="FAServicer"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =3495
                    LayoutCachedWidth =4140
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
                    Left =6225
                    Width =3015
                    Height =230
                    FontSize =8
                    TabIndex =6
                    BorderColor =14211288
                    Name ="BillingInstructions"
                    ControlSource ="BillingInstructions"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =6225
                    LayoutCachedWidth =9240
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
                    Left =9240
                    Width =2235
                    Height =230
                    FontSize =8
                    TabIndex =7
                    BorderColor =14211288
                    Name ="TaxInstructions"
                    ControlSource ="TaxInstructions"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =9240
                    LayoutCachedWidth =11475
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
                    Left =4140
                    Width =645
                    Height =230
                    FontSize =8
                    TabIndex =3
                    BorderColor =14211288
                    Name ="Text297"
                    ControlSource ="FAInvoicing"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =4140
                    LayoutCachedWidth =4785
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
                    Left =4785
                    Width =705
                    Height =230
                    FontSize =8
                    TabIndex =4
                    BorderColor =14211288
                    Name ="Text305"
                    ControlSource ="FACollecting"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =4785
                    LayoutCachedWidth =5490
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
                    Left =5490
                    Width =735
                    Height =230
                    FontSize =8
                    TabIndex =5
                    BorderColor =14211288
                    Name ="Text315"
                    ControlSource ="FA3rdParty"
                    FontName ="Tahoma"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =5490
                    LayoutCachedWidth =6225
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
