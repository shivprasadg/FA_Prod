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
    Width =11703
    DatasheetFontHeight =11
    ItemSuffix =318
    Top =600
    RecSrcDt = Begin
        0x5ccdc4f014f6e440
    End
    RecordSource ="ATF_Report_SchNotesActiveQ"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
            SortOrder = NotDefault
            GroupHeader = NotDefault
            ControlSource ="SchID"
        End
        Begin BreakLevel
            ControlSource ="RevID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="RevID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            RepeatSection = NotDefault
            Height =680
            Name ="GroupHeader1"
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    SpecialEffect =5
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =60
                    Width =11405
                    Height =330
                    ColumnOrder =0
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label15"
                    ControlSource ="=\"Schedule \" & DLookUp(\"SchNo\",\"Schedule\",\"SchId=\" & [schid]) & \" Revis"
                        "ion - Notes\" & IIf([Page]>1,\" Cont'd\",\"\")"
                    GridlineColor =10921638

                    LayoutCachedTop =60
                    LayoutCachedWidth =11405
                    LayoutCachedHeight =390
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =525
                    Top =440
                    Width =1008
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label247"
                    Caption ="Rev Details"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =525
                    LayoutCachedTop =440
                    LayoutCachedWidth =1533
                    LayoutCachedHeight =680
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =45
                    Top =440
                    Width =480
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label27"
                    Caption ="Rev"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =45
                    LayoutCachedTop =440
                    LayoutCachedWidth =525
                    LayoutCachedHeight =680
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =1533
                    Top =440
                    Width =7680
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label289"
                    Caption ="Revision Comments"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =1533
                    LayoutCachedTop =440
                    LayoutCachedWidth =9213
                    LayoutCachedHeight =680
                    ColumnStart =2
                    ColumnEnd =2
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10065
                    Top =60
                    Width =600
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =1
                    BorderColor =14277081
                    Name ="SchID"
                    ControlSource ="SchID"
                    StatusBarText ="deal id"
                    ValidationRule ="Is Not Null"
                    ValidationText ="FAID is required"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =10065
                    LayoutCachedTop =60
                    LayoutCachedWidth =10665
                    LayoutCachedHeight =300
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Line
                    BorderWidth =2
                    Width =11405
                    Name ="Line260"
                    GridlineColor =10921638
                    LayoutCachedWidth =11405
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10469
                    Top =440
                    Width =1008
                    FontSize =8
                    TabIndex =2
                    BorderColor =16777215
                    Name ="ReviewLabel"
                    ControlSource ="=IIf(Nz([ReviewedBy],\"\")=\"\",Null,\"Reviewed by\")"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638

                    LayoutCachedLeft =10469
                    LayoutCachedTop =440
                    LayoutCachedWidth =11477
                    LayoutCachedHeight =680
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
            End
        End
        Begin BreakHeader
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =485
            BreakLevel =2
            Name ="GroupHeader0"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =540
                    Width =1008
                    Height =260
                    FontSize =9
                    TabIndex =1
                    BorderColor =14211288
                    Name ="VendorSNm"
                    ControlSource ="ActionDate"
                    Format ="Short Date"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =540
                    LayoutCachedWidth =1548
                    LayoutCachedHeight =260
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
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
                    Left =60
                    Width =480
                    Height =308
                    FontSize =10
                    FontWeight =700
                    BorderColor =14211288
                    Name ="contactPrimaryPhone"
                    ControlSource ="RevID"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =60
                    LayoutCachedWidth =540
                    LayoutCachedHeight =330
                    RowStart =1
                    RowEnd =1
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
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =540
                    Top =210
                    Width =1008
                    Height =260
                    FontSize =9
                    TabIndex =3
                    BorderColor =14211288
                    ForeColor =5855577
                    Name ="Text270"
                    ControlSource ="SysUser"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =540
                    LayoutCachedTop =210
                    LayoutCachedWidth =1548
                    LayoutCachedHeight =470
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1548
                    Width =8895
                    Height =485
                    FontSize =8
                    TabIndex =2
                    BorderColor =14211288
                    Name ="ActionDesc"
                    ControlSource ="ActionDesc"
                    FontName ="Segoe UI"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    TextFormat =1
                    ShowDatePicker =1

                    LayoutCachedLeft =1548
                    LayoutCachedWidth =10443
                    LayoutCachedHeight =485
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
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
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10469
                    Top =210
                    Width =1008
                    Height =260
                    FontSize =9
                    TabIndex =4
                    BorderColor =14211288
                    ForeColor =5855577
                    Name ="ReviewedByTxt"
                    ControlSource ="=IIf(Nz([ReviewedBy],\"\")=\"\",Null,[ReviewedBy])"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =10469
                    LayoutCachedTop =210
                    LayoutCachedWidth =11477
                    LayoutCachedHeight =470
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10469
                    Width =1008
                    Height =260
                    FontSize =9
                    TabIndex =5
                    BorderColor =14211288
                    Name ="ReviewedDateTxt"
                    ControlSource ="=IIf(Nz([ReviewedDate],\"\")=\"\",Null,[ReviewedDate])"
                    Format ="Short Date"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =10469
                    LayoutCachedWidth =11477
                    LayoutCachedHeight =260
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
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
            Height =0
            Name ="Detail"
            BackThemeColorIndex =1
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
' See "ATF_Report_SchNotesActive.cls"
