Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11760
    DatasheetFontHeight =11
    ItemSuffix =49
    Left =4980
    Top =4065
    ShortcutMenuBar ="PrintMe"
    RecSrcDt = Begin
        0xf6ac7ebf15f6e440
    End
    RecordSource ="ATF_ReportActiveQ"
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
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="SchID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="GroupNo"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =315
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Width =9600
                    Height =315
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    Name ="ATFHeader"
                    ControlSource ="=\"MLA: \" & [ATFDesc] & \"    Schedule: \" & [SchNo] & \"   \" & [SchDescriptio"
                        "n] & IIf([Page]>1,\" - Cont'd\",\"\")"

                    LayoutCachedWidth =9600
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9660
                    Width =1830
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="Text36"
                    ControlSource ="=\"FAID: \" & [FAID] & IIf([Page]>1,\" - Cont'd\",\"\")"

                    LayoutCachedLeft =9660
                    LayoutCachedWidth =11490
                    LayoutCachedHeight =315
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2367
            Name ="GroupHeader0"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =540
                    Width =936
                    FontSize =9
                    TabIndex =1
                    BorderColor =6108695
                    ForeColor =0
                    Name ="ATFDate"
                    ControlSource ="ATFDate"
                    Format ="Short Date"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =60
                    LayoutCachedTop =540
                    LayoutCachedWidth =996
                    LayoutCachedHeight =780
                    RowStart =7
                    RowEnd =7
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
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4980
                    Top =540
                    Width =936
                    FontSize =9
                    TabIndex =2
                    BorderColor =6108695
                    ForeColor =0
                    Name ="LPAFDate"
                    ControlSource ="LPAFDate"
                    Format ="Short Date"
                    ControlTipText ="In SFDC, field is Actual Close Date on the Opportunity form"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =4980
                    LayoutCachedTop =540
                    LayoutCachedWidth =5916
                    LayoutCachedHeight =780
                    RowStart =7
                    RowEnd =7
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =996
                    Top =300
                    Width =2304
                    Height =240
                    FontSize =8
                    Name ="Label314"
                    Caption ="Master Lease"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =996
                    LayoutCachedTop =300
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =540
                    RowStart =4
                    RowEnd =4
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =60
                    Top =300
                    Width =936
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label96"
                    Caption ="Import Date"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =60
                    LayoutCachedTop =300
                    LayoutCachedWidth =996
                    LayoutCachedHeight =540
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    Left =4980
                    Top =300
                    Width =936
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label97"
                    Caption ="LPAF Signed"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4980
                    LayoutCachedTop =300
                    LayoutCachedWidth =5916
                    LayoutCachedHeight =540
                    RowStart =6
                    RowEnd =6
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =4425
                    Top =300
                    Width =555
                    Height =240
                    FontSize =8
                    Name ="Label585"
                    Caption ="BDEX"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4425
                    LayoutCachedTop =300
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =540
                    RowStart =8
                    RowEnd =8
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =3900
                    Top =300
                    Width =525
                    Height =240
                    FontSize =8
                    Name ="Label384"
                    Caption ="Units"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3900
                    LayoutCachedTop =300
                    LayoutCachedWidth =4425
                    LayoutCachedHeight =540
                    RowStart =8
                    RowEnd =8
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4425
                    Top =540
                    Width =555
                    FontSize =9
                    BorderColor =14277081
                    ForeColor =0
                    Name ="BDEXname"
                    ControlSource ="BDEX"
                    StatusBarText ="Required - The BDEX (Salesperson) for the Client"
                    ValidationRule ="Is Not Null"
                    ValidationText ="BDEX is Required"
                    DefaultValue ="\"HS\""
                    Tag ="Required"
                    ControlTipText ="Required - The BDEX (Salesperson) for the Client"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4425
                    LayoutCachedTop =540
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =780
                    RowStart =9
                    RowEnd =9
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
                    Left =3900
                    Top =540
                    Width =525
                    FontSize =9
                    TabIndex =4
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Combo383"
                    ControlSource ="NumOfUnits"
                    StatusBarText ="Required - The BDEX (Salesperson) for the Client"
                    ValidationRule ="Is Not Null"
                    ValidationText ="BDEX is Required"
                    DefaultValue ="\"HS\""
                    Tag ="Required"
                    ControlTipText ="Required - The BDEX (Salesperson) for the Client"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3900
                    LayoutCachedTop =540
                    LayoutCachedWidth =4425
                    LayoutCachedHeight =780
                    RowStart =9
                    RowEnd =9
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
                    Left =996
                    Top =540
                    Width =2304
                    FontSize =9
                    TabIndex =3
                    BorderColor =14277081
                    ForeColor =0
                    Name ="MLIDName"
                    ControlSource ="MLNo"
                    StatusBarText ="deal id"
                    ValidationRule ="Is Not Null"
                    ValidationText ="MLID is required"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =996
                    LayoutCachedTop =540
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =780
                    RowStart =5
                    RowEnd =5
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =3300
                    Top =300
                    Width =600
                    Height =240
                    FontSize =8
                    Name ="Label3"
                    Caption ="Groups"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3300
                    LayoutCachedTop =300
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =540
                    RowStart =8
                    RowEnd =8
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3300
                    Top =540
                    Width =600
                    FontSize =9
                    TabIndex =5
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text5"
                    ControlSource ="NumofGroups"
                    StatusBarText ="Required - The BDEX (Salesperson) for the Client"
                    ValidationRule ="Is Not Null"
                    ValidationText ="BDEX is Required"
                    DefaultValue ="\"HS\""
                    Tag ="Required"
                    ControlTipText ="Required - The BDEX (Salesperson) for the Client"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3300
                    LayoutCachedTop =540
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =780
                    RowStart =9
                    RowEnd =9
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
                    Left =5916
                    Top =540
                    Width =1140
                    FontSize =9
                    TabIndex =6
                    BorderColor =6108695
                    ForeColor =0
                    Name ="Text6"
                    ControlSource ="[Assignee_Bank]"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =5916
                    LayoutCachedTop =540
                    LayoutCachedWidth =7056
                    LayoutCachedHeight =780
                    RowStart =7
                    RowEnd =7
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7056
                    Top =540
                    Width =979
                    FontSize =9
                    TabIndex =7
                    BorderColor =6108695
                    ForeColor =0
                    Name ="Text8"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =7056
                    LayoutCachedTop =540
                    LayoutCachedWidth =8035
                    LayoutCachedHeight =780
                    RowStart =7
                    RowEnd =7
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =7056
                    Top =300
                    Width =979
                    Height =240
                    FontSize =8
                    BorderColor =16777215
                    Name ="Label9"
                    Caption ="Interim Fund"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7056
                    LayoutCachedTop =300
                    LayoutCachedWidth =8035
                    LayoutCachedHeight =540
                    RowStart =6
                    RowEnd =6
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5916
                    Top =300
                    Width =1140
                    FontSize =8
                    TabIndex =8
                    BorderColor =16777215
                    Name ="AssignedLabel"
                    ControlSource ="=IIf([Assigned]=\"F\",\"Assignee\",\"Proj. Assignee\")"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5916
                    LayoutCachedTop =300
                    LayoutCachedWidth =7056
                    LayoutCachedHeight =540
                    RowStart =6
                    RowEnd =6
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =9150
                    Top =300
                    Width =1485
                    Height =240
                    FontSize =8
                    Name ="Label37"
                    Caption ="Grand Total"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =9150
                    LayoutCachedTop =300
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =540
                    RowStart =8
                    RowEnd =8
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9150
                    Top =540
                    Width =1485
                    FontSize =9
                    TabIndex =9
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text38"
                    ControlSource ="=DSum(\"TotalGroupOEC\",\"vw_OEC_Rents_LRF_Cost\",\"SchId=\" & [SchID])"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Required - The BDEX (Salesperson) for the Client"
                    ValidationRule ="Is Not Null"
                    ValidationText ="BDEX is Required"
                    DefaultValue ="\"HS\""
                    Tag ="Required"
                    ControlTipText ="Required - The BDEX (Salesperson) for the Client"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =9150
                    LayoutCachedTop =540
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =780
                    RowStart =9
                    RowEnd =9
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8040
                    Top =540
                    Width =1110
                    FontSize =9
                    TabIndex =10
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text40"
                    ControlSource ="SFDCType"
                    StatusBarText ="Required - The BDEX (Salesperson) for the Client"
                    ValidationRule ="Is Not Null"
                    ValidationText ="BDEX is Required"
                    DefaultValue ="\"HS\""
                    Tag ="Required"
                    ControlTipText ="Required - The BDEX (Salesperson) for the Client"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8040
                    LayoutCachedTop =540
                    LayoutCachedWidth =9150
                    LayoutCachedHeight =780
                    RowStart =9
                    RowEnd =9
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =8040
                    Top =300
                    Width =1110
                    Height =240
                    FontSize =8
                    Name ="Label41"
                    Caption ="SFDC Type"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8040
                    LayoutCachedTop =300
                    LayoutCachedWidth =9150
                    LayoutCachedHeight =540
                    RowStart =8
                    RowEnd =8
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10635
                    Top =540
                    Width =885
                    FontSize =9
                    TabIndex =11
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text42"
                    ControlSource ="=IIf(DSum(\"UTASyndGroup\",\"vw_OEC_Rents_LRF_Cost\",\"SchId=\" & [SchID])>0,DSu"
                        "m(\"UTASyndGroup\",\"vw_OEC_Rents_LRF_Cost\",\"SchId=\" & [SchID]),DSum(\"UTAPro"
                        "jGroup\",\"vw_OEC_Rents_LRF_Cost\",\"SchId=\" & [SchID]))"
                    Format ="$#,##0.00;($#,##0.00)"
                    StatusBarText ="Required - The BDEX (Salesperson) for the Client"
                    ValidationRule ="Is Not Null"
                    ValidationText ="BDEX is Required"
                    DefaultValue ="\"HS\""
                    Tag ="Required"
                    ControlTipText ="Required - The BDEX (Salesperson) for the Client"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10635
                    LayoutCachedTop =540
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =780
                    RowStart =9
                    RowEnd =9
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    CurrencySymbol ="$"
                End
                Begin Label
                    TextAlign =2
                    Left =10635
                    Top =300
                    Width =885
                    Height =240
                    FontSize =8
                    Name ="Label43"
                    Caption ="Sch UTA"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10635
                    LayoutCachedTop =300
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =540
                    RowStart =8
                    RowEnd =8
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Label
                    SpecialEffect =5
                    OldBorderStyle =1
                    Width =11523
                    Height =304
                    FontWeight =700
                    BorderColor =7500402
                    ForeColor =0
                    Name ="Label1"
                    Caption ="Lease Origination"
                    LayoutCachedWidth =11523
                    LayoutCachedHeight =304
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =60
                    Top =1140
                    Width =864
                    Height =240
                    FontSize =8
                    Name ="Label304"
                    Caption ="Lessee"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =60
                    LayoutCachedTop =1140
                    LayoutCachedWidth =924
                    LayoutCachedHeight =1380
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =924
                    Top =1140
                    Width =4320
                    FontSize =9
                    TabIndex =12
                    BorderColor =14277081
                    ForeColor =0
                    Name ="LesseeName"
                    ControlSource ="CompanyName"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =924
                    LayoutCachedTop =1140
                    LayoutCachedWidth =5244
                    LayoutCachedHeight =1380
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
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
                    Left =924
                    Top =1410
                    Width =4320
                    FontSize =9
                    TabIndex =13
                    BorderColor =14277081
                    ForeColor =0
                    Name ="CoLesseeName"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)>0,[CoLessee],Null)"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =924
                    LayoutCachedTop =1410
                    LayoutCachedWidth =5244
                    LayoutCachedHeight =1650
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    SpecialEffect =5
                    OldBorderStyle =1
                    TextAlign =1
                    Top =825
                    Width =11448
                    Height =304
                    FontWeight =700
                    BorderColor =7500402
                    ForeColor =0
                    Name ="Label10"
                    Caption ="Lessee Information"
                    LayoutCachedTop =825
                    LayoutCachedWidth =11448
                    LayoutCachedHeight =1129
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =5244
                    Top =1140
                    Width =576
                    Height =240
                    FontSize =8
                    Name ="Label11"
                    Caption ="Tax ID"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =5244
                    LayoutCachedTop =1140
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =1380
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5820
                    Top =1140
                    Width =1920
                    FontSize =9
                    TabIndex =14
                    BorderColor =14277081
                    ForeColor =0
                    Name ="TaxID"
                    ControlSource ="clientTaxID"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5820
                    LayoutCachedTop =1140
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =1380
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5820
                    Top =1410
                    Width =1920
                    FontSize =9
                    TabIndex =15
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text14"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)>0,DLookUp(\"clientTaxID\",\"Clients\",\"ClientId=\" & [C"
                        "oLesseeID]),Null)"
                    StatusBarText ="Also enter Short Name Nm for Reports - it is required."
                    Tag ="Required"
                    ControlTipText ="Also enter Short Name Nm for Reports - it is required."
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5820
                    LayoutCachedTop =1410
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =1650
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Subform
                    Locked = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Top =1919
                    Width =11760
                    Height =224
                    TabIndex =16
                    Name ="ATF_Report_Guarantors"
                    SourceObject ="Report.ATF_Report_GuarantorsActive"
                    LinkChildFields ="FAID"
                    LinkMasterFields ="FAID"

                    LayoutCachedTop =1919
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =2143
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =1410
                    Width =864
                    FontSize =8
                    TabIndex =17
                    BorderColor =14277081
                    Name ="Label305"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)>0,\"Co-Lessee\",Null)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =1410
                    LayoutCachedWidth =924
                    LayoutCachedHeight =1650
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5244
                    Top =1410
                    Width =576
                    FontSize =8
                    TabIndex =18
                    BorderColor =14277081
                    Name ="Label12"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)>0,\"Tax ID\",Null)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =5244
                    LayoutCachedTop =1410
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =1650
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =2143
                    Width =11760
                    Height =224
                    TabIndex =19
                    Name ="Report.ATF_Report_ContactsActive"
                    SourceObject ="Report.ATF_Report_ContactsActive"
                    LinkChildFields ="SCHID"
                    LinkMasterFields ="SCHID"
                    EventProcPrefix ="Report_ATF_Report_ContactsActive"

                    LayoutCachedTop =2143
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =2367
                End
                Begin Subform
                    OldBorderStyle =0
                    Top =1695
                    Width =11760
                    Height =224
                    TabIndex =20
                    Name ="ATF_Report_CoLesseeActive"
                    SourceObject ="Report.ATF_Report_CoLesseeActive"
                    LinkChildFields ="ScheduleID"
                    LinkMasterFields ="SchID"

                    LayoutCachedTop =1695
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =1919
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7245
                    Width =4266
                    FontSize =8
                    TabIndex =21
                    BorderColor =6108695
                    ForeColor =0
                    Name ="Text46"
                    ControlSource ="=Now()"
                    Format ="General Date"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ShowDatePicker =1

                    LayoutCachedLeft =7245
                    LayoutCachedWidth =11511
                    LayoutCachedHeight =240
                    RowStart =7
                    RowEnd =7
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
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1364
            BreakLevel =1
            Name ="GroupHeader1"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    Locked = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Width =11745
                    Height =1364
                    Name ="ATF_Report_StructureActive"
                    SourceObject ="Report.ATF_Report_StructureActive"
                    LinkChildFields ="SGrpID"
                    LinkMasterFields ="GroupID"

                    LayoutCachedWidth =11745
                    LayoutCachedHeight =1364
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =344
            Name ="Detail"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Top =45
                    Width =11745
                    Height =299
                    Name ="ATF_Report_BillingActive"
                    SourceObject ="Report.ATF_Report_BillingActive"
                    LinkChildFields ="SCHID"
                    LinkMasterFields ="SCHID"

                    LayoutCachedTop =45
                    LayoutCachedWidth =11745
                    LayoutCachedHeight =344
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =614
            Name ="GroupFooter2"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =315
                    Width =11745
                    Height =299
                    Name ="ATF_Report_SchNotesActive"
                    SourceObject ="Report.ATF_Report_SchNotesActive"
                    LinkChildFields ="SchID"
                    LinkMasterFields ="SchID"

                    LayoutCachedTop =315
                    LayoutCachedWidth =11745
                    LayoutCachedHeight =614
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Top =15
                    Width =11745
                    Height =299
                    TabIndex =1
                    Name ="ATf_Report_ScheduleSumsActive"
                    SourceObject ="Report.ATf_Report_ScheduleSumsActive"
                    LinkChildFields ="SCHID"
                    LinkMasterFields ="SCHID"

                    LayoutCachedTop =15
                    LayoutCachedWidth =11745
                    LayoutCachedHeight =314
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =240
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6435
                    Width =1260
                    FontSize =9
                    BorderColor =14277081
                    ForeColor =0
                    Name ="ATFID"
                    ControlSource ="ATFID"
                    StatusBarText ="deal id"
                    ValidationRule ="Is Not Null"
                    ValidationText ="FAID is required"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6435
                    LayoutCachedWidth =7695
                    LayoutCachedHeight =240
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
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
                    Left =4935
                    Width =1260
                    FontSize =9
                    TabIndex =1
                    BorderColor =14277081
                    ForeColor =0
                    Name ="Text33"
                    ControlSource ="schID"
                    StatusBarText ="deal id"
                    ValidationRule ="Is Not Null"
                    ValidationText ="FAID is required"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4935
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =240
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    BackThemeColorIndex =-1
                    BorderShade =85.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7890
                    Width =3600
                    FontSize =9
                    TabIndex =2
                    BorderColor =0
                    ForeColor =0
                    Name ="Text66"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =0

                    LayoutCachedLeft =7890
                    LayoutCachedWidth =11490
                    LayoutCachedHeight =240
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
            End
        End
    End
End
