Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3300
    DatasheetFontHeight =11
    ItemSuffix =244
    Left =16155
    Top =5130
    Right =19890
    Bottom =6600
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x0b7a20f0903fe540
    End
    RecordSource ="Select * From ScheduleInfo Where SchID =1496;"
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =1530
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =270
                    Width =3297
                    Height =225
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="dealShortDesc"
                    ControlSource ="=StrConv([SchShortDesc],3)"

                    LayoutCachedTop =270
                    LayoutCachedWidth =3297
                    LayoutCachedHeight =495
                    ThemeFontIndex =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2220
                    Top =1080
                    Width =1080
                    Height =225
                    FontSize =8
                    TabIndex =5
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="docSchedSent"
                    ControlSource ="=IIf([mlorig]=\"Portfolio\",Null,\"Sent \" & Format([SchOut],\"mm/dd/yy\"))"
                    Format ="Short Date"
                    InputMask ="\"Sched Sent\""
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2220
                    LayoutCachedTop =1080
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =1305
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2220
                    Top =1305
                    Width =1080
                    Height =225
                    FontSize =8
                    TabIndex =8
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="docSchedBack"
                    ControlSource ="=IIf([mlorig]=\"Portfolio\",Null,\"Back \" & Format([SchIn],\"mm/dd/yy\"))"
                    Format ="Short Date"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2220
                    LayoutCachedTop =1305
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =1530
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Width =3297
                    FontSize =10
                    FontWeight =700
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =8210719
                    Name ="Schedule"
                    ControlSource ="=\"Schedule \" & [SchNo]"
                    StatusBarText ="schedule number starts with 1"
                    FontName ="Segoe UI Semibold"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Double click to open schedule"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =3297
                    LayoutCachedHeight =240
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =600
                    Top =1305
                    Width =1560
                    Height =225
                    FontSize =8
                    TabIndex =7
                    BackColor =-2147483643
                    BorderColor =8965045
                    Name ="SumOfOrigEquipCost"
                    ControlSource ="=IIf([mlorig]=\"Portfolio\",Null,\"OEC \" & Format([ECost],\"$#,###\"))"
                    Format ="$#,##0.00;($#,##0.00)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =600
                    LayoutCachedTop =1305
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =1530
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =600
                    Top =1080
                    Width =1560
                    Height =225
                    FontSize =8
                    TabIndex =4
                    BackColor =-2147483643
                    Name ="Text65"
                    ControlSource ="=IIf([mlorig]=\"Portfolio\",Null,\"Assignee \" & UCase([Bassign]))"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =600
                    LayoutCachedTop =1080
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =1305
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    BorderThemeColorIndex =1
                    BorderShade =65.0
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =660
                    Width =3300
                    FontSize =8
                    TabIndex =2
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="FAID"
                    ControlSource ="=IIf([mlorig]=\"Portfolio\",\"Tracking Only\",\"FAID \" & [SchFAID])"
                    StatusBarText ="deal id"
                    ShortcutMenuBar ="Sch FAID may be overriden at group or unit level"
                    ControlTipText ="Sch FAID may be overriden at group or unit level"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =660
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =900
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =1095
                    Width =480
                    Height =345
                    FontWeight =700
                    TabIndex =6
                    LeftMargin =45
                    TopMargin =45
                    RightMargin =45
                    BottomMargin =45
                    BackColor =0
                    BorderColor =0
                    Name ="AssignedIndicator\015\012\015\012\015\012\015\012"
                    ControlSource ="=IIf([mlorig]=\"Portfolio\",\"n/a\",[Assigned_Indicator])"
                    EventProcPrefix ="AssignedIndicator________"

                    LayoutCachedLeft =60
                    LayoutCachedTop =1095
                    LayoutCachedWidth =540
                    LayoutCachedHeight =1440
                    ForeThemeColorIndex =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =238
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1800
                    Top =720
                    Width =345
                    Height =165
                    FontSize =8
                    TabIndex =3
                    BackColor =-2147483643
                    BorderColor =0
                    Name ="SchID"
                    ControlSource ="SchID"
                    StatusBarText ="Master Lease No used on documents i.e. CofA"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =720
                    LayoutCachedWidth =2145
                    LayoutCachedHeight =885
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =60
                    Top =930
                    Width =480
                    Height =240
                    FontSize =7
                    Name ="Label66"
                    Caption ="Funded"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =60
                    LayoutCachedTop =930
                    LayoutCachedWidth =540
                    LayoutCachedHeight =1170
                    ThemeFontIndex =1
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
            End
        End
    End
End
CodeBehindForm
' See "ScheduleInfo_SubForm2.cls"
