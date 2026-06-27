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
    Width =10087
    DatasheetFontHeight =11
    ItemSuffix =85
    Top =600
    RecSrcDt = Begin
        0x69012943b579e640
    End
    RecordSource ="RALGroupDetails"
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="UnitGroup"
        End
        Begin BreakLevel
            ControlSource ="UnitGroup"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =3180
            OnPrint ="[Event Procedure]"
            Name ="GroupHeader0"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3360
                    Top =2940
                    Width =6719
                    TabIndex =8
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =16711422
                    ForeColor =0
                    Name ="Text55"
                    FontName ="Times New Roman"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3360
                    LayoutCachedTop =2940
                    LayoutCachedWidth =10079
                    LayoutCachedHeight =3180
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3360
                    Top =360
                    Width =6719
                    Height =315
                    ColumnWidth =2625
                    TabIndex =1
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =0
                    ForeColor =0
                    Name ="Desc"
                    ControlSource ="Desc"
                    FontName ="Times New Roman"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3360
                    LayoutCachedTop =360
                    LayoutCachedWidth =10079
                    LayoutCachedHeight =675
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =3
                            TextFontFamily =18
                            Left =150
                            Top =360
                            Width =3210
                            Height =315
                            RightMargin =29
                            BorderColor =16777215
                            ForeColor =0
                            Name ="DescLabel"
                            Caption ="Equipment:"
                            FontName ="Times New Roman"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =150
                            LayoutCachedTop =360
                            LayoutCachedWidth =3360
                            LayoutCachedHeight =675
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3360
                    Top =675
                    Width =6719
                    Height =315
                    TabIndex =2
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =0
                    ForeColor =0
                    Name ="hereto"
                    ControlSource ="=\"See Attachment 1 attached hereto\""
                    FontName ="Times New Roman"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3360
                    LayoutCachedTop =675
                    LayoutCachedWidth =10079
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =3
                            TextFontFamily =18
                            Left =150
                            Top =675
                            Width =3210
                            Height =315
                            RightMargin =29
                            BorderColor =16777215
                            ForeColor =0
                            Name ="HereToLabel"
                            Caption ="VINS:"
                            FontName ="Times New Roman"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =150
                            LayoutCachedTop =675
                            LayoutCachedWidth =3360
                            LayoutCachedHeight =990
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3360
                    Top =990
                    Width =6719
                    Height =435
                    TabIndex =3
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =0
                    ForeColor =0
                    Name ="Lessorcostx"
                    ControlSource ="=\"<u>\" & Format([Lessorcost],\"$#,###.##\") & \"</u><br> (the Lessor’s Cost am"
                        "ount per unit is as set forth on Attachment 1, attached hereto)\012 \""
                    FontName ="Times New Roman"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =3360
                    LayoutCachedTop =990
                    LayoutCachedWidth =10079
                    LayoutCachedHeight =1425
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =3
                            TextFontFamily =18
                            Left =150
                            Top =990
                            Width =3210
                            Height =435
                            RightMargin =29
                            BorderColor =16777215
                            ForeColor =0
                            Name ="LessorCostLabel"
                            Caption ="Total Lessor's Cost:"
                            FontName ="Times New Roman"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =150
                            LayoutCachedTop =990
                            LayoutCachedWidth =3360
                            LayoutCachedHeight =1425
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3360
                    Top =1425
                    Width =6719
                    Height =576
                    ColumnWidth =3135
                    TabIndex =4
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =0
                    ForeColor =0
                    Name ="SyndRent"
                    ControlSource ="SyndRent"
                    FontName ="Times New Roman"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =3360
                    LayoutCachedTop =1425
                    LayoutCachedWidth =10079
                    LayoutCachedHeight =2001
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =3
                            TextFontFamily =18
                            Left =150
                            Top =1425
                            Width =3210
                            Height =576
                            RightMargin =29
                            BorderColor =16777215
                            ForeColor =0
                            Name ="SyndRentLabel"
                            Caption ="Basic Rent \015\012(after all adjustments):"
                            FontName ="Times New Roman"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =150
                            LayoutCachedTop =1425
                            LayoutCachedWidth =3360
                            LayoutCachedHeight =2001
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3360
                    Top =1995
                    Width =6719
                    Height =315
                    TabIndex =5
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =0
                    ForeColor =0
                    Name ="BLCD"
                    ControlSource ="BLCD"
                    FontName ="Times New Roman"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3360
                    LayoutCachedTop =1995
                    LayoutCachedWidth =10079
                    LayoutCachedHeight =2310
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =3
                            TextFontFamily =18
                            Left =150
                            Top =1995
                            Width =3210
                            Height =315
                            RightMargin =29
                            BorderColor =16777215
                            ForeColor =0
                            Name ="BLCDLabel"
                            Caption ="Base Lease Commencement Date:"
                            FontName ="Times New Roman"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =150
                            LayoutCachedTop =1995
                            LayoutCachedWidth =3360
                            LayoutCachedHeight =2310
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3360
                    Top =2625
                    Width =6719
                    Height =315
                    TabIndex =7
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =0
                    ForeColor =0
                    Name ="AssignedRents"
                    ControlSource ="=\"All Rents due during the Primary Term at Basic Rent beginning on the Base Lea"
                        "se Commencement Date; (unless ExchangeIT Option is elected), and any rents due t"
                        "hereafter, as applicable (as set forth in the Schedule under the Section EXchang"
                        "eIT®, Termination and ExtendIT Options.)<br>\""
                    FontName ="Times New Roman"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    TextFormat =1

                    LayoutCachedLeft =3360
                    LayoutCachedTop =2625
                    LayoutCachedWidth =10079
                    LayoutCachedHeight =2940
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =3
                            TextFontFamily =18
                            Left =150
                            Top =2625
                            Width =3210
                            Height =315
                            RightMargin =29
                            BorderColor =16777215
                            ForeColor =0
                            Name ="AssignedLabel"
                            Caption ="Assigned Rents:"
                            FontName ="Times New Roman"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =150
                            LayoutCachedTop =2625
                            LayoutCachedWidth =3360
                            LayoutCachedHeight =2940
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =9014
                    Top =75
                    Width =432
                    Height =150
                    TabIndex =9
                    Name ="SchID"
                    ControlSource ="SchID"

                    LayoutCachedLeft =9014
                    LayoutCachedTop =75
                    LayoutCachedWidth =9446
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8640
                    Top =144
                    Width =432
                    Height =150
                    TabIndex =10
                    Name ="SGrpID"
                    ControlSource ="SGrpID"

                    LayoutCachedLeft =8640
                    LayoutCachedTop =144
                    LayoutCachedWidth =9072
                    LayoutCachedHeight =294
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =150
                    Top =30
                    Width =9929
                    Height =330
                    FontWeight =700
                    LeftMargin =29
                    BorderColor =0
                    ForeColor =0
                    Name ="GroupHeaderLine"
                    ControlSource ="=\"Schedule No. \" & [Schedule] & \" - Group \" & [UnitGroup]"
                    FontName ="Times New Roman"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =150
                    LayoutCachedTop =30
                    LayoutCachedWidth =10079
                    LayoutCachedHeight =360
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =2
                    BackTint =20.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =150
                    Top =2940
                    Width =3210
                    Name ="EmptyCell71"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =150
                    LayoutCachedTop =2940
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =3180
                    RowStart =8
                    RowEnd =8
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3360
                    Top =2310
                    Width =6719
                    Height =315
                    TabIndex =6
                    LeftMargin =29
                    RightMargin =29
                    BorderColor =0
                    ForeColor =0
                    Name ="LeaseTerm"
                    ControlSource ="=([LeaseTermPrimary]-3) & \" months\""
                    FontName ="Times New Roman"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3360
                    LayoutCachedTop =2310
                    LayoutCachedWidth =10079
                    LayoutCachedHeight =2625
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =3
                            TextFontFamily =18
                            Left =150
                            Top =2310
                            Width =3210
                            Height =315
                            RightMargin =29
                            BorderColor =16777215
                            ForeColor =0
                            Name ="LeaseTermLabel"
                            Caption ="Primary Term:"
                            FontName ="Times New Roman"
                            GroupTable =2
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =150
                            LayoutCachedTop =2310
                            LayoutCachedWidth =3360
                            LayoutCachedHeight =2625
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =1
                            ThemeFontIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GroupTable =2
                        End
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =0
            Name ="Detail"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "RAL_Group.cls"
