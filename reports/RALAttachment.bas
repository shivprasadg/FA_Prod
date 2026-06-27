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
    Width =14400
    DatasheetFontHeight =11
    ItemSuffix =129
    Left =2445
    Top =510
    RecSrcDt = Begin
        0x10eded5a6df4e540
    End
    RecordSource ="RALUnits"
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
            KeepTogether =2
            ControlSource ="SchID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="GroupNo"
        End
        Begin BreakLevel
            ControlSource ="GroupNo"
        End
        Begin BreakLevel
            ControlSource ="AssetID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =360
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Width =10980
                    Height =315
                    ColumnWidth =1110
                    ColumnOrder =0
                    FontWeight =700
                    ForeColor =0
                    Name ="ScheduleHeader"
                    ControlSource ="=\"Attachment 1 <br />To Rent Assignment Letter dated: \" & Format([Letterdate],"
                        "\"mmmm dd\"\", \"\"yyyy\") & \"<br />Re: Lease No. \" & [MLNO] & \"; Schedule No"
                        ". \" & [Schedule] & \" (\" & [GroupsIn] & \")\""
                    FontName ="Times New Roman"
                    TextFormat =1

                    LayoutCachedLeft =60
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =315
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="GroupHeader0"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =795
            BreakLevel =1
            Name ="GroupHeader1"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =195
                    Width =3150
                    Height =315
                    FontWeight =700
                    Name ="GroupNo"
                    ControlSource ="GroupNo"
                    FontName ="Times New Roman"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =195
                    LayoutCachedWidth =3345
                    LayoutCachedHeight =315
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =5
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =195
                    Top =315
                    Width =1185
                    Height =480
                    FontSize =10
                    FontWeight =700
                    BackColor =8210719
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Label17"
                    Caption ="Unit No."
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =195
                    LayoutCachedTop =315
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =795
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1380
                    Top =315
                    Width =1965
                    Height =480
                    FontSize =10
                    FontWeight =700
                    BackColor =8210719
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Label26"
                    Caption ="VIN"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1380
                    LayoutCachedTop =315
                    LayoutCachedWidth =3345
                    LayoutCachedHeight =795
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3345
                    Top =315
                    Width =3045
                    Height =480
                    FontSize =10
                    FontWeight =700
                    BackColor =8210719
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Label33"
                    Caption ="Year / Make / Model / Type"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3345
                    LayoutCachedTop =315
                    LayoutCachedWidth =6390
                    LayoutCachedHeight =795
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8850
                    Top =315
                    Width =1950
                    Height =480
                    FontSize =10
                    FontWeight =700
                    BackColor =8210719
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Label40"
                    Caption ="Location Address"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8850
                    LayoutCachedTop =315
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =795
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =10800
                    Top =315
                    Width =1995
                    Height =480
                    FontSize =10
                    FontWeight =700
                    BackColor =8210719
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Label47"
                    Caption ="City"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10800
                    LayoutCachedTop =315
                    LayoutCachedWidth =12795
                    LayoutCachedHeight =795
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =12795
                    Top =315
                    Width =525
                    Height =480
                    FontSize =10
                    FontWeight =700
                    BackColor =8210719
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Label54"
                    Caption ="State"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12795
                    LayoutCachedTop =315
                    LayoutCachedWidth =13320
                    LayoutCachedHeight =795
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =13320
                    Top =315
                    Width =1035
                    Height =480
                    FontSize =10
                    FontWeight =700
                    BackColor =8210719
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Label61"
                    Caption ="Zip"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13320
                    LayoutCachedTop =315
                    LayoutCachedWidth =14355
                    LayoutCachedHeight =795
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =7770
                    Top =315
                    Width =1080
                    Height =480
                    FontSize =10
                    FontWeight =700
                    BackColor =8210719
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Label96"
                    Caption ="Basic Rent \015\012(per unit)"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7770
                    LayoutCachedTop =315
                    LayoutCachedWidth =8850
                    LayoutCachedHeight =795
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =6390
                    Top =315
                    Width =1380
                    Height =480
                    FontSize =10
                    FontWeight =700
                    BackColor =8210719
                    BorderColor =9605778
                    ForeColor =16777215
                    Name ="Label98"
                    Caption ="Lessor's Cost \015\012(per unit)"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6390
                    LayoutCachedTop =315
                    LayoutCachedWidth =7770
                    LayoutCachedHeight =795
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin EmptyCell
                    Left =3345
                    Width =3045
                    Height =315
                    Name ="EmptyCell122"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3345
                    LayoutCachedWidth =6390
                    LayoutCachedHeight =315
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =5
                End
                Begin EmptyCell
                    Left =6390
                    Width =1380
                    Height =315
                    Name ="EmptyCell123"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6390
                    LayoutCachedWidth =7770
                    LayoutCachedHeight =315
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =5
                End
                Begin EmptyCell
                    Left =7770
                    Width =1080
                    Height =315
                    Name ="EmptyCell124"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7770
                    LayoutCachedWidth =8850
                    LayoutCachedHeight =315
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =5
                End
                Begin EmptyCell
                    Left =8850
                    Width =1950
                    Height =315
                    Name ="EmptyCell125"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8850
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =315
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =5
                End
                Begin EmptyCell
                    Left =10800
                    Width =1995
                    Height =315
                    Name ="EmptyCell126"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =10800
                    LayoutCachedWidth =12795
                    LayoutCachedHeight =315
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =5
                End
                Begin EmptyCell
                    Left =12795
                    Width =525
                    Height =315
                    Name ="EmptyCell127"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =12795
                    LayoutCachedWidth =13320
                    LayoutCachedHeight =315
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =5
                End
                Begin EmptyCell
                    Left =13320
                    Width =1035
                    Height =315
                    Name ="EmptyCell128"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =13320
                    LayoutCachedWidth =14355
                    LayoutCachedHeight =315
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =5
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =315
            Name ="Detail"
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =195
                    Width =1185
                    Height =315
                    FontSize =10
                    BorderColor =0
                    ForeColor =0
                    Name ="Unitnum"
                    ControlSource ="Unitnum"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =195
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =315
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1380
                    Width =1965
                    Height =315
                    ColumnWidth =3225
                    FontSize =10
                    TabIndex =1
                    BorderColor =0
                    ForeColor =0
                    Name ="VIN"
                    ControlSource ="VIN"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1380
                    LayoutCachedWidth =3345
                    LayoutCachedHeight =315
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =3345
                    Width =3045
                    Height =315
                    ColumnWidth =4995
                    FontSize =10
                    TabIndex =2
                    BorderColor =0
                    ForeColor =0
                    Name ="UnitDesc"
                    ControlSource ="UnitDescDocVer"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3345
                    LayoutCachedWidth =6390
                    LayoutCachedHeight =315
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =8850
                    Width =1950
                    Height =315
                    FontSize =10
                    TabIndex =5
                    BorderColor =0
                    ForeColor =0
                    Name ="LAddress"
                    ControlSource ="LAddress"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8850
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =315
                    RowStart =2
                    RowEnd =2
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =10800
                    Width =1995
                    Height =315
                    FontSize =10
                    TabIndex =6
                    BorderColor =0
                    ForeColor =0
                    Name ="LCity"
                    ControlSource ="LCity"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10800
                    LayoutCachedWidth =12795
                    LayoutCachedHeight =315
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12795
                    Width =525
                    Height =315
                    FontSize =10
                    TabIndex =7
                    BorderColor =0
                    ForeColor =0
                    Name ="LState"
                    ControlSource ="LState"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12795
                    LayoutCachedWidth =13320
                    LayoutCachedHeight =315
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13320
                    Width =1035
                    Height =315
                    FontSize =10
                    TabIndex =8
                    BorderColor =0
                    ForeColor =0
                    Name ="LZip2"
                    ControlSource ="=Left([LZip],5)"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13320
                    LayoutCachedWidth =14355
                    LayoutCachedHeight =315
                    RowStart =2
                    RowEnd =2
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6390
                    Width =1380
                    Height =315
                    FontSize =10
                    TabIndex =3
                    BorderColor =0
                    ForeColor =0
                    Name ="Text84"
                    ControlSource ="LessorCost"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6390
                    LayoutCachedWidth =7770
                    LayoutCachedHeight =315
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7770
                    Width =1080
                    Height =315
                    FontSize =10
                    TabIndex =4
                    BorderColor =0
                    ForeColor =0
                    Name ="Text85"
                    ControlSource ="UnitRent"
                    GroupTable =5
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7770
                    LayoutCachedWidth =8850
                    LayoutCachedHeight =315
                    RowStart =2
                    RowEnd =2
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =5
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =345
            BreakLevel =1
            Name ="GroupFooter0"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =195
                    Width =1185
                    Height =315
                    FontSize =10
                    BackColor =15523798
                    BorderColor =9605778
                    ForeColor =0
                    Name ="Text108"
                    ControlSource ="=\"Units: \" & Count([AssetID])"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =195
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =315
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1380
                    Width =1965
                    Height =315
                    FontSize =10
                    TabIndex =1
                    BackColor =15523798
                    BorderColor =9605778
                    ForeColor =0
                    Name ="Text109"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1380
                    LayoutCachedWidth =3345
                    LayoutCachedHeight =315
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3345
                    Width =3045
                    Height =315
                    FontSize =10
                    TabIndex =2
                    BackColor =15523798
                    BorderColor =9605778
                    ForeColor =0
                    Name ="Text110"
                    ControlSource ="=\"Total \" & [GroupNo] & \": \""
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3345
                    LayoutCachedWidth =6390
                    LayoutCachedHeight =315
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =8850
                    Width =1950
                    Height =315
                    FontSize =10
                    TabIndex =3
                    BackColor =15523798
                    BorderColor =9605778
                    ForeColor =0
                    Name ="Text111"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8850
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =315
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =10800
                    Width =1995
                    Height =315
                    FontSize =10
                    TabIndex =4
                    BackColor =15523798
                    BorderColor =9605778
                    ForeColor =0
                    Name ="Text112"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =10800
                    LayoutCachedWidth =12795
                    LayoutCachedHeight =315
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12795
                    Width =510
                    Height =315
                    FontSize =10
                    TabIndex =5
                    BackColor =15523798
                    BorderColor =9605778
                    ForeColor =0
                    Name ="Text113"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =12795
                    LayoutCachedWidth =13305
                    LayoutCachedHeight =315
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13305
                    Width =1050
                    Height =315
                    FontSize =10
                    TabIndex =6
                    BackColor =15523798
                    BorderColor =9605778
                    ForeColor =0
                    Name ="Text114"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =13305
                    LayoutCachedWidth =14355
                    LayoutCachedHeight =315
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6390
                    Width =1380
                    Height =315
                    FontSize =10
                    TabIndex =7
                    BackColor =15523798
                    BorderColor =9605778
                    ForeColor =0
                    Name ="Text115"
                    ControlSource ="=Sum([LessorCost])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6390
                    LayoutCachedWidth =7770
                    LayoutCachedHeight =315
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7770
                    Width =1080
                    Height =315
                    FontSize =10
                    TabIndex =8
                    BackColor =15523798
                    BorderColor =9605778
                    ForeColor =0
                    Name ="Text116"
                    ControlSource ="=Sum([UnitRent])"
                    Format ="$#,##0.00;($#,##0.00)"
                    GroupTable =14
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7770
                    LayoutCachedWidth =8850
                    LayoutCachedHeight =315
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =14
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
