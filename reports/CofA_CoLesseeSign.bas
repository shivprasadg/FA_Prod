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
    Width =11460
    DatasheetFontHeight =11
    ItemSuffix =67
    Left =6330
    Top =2775
    RecSrcDt = Begin
        0x990addfd3602e540
    End
    RecordSource ="SELECT Clients.clientCompanyName, ScheduleCoLessee.ScheduleID, ScheduleCoLessee."
        "CoLesseeID, ScheduleCoLessee.CoLesseeLabel, ScheduleCoLessee.CoLesseeTaxID, Sche"
        "duleCoLessee.HideOnCOA FROM ScheduleCoLessee INNER JOIN Clients ON ScheduleCoLes"
        "see.CoLesseeID = Clients.clientID WHERE (((ScheduleCoLessee.HideOnCOA)=False));"
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
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =0
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =2280
            Name ="Detail"
            AlternateBackColor =16777215
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    BorderWidth =1
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1305
                    Top =345
                    Width =4935
                    Height =345
                    TabIndex =3
                    BorderColor =0
                    ForeColor =0
                    Name ="Text498"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)=0,Null,\" \")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1305
                    LayoutCachedTop =345
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =690
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
                    GroupTable =2
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1305
                    Width =8295
                    Height =345
                    TabIndex =1
                    BorderColor =0
                    ForeColor =0
                    Name ="Text500"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)=0,Null,[clientCompanyName])"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1305
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =345
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
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
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Width =1305
                    Height =345
                    BorderColor =0
                    ForeColor =0
                    Name ="Text501"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)=0,Null,[CoLesseeLabel] & \":\")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =1305
                    LayoutCachedHeight =345
                    LayoutGroup =1
                    BackThemeColorIndex =-1
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
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Top =345
                    Width =1305
                    Height =345
                    TabIndex =2
                    BorderColor =0
                    ForeColor =0
                    Name ="Label486"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)=0,Null,\"By: \")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =345
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =690
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
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
                    OldBorderStyle =0
                    BorderWidth =1
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6240
                    Top =345
                    Width =630
                    Height =345
                    TabIndex =4
                    BorderColor =0
                    ForeColor =0
                    Name ="Text687"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)=0,Null,\"Date: \")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6240
                    LayoutCachedTop =345
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =690
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
                    GroupTable =2
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    BorderWidth =1
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6870
                    Top =345
                    Width =2730
                    Height =345
                    TabIndex =5
                    BorderColor =0
                    ForeColor =0
                    Name ="Text692"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6870
                    LayoutCachedTop =345
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =690
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
                    GroupTable =2
                End
                Begin EmptyCell
                    Top =690
                    Width =1305
                    Height =300
                    Name ="EmptyCell39"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =690
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6240
                    Top =690
                    Width =630
                    Height =300
                    Name ="EmptyCell41"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6240
                    LayoutCachedTop =690
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6870
                    Top =690
                    Width =2730
                    Height =300
                    Name ="EmptyCell42"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6870
                    LayoutCachedTop =690
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1305
                    Top =690
                    Width =4935
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =0
                    ForeColor =5855577
                    Name ="Label499"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)=0,Null,\"Signature\")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1305
                    LayoutCachedTop =690
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =990
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
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
                    SpecialEffect =5
                    OldBorderStyle =0
                    BorderWidth =1
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1305
                    Top =990
                    Width =8295
                    Height =345
                    TabIndex =8
                    BorderColor =0
                    ForeColor =0
                    Name ="Text487"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)=0,Null,\" \")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1305
                    LayoutCachedTop =990
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
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
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Top =990
                    Width =1305
                    Height =345
                    TabIndex =7
                    BorderColor =0
                    ForeColor =0
                    Name ="Label488"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)=0,Null,\"Name:\")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =990
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =1335
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Top =1335
                    Width =1305
                    Height =300
                    Name ="EmptyCell50"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =1335
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =1635
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6240
                    Top =1335
                    Width =630
                    Height =300
                    Name ="EmptyCell52"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6240
                    LayoutCachedTop =1335
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =1635
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    BorderWidth =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1305
                    Top =1335
                    Width =4935
                    Height =300
                    FontSize =8
                    TabIndex =9
                    BorderColor =0
                    ForeColor =5855577
                    Name ="Label493"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)=0,Null,\"Please Print\")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1305
                    LayoutCachedTop =1335
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =1635
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6870
                    Top =1335
                    Width =2730
                    Height =300
                    Name ="EmptyCell54"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6870
                    LayoutCachedTop =1335
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =1635
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OldBorderStyle =0
                    BorderWidth =1
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1305
                    Top =1635
                    Width =8295
                    Height =345
                    TabIndex =11
                    BorderColor =0
                    ForeColor =0
                    Name ="Text489"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)=0,Null,\" \")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1305
                    LayoutCachedTop =1635
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =1980
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =3
                    LayoutGroup =1
                    BackThemeColorIndex =-1
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
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Top =1635
                    Width =1305
                    Height =345
                    TabIndex =10
                    BorderColor =0
                    ForeColor =0
                    Name ="Label490"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)=0,Null,\"Title:\")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =1635
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =1980
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =-1
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
                    OldBorderStyle =0
                    BorderWidth =1
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1305
                    Top =1980
                    Width =4935
                    Height =300
                    FontSize =8
                    TabIndex =12
                    BorderColor =0
                    ForeColor =5855577
                    Name ="Label495"
                    ControlSource ="=IIf(Nz([CoLesseeID],0)=0,Null,\"Please Print\")"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1305
                    LayoutCachedTop =1980
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =2280
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Top =1980
                    Width =1305
                    Height =300
                    Name ="EmptyCell62"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedTop =1980
                    LayoutCachedWidth =1305
                    LayoutCachedHeight =2280
                    RowStart =6
                    RowEnd =6
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6240
                    Top =1980
                    Width =630
                    Height =300
                    Name ="EmptyCell64"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6240
                    LayoutCachedTop =1980
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =2280
                    RowStart =6
                    RowEnd =6
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =6870
                    Top =1980
                    Width =2730
                    Height =300
                    Name ="EmptyCell66"
                    GroupTable =2
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6870
                    LayoutCachedTop =1980
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =2280
                    RowStart =6
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =2
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10140
                    Top =30
                    Width =819
                    Height =479
                    TabIndex =13
                    BorderColor =0
                    ForeColor =0
                    Name ="ScheduleID"
                    ControlSource ="ScheduleID"
                    GridlineColor =0

                    LayoutCachedLeft =10140
                    LayoutCachedTop =30
                    LayoutCachedWidth =10959
                    LayoutCachedHeight =509
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
