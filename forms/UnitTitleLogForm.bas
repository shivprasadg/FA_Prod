Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =163
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8588
    DatasheetFontHeight =11
    ItemSuffix =50
    Left =14745
    Top =17400
    Right =22785
    Bottom =20250
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xffc63b1e7156e640
    End
    RecordSource ="SELECT Units.UnitID , UnitTitleLog.UnitTitleLogDate , IIf([entrytypeid]=1,'In','"
        "Out') AS Status, UnitTitleLog.TitleNumber , UnitTitleLog.UnitTitleLogId, UnitTit"
        "leLog.UnitRef, UnitTitleLog.TitleSt , UnitTitleLog.TitleSt AS TitleState, UnitTi"
        "tleLog.EntryReason AS Reason , IIf([ExpectedReturn]=-1,'Y','N') AS RTN , UnitTit"
        "leLog.UserID FROM Units INNER JOIN UnitTitleLog ON Units.UnitID = UnitTitleLog.U"
        "nitRef Where Units.UnitID=363218 AND EntryReasonID <> 41 ORDER BY UnitTitleLog.U"
        "nitTitleLogDate DESC , UnitTitleLog.UnitTitleLogId DESC;"
    Caption ="Unit Title Log"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =255
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =120
                    Width =960
                    Height =255
                    FontSize =9
                    Name ="Label1"
                    Caption ="Date"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =255
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1080
                    Width =540
                    Height =255
                    FontSize =9
                    Name ="Label2"
                    Caption ="Status"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1080
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =255
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =1620
                    Width =2685
                    Height =255
                    FontSize =9
                    Name ="Label3"
                    Caption ="Reason"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1620
                    LayoutCachedWidth =4305
                    LayoutCachedHeight =255
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =4305
                    Width =2100
                    Height =255
                    FontSize =9
                    Name ="Label4"
                    Caption ="Title Number"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =4305
                    LayoutCachedWidth =6405
                    LayoutCachedHeight =255
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =75
                    Width =0
                    Height =255
                    Name ="Label5"
                    Caption ="Asset Id:"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =75
                    LayoutCachedWidth =75
                    LayoutCachedHeight =255
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6405
                    Width =495
                    Height =255
                    FontSize =9
                    Name ="Label9"
                    Caption ="State"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6405
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =255
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8580
                    Width =0
                    Height =255
                    Name ="EmptyCell12"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8580
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =255
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =6900
                    Width =435
                    Height =255
                    FontSize =9
                    Name ="Label14"
                    Caption ="Rtn?"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =6900
                    LayoutCachedWidth =7335
                    LayoutCachedHeight =255
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7335
                    Width =345
                    Height =255
                    Name ="EmptyCell43"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7335
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =255
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =7680
                    Width =900
                    Height =255
                    Name ="EmptyCell48"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =7680
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =255
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin Section
            Height =285
            Name ="Detail"
            AlternateBackColor =15590879
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =75
                    Width =0
                    Height =285
                    Name ="UnitID"
                    ControlSource ="UnitID"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0

                    LayoutCachedLeft =75
                    LayoutCachedWidth =75
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Width =960
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="UnitTitleLogDate"
                    ControlSource ="UnitTitleLogDate"
                    Format ="mm/dd/yy"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =120
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1080
                    Width =540
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Status"
                    ControlSource ="Status"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1080
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1620
                    Width =2685
                    Height =285
                    ColumnWidth =3465
                    FontSize =10
                    TabIndex =3
                    Name ="Reason"
                    ControlSource ="Reason"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1620
                    LayoutCachedWidth =4305
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4305
                    Width =2100
                    Height =285
                    ColumnWidth =1605
                    FontSize =10
                    TabIndex =4
                    Name ="TitleNumber"
                    ControlSource ="TitleNumber"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =4305
                    LayoutCachedWidth =6405
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8580
                    Width =0
                    Height =285
                    FontWeight =700
                    TabIndex =9
                    ForeColor =5026082
                    Name ="UnitTitleLogId"
                    ControlSource ="UnitTitleLogId"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =8580
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =9
                    ColumnEnd =9
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6405
                    Width =495
                    Height =285
                    FontSize =10
                    TabIndex =5
                    Name ="TitleState"
                    ControlSource ="TitleState"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6405
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6900
                    Width =435
                    Height =285
                    FontSize =10
                    TabIndex =6
                    Name ="RTN"
                    ControlSource ="RTN"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6900
                    LayoutCachedWidth =7335
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =7335
                    Width =345
                    Height =285
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    Name ="DeleteRecord"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete This Record"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000003255d654 ,
                        0x3255d6e73255d68d000000000000000000000000000000000000000000000000 ,
                        0x3255d6423255d67500000000000000000000000000000000000000003255d68d ,
                        0x3255d6ff3255d6ff3255d660000000000000000000000000000000003255d642 ,
                        0x3255d6f63255d64b00000000000000000000000000000000000000003255d61b ,
                        0x3255d6ea3255d6ff3255d6f93255d64500000000000000003255d6423255d6f6 ,
                        0x3255d6b100000000000000000000000000000000000000000000000000000000 ,
                        0x3255d63f3255d6f63255d6ff3255d6f63255d6423255d6423255d6f63255d6f0 ,
                        0x3255d61b0000000000000000000000000000000072727296727272f0727272a2 ,
                        0x000000003255d6363255d6e73255d6ff3255d6f63255d6f63255d6ff3255d64b ,
                        0x0000000000000000000000000000000000000000727272cc727272ff727272e1 ,
                        0x72727203000000003255d6213255d6f03255d6ff3255d6ff3255d6b700000000 ,
                        0x0000000000000000000000000000000000000000000000007272722a00000000 ,
                        0x000000003255d6213255d6c03255d6ff3255d6ff3255d6ff3255d6ff3255d6ab ,
                        0x3255d60c00000000000000000000000000000000727272187272723372727203 ,
                        0x3255d65a3255d6f33255d6ff3255d6ff3255d6873255d62a3255d6ae3255d6ff ,
                        0x3255d6de3255d63f0000000000000000000000007272727e727272ff72727212 ,
                        0x3255d6e43255d6ff3255d6ea3255d6570000000000000000000000003255d636 ,
                        0x3255d68d3255d62700000000000000000000000072727287727272ff72727269 ,
                        0x3255d6393255d6873255d60f0000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000072727293727272ff727272a5 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000072727299727272ff727272b1 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272a5727272ff727272ba ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272b1727272ff727272c3 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272b7727272ff727272d2 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000727272c3727272ff727272db ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =7335
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =15527148
                    BorderColor =10855845
                    HoverThemeColorIndex =2
                    HoverTint =40.0
                    PressedThemeColorIndex =1
                    PressedShade =75.0
                    HoverForeThemeColorIndex =0
                    PressedForeThemeColorIndex =0
                    GroupTable =1
                    QuickStyle =22
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7680
                    Width =900
                    Height =285
                    FontSize =8
                    TabIndex =8
                    Name ="UserID"
                    ControlSource ="UserID"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =7680
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =8
                    ColumnEnd =8
                    LayoutGroup =1
                    ThemeFontIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "UnitTitleLogForm.cls"
