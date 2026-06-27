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
    ScrollBars =2
    TabularCharSet =163
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8618
    DatasheetFontHeight =11
    ItemSuffix =41
    Left =14715
    Top =14235
    Right =22755
    Bottom =16200
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xa1f8311e7156e640
    End
    RecordSource =" SELECT H.*, LOld.LocationName AS PrevLocation, LNew.LocationName AS NewLocation"
        " FROM (LocationChangeHistory AS H INNER JOIN Locations AS LOld ON H.OldLocationI"
        "d = LOld.LocationID) INNER JOIN Locations AS LNew ON H.NewLocationId = LNew.Loca"
        "tionID WHERE AssetID =363218;"
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
                    Width =1005
                    Height =255
                    FontSize =9
                    Name ="Label1"
                    Caption ="Date"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =120
                    LayoutCachedWidth =1125
                    LayoutCachedHeight =255
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1125
                    Width =1980
                    Height =255
                    FontSize =9
                    Name ="Label2"
                    Caption ="Previous Location           "
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1125
                    LayoutCachedWidth =3105
                    LayoutCachedHeight =255
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
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
                Begin EmptyCell
                    Left =8610
                    Width =0
                    Height =255
                    Name ="EmptyCell12"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8610
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =255
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    FontItalic = NotDefault
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6045
                    Width =1986
                    Height =255
                    ColumnOrder =0
                    FontSize =7
                    FontWeight =700
                    BorderColor =0
                    Name ="HideShow"
                    ControlSource ="=\"Hide/Show Details\""
                    OnGotFocus ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =6045
                    LayoutCachedWidth =8031
                    LayoutCachedHeight =255
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    ThemeFontIndex =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4575
                            Width =1470
                            Height =255
                            FontSize =9
                            Name ="Label4"
                            Caption ="New Location"
                            FontName ="Segoe UI"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =4575
                            LayoutCachedWidth =6045
                            LayoutCachedHeight =255
                            ColumnStart =4
                            ColumnEnd =4
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GroupTable =1
                        End
                    End
                End
                Begin EmptyCell
                    Left =3105
                    Width =1476
                    Height =255
                    Name ="EmptyCell36"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3105
                    LayoutCachedWidth =4581
                    LayoutCachedHeight =255
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8025
                    Width =570
                    Height =255
                    Name ="EmptyCell38"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =0
                    LayoutCachedLeft =8025
                    LayoutCachedWidth =8595
                    LayoutCachedHeight =255
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =300
            Name ="Detail"
            AlternateBackColor =16773853
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =75
                    Width =0
                    Height =285
                    Name ="AssetId"
                    ControlSource ="AssetId"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =15

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
                    Width =1005
                    Height =285
                    FontSize =9
                    TabIndex =1
                    Name ="LocationChangeDate"
                    ControlSource ="LocationChangeDate"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =120
                    LayoutCachedWidth =1125
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =300
                    Width =1005
                    Height =0
                    FontSize =9
                    FontWeight =700
                    TabIndex =6
                    Name ="FAUserName"
                    ControlSource ="FAUserName"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =120
                    LayoutCachedTop =300
                    LayoutCachedWidth =1125
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =5
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1125
                    Top =300
                    Width =6906
                    Height =0
                    ColumnWidth =3465
                    FontSize =9
                    TabIndex =7
                    BackColor =13434879
                    Name ="Reason"
                    ControlSource ="Reason"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1125
                    LayoutCachedTop =300
                    LayoutCachedWidth =8031
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1125
                    Width =3456
                    Height =285
                    ColumnWidth =1605
                    FontSize =10
                    TabIndex =2
                    Name ="PrevLocation"
                    ControlSource ="PrevLocation"
                    FontName ="Segoe UI"
                    OnLostFocus ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =1125
                    LayoutCachedWidth =4581
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8610
                    Width =0
                    Height =285
                    FontWeight =700
                    TabIndex =5
                    ForeColor =5026082
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =8610
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4575
                    Width =3456
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="NewLocation"
                    ControlSource ="NewLocation"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =15

                    LayoutCachedLeft =4575
                    LayoutCachedWidth =8031
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =75
                    Top =300
                    Width =0
                    Height =0
                    Name ="EmptyCell20"
                    GroupTable =1
                    TopPadding =0
                    RightPadding =38
                    BottomPadding =0
                    LayoutCachedLeft =75
                    LayoutCachedTop =300
                    LayoutCachedWidth =75
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =8610
                    Top =300
                    Width =0
                    Height =0
                    Name ="EmptyCell26"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =8610
                    LayoutCachedTop =300
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =8025
                    Width =570
                    Height =285
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="DeleteRecord"
                    Caption ="DeleteSelected"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete This Record"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =15
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

                    LayoutCachedLeft =8025
                    LayoutCachedWidth =8595
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
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
                Begin EmptyCell
                    Left =8025
                    Top =300
                    Width =570
                    Height =0
                    Name ="EmptyCell40"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =15
                    BottomPadding =0
                    LayoutCachedLeft =8025
                    LayoutCachedTop =300
                    LayoutCachedWidth =8595
                    LayoutCachedHeight =300
                    RowStart =2
                    RowEnd =2
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
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
' See "UnitLocationChangeHistoryForm.cls"
