Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ViewsAllowed =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10230
    DatasheetFontHeight =9
    ItemSuffix =48
    Left =330
    Top =6435
    Right =10440
    Bottom =9030
    RecSrcDt = Begin
        0xf3bbea7eece2e440
    End
    RecordSource ="vw_ATFRev_Leesee"
    Caption ="ATFRevLeesee"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Segoe UI"
    AllowFormView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
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
        Begin FormHeader
            Height =0
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin Section
            Height =8362
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1812
                    Top =60
                    Width =2460
                    Height =330
                    ColumnWidth =1530
                    ColumnOrder =0
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1812
                    LayoutCachedTop =60
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =390
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    GroupTable =1
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =60
                            Width =1620
                            Height =330
                            Name ="ID_Label"
                            Caption ="ID"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =60
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =390
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1812
                    Top =408
                    Width =2460
                    Height =336
                    ColumnWidth =3000
                    ColumnOrder =1
                    TabIndex =1
                    Name ="ATFID"
                    ControlSource ="ATFID"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1812
                    LayoutCachedTop =408
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =744
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =408
                            Width =1620
                            Height =336
                            Name ="Application_Label"
                            Caption ="ATFID"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =408
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =744
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1812
                    Top =756
                    Width =2460
                    Height =336
                    ColumnWidth =3000
                    ColumnOrder =2
                    TabIndex =2
                    Name ="tableName"
                    ControlSource ="tableName"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1812
                    LayoutCachedTop =756
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =1092
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =756
                            Width =1620
                            Height =336
                            Name ="tableName_Label"
                            Caption ="tableName"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =756
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =1092
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1812
                    Top =1104
                    Width =2460
                    Height =336
                    ColumnWidth =1875
                    ColumnOrder =4
                    TabIndex =3
                    Name ="ActionDate"
                    ControlSource ="ActionDate"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1812
                    LayoutCachedTop =1104
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =1440
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =1104
                            Width =1620
                            Height =336
                            Name ="ActionDate_Label"
                            Caption ="ActionDate"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =1104
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =1440
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1812
                    Top =1452
                    Width =2460
                    Height =336
                    ColumnWidth =915
                    ColumnOrder =5
                    TabIndex =4
                    Name ="Action"
                    ControlSource ="Action"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1812
                    LayoutCachedTop =1452
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =1788
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =1452
                            Width =1620
                            Height =336
                            Name ="Action_Label"
                            Caption ="Action"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =1452
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =1788
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1812
                    Top =1800
                    Width =2460
                    Height =336
                    ColumnWidth =1005
                    ColumnOrder =6
                    TabIndex =5
                    Name ="SysUser"
                    ControlSource ="SysUser"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1812
                    LayoutCachedTop =1800
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =2136
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =1800
                            Width =1620
                            Height =336
                            Name ="SysUser_Label"
                            Caption ="SysUser"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =1800
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =2136
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1812
                    Top =2148
                    Width =2460
                    Height =336
                    ColumnWidth =1620
                    ColumnOrder =7
                    TabIndex =6
                    Name ="FieldName"
                    ControlSource ="FieldName"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1812
                    LayoutCachedTop =2148
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =2484
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =2148
                            Width =1620
                            Height =336
                            Name ="FieldName_Label"
                            Caption ="FieldName"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =2148
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =2484
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1812
                    Top =2496
                    Width =2460
                    Height =336
                    ColumnWidth =1845
                    ColumnOrder =8
                    TabIndex =7
                    Name ="PrevValue"
                    ControlSource ="PrevValue"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1812
                    LayoutCachedTop =2496
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =2832
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =2496
                            Width =1620
                            Height =336
                            Name ="PrevValue_Label"
                            Caption ="PrevValue"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =2496
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =2832
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1812
                    Top =2844
                    Width =2460
                    Height =336
                    ColumnWidth =3000
                    ColumnOrder =9
                    TabIndex =8
                    Name ="NewValue"
                    ControlSource ="NewValue"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1812
                    LayoutCachedTop =2844
                    LayoutCachedWidth =4272
                    LayoutCachedHeight =3180
                    RowStart =8
                    RowEnd =8
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =2844
                            Width =1620
                            Height =336
                            Name ="NewValue_Label"
                            Caption ="NewValue"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =180
                            LayoutCachedTop =2844
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =3180
                            RowStart =8
                            RowEnd =8
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
' See "ATFRevATF.cls"
