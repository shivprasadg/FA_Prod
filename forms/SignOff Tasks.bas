Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12960
    DatasheetFontHeight =11
    ItemSuffix =86
    Left =6540
    Top =4965
    Right =19785
    Bottom =10290
    OrderBy ="[DepartmentContacts].[SignOffLevel]"
    RecSrcDt = Begin
        0x41926cc0ac59e540
    End
    RecordSource ="DepartmentContacts"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =705
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            BackShade =85.0
            Begin
                Begin Image
                    PictureTiling = NotDefault
                    SizeMode =0
                    PictureAlignment =0
                    PictureType =2
                    Top =465
                    Width =12960
                    Height =240
                    BorderColor =10921638
                    Name ="Image439"
                    Picture ="LtGrey2_Shadow"
                    GridlineColor =10921638
                    HorizontalAnchor =2

                    LayoutCachedTop =465
                    LayoutCachedWidth =12960
                    LayoutCachedHeight =705
                    TabIndex =1
                    RowStart =1
                    RowEnd =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Top =30
                    Width =10470
                    Height =465
                    FontSize =18
                    Name ="Auto_Header0"
                    Caption ="Sign Off Events"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    HorizontalAnchor =2
                    LayoutCachedTop =30
                    LayoutCachedWidth =10470
                    LayoutCachedHeight =495
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11638
                    Width =823
                    Height =285
                    ColumnOrder =0
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="INDX"
                    ControlSource ="INDX"
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =11638
                    LayoutCachedWidth =12461
                    LayoutCachedHeight =285
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =11010
                            Width =823
                            Height =285
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label0"
                            Caption ="INDX"
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =11010
                            LayoutCachedWidth =11833
                            LayoutCachedHeight =285
                        End
                    End
                End
            End
        End
        Begin Section
            Height =4605
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2248
                    Top =660
                    Width =2640
                    Height =360
                    ColumnWidth =1770
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Department"
                    ControlSource ="Department"
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2248
                    LayoutCachedTop =660
                    LayoutCachedWidth =4888
                    LayoutCachedHeight =1020
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =225
                            Top =660
                            Width =2023
                            Height =360
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label3"
                            Caption ="Department"
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =225
                            LayoutCachedTop =660
                            LayoutCachedWidth =2248
                            LayoutCachedHeight =1020
                            RowStart =1
                            RowEnd =1
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3240
                    Left =2248
                    Top =1020
                    Width =2640
                    Height =360
                    ColumnWidth =2040
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="PrimaryContactID"
                    ControlSource ="PrimaryContactID"
                    RowSourceType ="Table/Query"
                    RowSource ="Select ID, [Employee Name] From TBL_Employees WHERE [Active User] =True Order By"
                        " [Employee Name]"
                    ColumnWidths ="0;2880"
                    BottomPadding =150
                    GridlineColor =10921638
                    AllowValueListEdits =0
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22

                    LayoutCachedLeft =2248
                    LayoutCachedTop =1020
                    LayoutCachedWidth =4888
                    LayoutCachedHeight =1380
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =225
                            Top =1020
                            Width =2023
                            Height =360
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label6"
                            Caption ="From"
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =225
                            LayoutCachedTop =1020
                            LayoutCachedWidth =2248
                            LayoutCachedHeight =1380
                            RowStart =2
                            RowEnd =2
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3240
                    Left =2248
                    Top =1380
                    Width =2640
                    Height =360
                    ColumnWidth =1770
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="SecondaryContactID"
                    ControlSource ="SecondaryContactID"
                    RowSourceType ="Table/Query"
                    RowSource ="Select ID, [Employee Name] From TBL_Employees WHERE [Active User] =True Order By"
                        " [Employee Name]"
                    ColumnWidths ="0;2880"
                    BottomPadding =150
                    GridlineColor =10921638
                    AllowValueListEdits =0
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22

                    LayoutCachedLeft =2248
                    LayoutCachedTop =1380
                    LayoutCachedWidth =4888
                    LayoutCachedHeight =1740
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =1
                            Left =225
                            Top =1380
                            Width =2023
                            Height =360
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label9"
                            Caption ="To"
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =225
                            LayoutCachedTop =1380
                            LayoutCachedWidth =2248
                            LayoutCachedHeight =1740
                            RowStart =3
                            RowEnd =3
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2248
                    Top =1845
                    Width =4725
                    Height =945
                    ColumnWidth =8535
                    TabIndex =5
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="DepartmentEmail"
                    ControlSource ="DepartmentEmail"
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2248
                    LayoutCachedTop =1845
                    LayoutCachedWidth =6973
                    LayoutCachedHeight =2790
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =225
                            Top =1845
                            Width =2023
                            Height =360
                            LeftMargin =44
                            TopMargin =22
                            RightMargin =44
                            BottomMargin =22
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label12"
                            Caption ="CC Group"
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =225
                            LayoutCachedTop =1845
                            LayoutCachedWidth =2248
                            LayoutCachedHeight =2205
                            RowStart =4
                            RowEnd =4
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7230
                    Top =195
                    Width =5685
                    Height =4305
                    ColumnWidth =4350
                    TabIndex =2
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SignOffTaskList"
                    ControlSource ="SignOffTaskList"
                    BottomPadding =150
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedLeft =7230
                    LayoutCachedTop =195
                    LayoutCachedWidth =12915
                    LayoutCachedHeight =4500
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2248
                    Top =300
                    Width =2640
                    Height =360
                    TabIndex =4
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SignOffLevel"
                    ControlSource ="SignOffLevel"
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2248
                    LayoutCachedTop =300
                    LayoutCachedWidth =4888
                    LayoutCachedHeight =660
                    RowStart =3
                    RowEnd =3
                    ColumnStart =3
                    ColumnEnd =3
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2250
                    Top =3120
                    Width =4320
                    Height =810
                    TabIndex =6
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SignOffGroupEmail"
                    ControlSource ="SignOffGroupEmail"
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2250
                    LayoutCachedTop =3120
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =3930
                    RowStart =4
                    RowEnd =4
                    ColumnStart =3
                    ColumnEnd =3
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =5205
                    Top =195
                    Width =2023
                    Height =810
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label15"
                    Caption ="SignOffTaskList"
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =5205
                    LayoutCachedTop =195
                    LayoutCachedWidth =7228
                    LayoutCachedHeight =1005
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =225
                    Top =300
                    Width =2023
                    Height =360
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label18"
                    Caption ="SignOffLevel"
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =225
                    LayoutCachedTop =300
                    LayoutCachedWidth =2248
                    LayoutCachedHeight =660
                    RowStart =3
                    RowEnd =3
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =210
                    Top =3120
                    Width =2023
                    Height =810
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label21"
                    Caption ="SignOffGroupEmail"
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =210
                    LayoutCachedTop =3120
                    LayoutCachedWidth =2233
                    LayoutCachedHeight =3930
                    RowStart =4
                    RowEnd =4
                    ColumnStart =2
                    ColumnEnd =2
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
