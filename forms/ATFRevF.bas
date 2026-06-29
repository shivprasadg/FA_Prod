Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =17595
    DatasheetFontHeight =11
    ItemSuffix =193
    Left =5205
    Top =12090
    Right =22470
    Bottom =19080
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0xa63b1363d6ebe440
    End
    RecordSource ="SELECT ATF_ChangeReasons.Action, ATF_ChangeReasons.Reason, ATF_ChangeReasons.Cha"
        "ngeDate, ATF_ChangeReasons.FieldName, ATF_ChangeReasons.SchNo, ATF_ChangeReasons"
        ".GroupNo, ATF_ChangeReasons.FieldGroup, ATF_ChangeReasons.OldValue, ATF_ChangeRe"
        "asons.NewValue, ATF_ChangeReasons.TableName, ATF_ChangeReasons.RevID, ATF_Change"
        "Reasons.SchID, ATF_ChangeReasons.GroupID, ATF_ChangeReasons.UnitID, ATF_ChangeRe"
        "asons.SysUser, ATF_ChangeReasons.ATFID, ATF_ChangeReasons.ATFUnitSpecID, ATF_Cha"
        "ngeReasons.ATFGroupSpecID, ATF_ChangeReasons.ATFSchSpecID, ATF_ChangeReasons.Rev"
        "iewedBy, ATF_ChangeReasons.ReviewedDate, ATF_ChangeReasons.INDX FROM ATF_ChangeR"
        "easons ORDER BY ATF_ChangeReasons.RevID DESC;"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    Moveable =0
    FilterOnLoad =255
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =7050
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =247
                    Left =3990
                    Top =15
                    Width =13230
                    Height =4890
                    Name ="ScheduleRevisionLog"
                    SourceObject ="Form.ScheduleRevisionLog"
                    LinkChildFields ="SchID"
                    LinkMasterFields ="schref"

                    LayoutCachedLeft =3990
                    LayoutCachedTop =15
                    LayoutCachedWidth =17220
                    LayoutCachedHeight =4905
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2196
                    Top =4776
                    Width =516
                    Height =192
                    FontSize =9
                    TabIndex =12
                    ForeColor =5026082
                    Name ="INDX"
                    ControlSource ="INDX"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2196
                    LayoutCachedTop =4776
                    LayoutCachedWidth =2712
                    LayoutCachedHeight =4968
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1668
                    Top =4776
                    Width =516
                    Height =192
                    FontSize =9
                    TabIndex =11
                    ForeColor =5026082
                    Name ="schref"
                    ControlSource ="SchID"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1668
                    LayoutCachedTop =4776
                    LayoutCachedWidth =2184
                    LayoutCachedHeight =4968
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1656
                    Top =1056
                    Width =2076
                    Height =312
                    FontSize =10
                    TabIndex =3
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    Format ="Short Date"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Revision Date Required"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1656
                    LayoutCachedTop =1056
                    LayoutCachedWidth =3732
                    LayoutCachedHeight =1368
                    RowStart =2
                    RowEnd =2
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =30
                    Top =4980
                    Width =17340
                    Height =2070
                    FontSize =10
                    TabIndex =17
                    Name ="Reason"
                    ControlSource ="Reason"
                    FontName ="Segoe UI"
                    OnChange ="[Event Procedure]"
                    TextFormat =1

                    LayoutCachedLeft =30
                    LayoutCachedTop =4980
                    LayoutCachedWidth =17370
                    LayoutCachedHeight =7050
                    RowEnd =6
                    ColumnStart =3
                    ColumnEnd =3
                    BackThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =30
                            Top =4770
                            Width =1080
                            Height =195
                            FontSize =9
                            FontWeight =700
                            Name ="Label80"
                            Caption ="Notes"
                            LayoutCachedLeft =30
                            LayoutCachedTop =4770
                            LayoutCachedWidth =1110
                            LayoutCachedHeight =4965
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1656
                    Top =1596
                    Width =2076
                    Height =312
                    FontSize =10
                    TabIndex =4
                    Name ="ReviewedDate"
                    ControlSource ="ReviewedDate"
                    Format ="Short Date"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1656
                    LayoutCachedTop =1596
                    LayoutCachedWidth =3732
                    LayoutCachedHeight =1908
                    RowStart =4
                    RowEnd =4
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1656
                    Top =2136
                    Width =2076
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="ReviewedBy"
                    ControlSource ="ReviewedBy"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TBL_Employees.[User ID], TBL_Employees.[Employee Name], TBL_Employees.Emp"
                        "Int FROM TBL_Employees WHERE (((TBL_Employees.[User ID])<>\"ALL\") AND ((TBL_Emp"
                        "loyees.[Active User])=Yes)) ORDER BY TBL_Employees.EmpInt;"
                    FontName ="Segoe UI"
                    Format =">"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1656
                    LayoutCachedTop =2136
                    LayoutCachedWidth =3732
                    LayoutCachedHeight =2436
                    RowStart =6
                    RowEnd =6
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2160
                    Left =1656
                    Top =2664
                    Width =2076
                    Height =300
                    FontSize =10
                    TabIndex =6
                    Name ="SysUserCbo"
                    ControlSource ="SysUser"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TBL_Employees.[User ID], TBL_Employees.[Employee Name], TBL_Employees.Emp"
                        "Int FROM TBL_Employees WHERE (((TBL_Employees.[User ID])<>\"ALL\") AND ((TBL_Emp"
                        "loyees.[Active User])=Yes)) ORDER BY TBL_Employees.EmpInt;"
                    ColumnWidths ="576;1728;0"
                    ValidationRule ="Is Not Null"
                    ValidationText ="NT By is required (even for 1st ATF)"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    AllowValueListEdits =0

                    LayoutCachedLeft =1656
                    LayoutCachedTop =2664
                    LayoutCachedWidth =3732
                    LayoutCachedHeight =2964
                    RowStart =8
                    RowEnd =8
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    Width =2025
                    Height =345
                    FontSize =14
                    FontWeight =700
                    ForeColor =11830108
                    Name ="Label15"
                    Caption ="ATF Revisions                                "
                    FontName ="Segoe UI"
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =345
                    BackThemeColorIndex =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =30
                    Top =360
                    Width =1545
                    Height =4305
                    FontSize =9
                    TabIndex =1
                    Name ="RevList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ATF_ChangeReasons.INDX, ATF_ChangeReasons.RevID AS Rev, Format([ChangeDat"
                        "e],\"mm/dd/yy\") AS [Date] FROM ATF_ChangeReasons WHERE (((ATF_ChangeReasons.Sch"
                        "ID)=[forms]![schedule_form]![schid])) ORDER BY ATF_ChangeReasons.RevID DESC , Fo"
                        "rmat([ChangeDate],\"mm/dd/yy\") DESC;"
                    ColumnWidths ="0;432;1008"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =30
                    LayoutCachedTop =360
                    LayoutCachedWidth =1575
                    LayoutCachedHeight =4665
                    RowEnd =4
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =1620
                    Top =3150
                    Width =2115
                    Height =285
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    Name ="SaveUpdate"
                    Caption =" Save Update"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffffb17d4affb17d4affb17d4aff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffff00000000 ,
                        0x000000000000000000000000727272ffffffffffffffffffffffffffffffffff ,
                        0xd5abc3ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff00000000727272ff727272ff727272ff727272ff00000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffffffffffffffffffffffffffffffffffab5989ff ,
                        0xab5989ffab5989ff000000000000000000000000000000000000000000000000 ,
                        0xab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ffab5989ff ,
                        0xab5989ffab5989ff
                    End

                    LayoutCachedLeft =1620
                    LayoutCachedTop =3150
                    LayoutCachedWidth =3735
                    LayoutCachedHeight =3435
                    PictureCaptionArrangement =5
                    RowStart =11
                    RowEnd =11
                    ColumnEnd =1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedColor =16777215
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =1635
                    Top =3570
                    Width =2115
                    Height =285
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="AddRevision"
                    Caption ="Add Rev Note"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4af9b17d4a78b17d4a18 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a81b17d4affb17d4af3 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a27b17d4af6b17d4a03 ,
                        0xb17d4ab7b17d4a6c000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a0cb17d4ab7 ,
                        0xb17d4affb17d4affb17d4a7b0000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000000000000b17d4a87 ,
                        0xb17d4affb17d4affb17d4affb17d4a8700000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xb17d4a8db17d4affb17d4affb17d4affb17d4a93000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000b17d4a90b17d4affb17d4affb17d4aabb17d4a12b17d4a0300000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a96b17d4aabb17d4a15b17d4acfb17d4aa500000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a12b17d4acfb17d4affb17d4af000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a03b17d4aa2b17d4afcb17d4a2a00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =1635
                    LayoutCachedTop =3570
                    LayoutCachedWidth =3750
                    LayoutCachedHeight =3855
                    PictureCaptionArrangement =5
                    RowStart =10
                    RowEnd =10
                    ColumnEnd =1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    HoverThemeColorIndex =6
                    HoverTint =100.0
                    HoverShade =75.0
                    PressedColor =16777215
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =1656
                    Top =2436
                    Width =2076
                    Height =216
                    FontSize =9
                    FontWeight =700
                    Name ="Label10"
                    Caption ="Notes By"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1656
                    LayoutCachedTop =2436
                    LayoutCachedWidth =3732
                    LayoutCachedHeight =2652
                    RowStart =7
                    RowEnd =7
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =1656
                    Top =1908
                    Width =2076
                    Height =216
                    FontSize =9
                    FontWeight =700
                    Name ="Label6"
                    Caption ="Reviewer"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1656
                    LayoutCachedTop =1908
                    LayoutCachedWidth =3732
                    LayoutCachedHeight =2124
                    RowStart =5
                    RowEnd =5
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =1
                    Left =1656
                    Top =1368
                    Width =2076
                    Height =216
                    FontSize =9
                    FontWeight =700
                    Name ="Label5"
                    Caption ="Reviewed"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1656
                    LayoutCachedTop =1368
                    LayoutCachedWidth =3732
                    LayoutCachedHeight =1584
                    RowStart =3
                    RowEnd =3
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =1656
                    Top =828
                    Width =2076
                    Height =216
                    FontSize =9
                    FontWeight =700
                    Name ="Label3"
                    Caption ="Rev Date"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1656
                    LayoutCachedTop =828
                    LayoutCachedWidth =3732
                    LayoutCachedHeight =1044
                    RowStart =1
                    RowEnd =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1140
                    Top =4776
                    Width =510
                    Height =192
                    FontSize =9
                    TabIndex =10
                    ForeColor =5026082
                    Name ="Action"
                    ControlSource ="Action"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1140
                    LayoutCachedTop =4776
                    LayoutCachedWidth =1650
                    LayoutCachedHeight =4968
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2724
                    Top =4776
                    Width =516
                    Height =192
                    FontSize =9
                    TabIndex =13
                    ForeColor =5026082
                    Name ="FieldGroup"
                    ControlSource ="FieldGroup"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2724
                    LayoutCachedTop =4776
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =4968
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3252
                    Top =4776
                    Width =516
                    Height =192
                    FontSize =9
                    TabIndex =14
                    ForeColor =5026082
                    Name ="NewValue"
                    ControlSource ="NewValue"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3252
                    LayoutCachedTop =4776
                    LayoutCachedWidth =3768
                    LayoutCachedHeight =4968
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3780
                    Top =4776
                    Width =516
                    Height =192
                    FontSize =9
                    TabIndex =16
                    ForeColor =5026082
                    Name ="SchNo"
                    ControlSource ="SchNo"
                    GroupTable =7
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =3780
                    LayoutCachedTop =4776
                    LayoutCachedWidth =4296
                    LayoutCachedHeight =4968
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =7
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =2160
                    Left =1656
                    Top =2988
                    Width =2076
                    Height =12
                    FontSize =10
                    TabIndex =7
                    Name ="ImpactsCBO"
                    ControlSource ="FieldGroup"
                    RowSourceType ="Value List"
                    RowSource ="Rent;Cost;Vendor;Schedule;Lease Terms;Unit Spec;Economics"
                    ColumnWidths ="2016"
                    ValidationRule ="Is Not Null"
                    ValidationText ="NT By is required (even for 1st ATF)"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =1656
                    LayoutCachedTop =2988
                    LayoutCachedWidth =3732
                    LayoutCachedHeight =3000
                    RowStart =10
                    RowEnd =10
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    Left =1656
                    Top =2964
                    Width =2076
                    Height =12
                    FontSize =9
                    FontWeight =700
                    Name ="Label164"
                    Caption ="Impacts..."
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =1656
                    LayoutCachedTop =2964
                    LayoutCachedWidth =3732
                    LayoutCachedHeight =2976
                    RowStart =9
                    RowEnd =9
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =1620
                    Top =4020
                    Width =2115
                    Height =285
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    Name ="DeleteRev"
                    Caption =" Delete Rev"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000003255d6943255d68300000000000000000000000000000000 ,
                        0x3255d6043255d681000000000000000000000000000000000000000000000000 ,
                        0x000000003255d60e3255d6f43255d6f63255d66a000000000000000000000000 ,
                        0x3255d6c13255d6b7000000000000000000000000000000000000000000000000 ,
                        0x00000000000000003255d68d3255d6f83255d6f43255d652000000003255d6ba ,
                        0x3255d6ec3255d618000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000003255d6943255d6f63255d6f43255d6d23255d6f6 ,
                        0x3255d65900000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000003255d68d3255d6f83255d6f83255d6cd ,
                        0x0000000000000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff00000000000000003255d6753255d6f13255d6f83255d6eb3255d6f8 ,
                        0x3255d6c63255d614727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffff423255d69b3255d6f83255d6f63255d6a1ffffff033255d65e ,
                        0x3255d6d93255d6ad727272ffffffffffb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affcead8da33255d6623255d6bf546acb52c9a582ccb48250ffc9a581cf ,
                        0xffffff3300000000727272ffffffffffb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb58454fecca988bbcdab8bb1ba8c5ffbb17d4affb17d4affb17d4aff ,
                        0xffffffff727272ff727272ffffffffffb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xffffffff727272ff727272ffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =1620
                    LayoutCachedTop =4020
                    LayoutCachedWidth =3735
                    LayoutCachedHeight =4305
                    PictureCaptionArrangement =5
                    RowStart =10
                    RowEnd =10
                    ColumnEnd =1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    HoverColor =4578793
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =16777215
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2736
                    Top =360
                    Width =996
                    Height =465
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    Name ="RevID"
                    ControlSource ="RevID"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Revision Number is Required"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Segoe UI"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedLeft =2736
                    LayoutCachedTop =360
                    LayoutCachedWidth =3732
                    LayoutCachedHeight =825
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1656
                            Top =360
                            Width =1068
                            Height =465
                            FontSize =9
                            FontWeight =700
                            Name ="Label2"
                            Caption ="Revision \015\012Number"
                            GroupTable =1
                            LeftPadding =0
                            TopPadding =0
                            RightPadding =0
                            BottomPadding =0
                            LayoutCachedLeft =1656
                            LayoutCachedTop =360
                            LayoutCachedWidth =2724
                            LayoutCachedHeight =825
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =1620
                    Top =4425
                    Width =2115
                    Height =285
                    FontSize =8
                    FontWeight =700
                    TabIndex =18
                    Name ="CleanUpNotes"
                    Caption =" Clean Up Notes"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000007ec8ed817ec8edff ,
                        0x000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff000000000000000000000000000000007ec8ed817ec8edff7ec8edff ,
                        0x000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff0000000000000000000000007ec8ed817ec8edff7ec8edff7ec8edff ,
                        0x000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff00000000000000007ec8ed817ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff00000000000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff00000000000000000000000000000000000000000000000000000000 ,
                        0x000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff00000000000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff00000000000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff00000000000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff00000000000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff00000000000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff00000000000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff00000000000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff00000000000000007ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff7ec8edff ,
                        0x7ec8edff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =1620
                    LayoutCachedTop =4425
                    LayoutCachedWidth =3735
                    LayoutCachedHeight =4710
                    PictureCaptionArrangement =5
                    RowStart =10
                    RowEnd =10
                    ColumnEnd =1
                    ForeThemeColorIndex =6
                    ForeTint =100.0
                    ForeShade =50.0
                    Shape =0
                    Gradient =0
                    BackColor =-2147483633
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    BorderShade =50.0
                    ThemeFontIndex =-1
                    HoverColor =4578793
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =16777215
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    QuickStyle =30
                    QuickStyleMask =-629
                End
            End
        End
    End
End
CodeBehindForm
' See "ATFRevF.cls"
