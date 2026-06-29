Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =163
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15900
    DatasheetFontHeight =11
    ItemSuffix =106
    Left =5205
    Top =12090
    Right =21015
    Bottom =17370
    ShortcutMenuBar ="SortClip"
    RecSrcDt = Begin
        0x94efa2368daee440
    End
    RecordSource ="ScheduleAmends"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            Name ="FormHeader"
        End
        Begin Section
            Height =4800
            Name ="Detail"
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5340
                    Top =480
                    Width =10500
                    Height =3885
                    TabIndex =3
                    Name ="AmendmentDesc"
                    ControlSource ="AmendmentDesc"
                    StatusBarText ="Descripiton of the Amendment"

                    LayoutCachedLeft =5340
                    LayoutCachedTop =480
                    LayoutCachedWidth =15840
                    LayoutCachedHeight =4365
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13140
                    Width =540
                    FontSize =9
                    Name ="ScheduleID"
                    ControlSource ="ScheduleID"
                    StatusBarText ="ML reference id"
                    BottomPadding =150

                    LayoutCachedLeft =13140
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =240
                    ColumnStart =4
                    ColumnEnd =4
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13740
                    Width =360
                    FontSize =9
                    TabIndex =1
                    Name ="AmendID"
                    ControlSource ="ID"
                    BottomPadding =150

                    LayoutCachedLeft =13740
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =240
                    ColumnStart =6
                    ColumnEnd =6
                    BorderThemeColorIndex =1
                    BorderShade =50.0
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =60
                    Top =480
                    Width =1680
                    Height =3960
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="RevList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ScheduleAmends.ID, ScheduleAmends.ScheduleID, ScheduleAmends.AmendmentNO "
                        "AS [No], ScheduleAmends.AmendmentDate AS [Date] FROM ScheduleAmends WHERE (((Sch"
                        "eduleAmends.ScheduleID)=[Forms]![Schedule_Form]![SchID])) ORDER BY ScheduleAmend"
                        "s.AmendmentNO DESC , ScheduleAmends.AmendmentDate DESC;"
                    ColumnWidths ="0;0;432;1138"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="List of Schedule Amendments"

                    LayoutCachedLeft =60
                    LayoutCachedTop =480
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =4440
                    RowEnd =4
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2652
                    Top =540
                    Width =2580
                    Height =270
                    FontWeight =700
                    TabIndex =4
                    Name ="AmendmentNO"
                    ControlSource ="AmendmentNO"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Revision No required"
                    DefaultValue ="0"
                    ControlTipText ="Amendment Number"
                    GroupTable =1

                    LayoutCachedLeft =2652
                    LayoutCachedTop =540
                    LayoutCachedWidth =5232
                    LayoutCachedHeight =810
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1800
                            Top =540
                            Width =780
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label38"
                            Caption ="No:"
                            GroupTable =1
                            LayoutCachedLeft =1800
                            LayoutCachedTop =540
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =810
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2652
                    Top =1236
                    Width =2580
                    Height =312
                    TabIndex =6
                    Name ="RevDt"
                    ControlSource ="AmendmentDate"
                    Format ="Short Date"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Revision Date Required"
                    ControlTipText ="Date the Amendment was enterd into ATLAAS"
                    GroupTable =1

                    LayoutCachedLeft =2652
                    LayoutCachedTop =1236
                    LayoutCachedWidth =5232
                    LayoutCachedHeight =1548
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1800
                            Top =1236
                            Width =780
                            Height =312
                            FontSize =9
                            FontWeight =700
                            Name ="Label3"
                            Caption ="Entered:"
                            GroupTable =1
                            LayoutCachedLeft =1800
                            LayoutCachedTop =1236
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =1548
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1800
                    Top =2772
                    Width =3432
                    Height =252
                    FontSize =9
                    FontWeight =700
                    TabIndex =11
                    Name ="AddRevision"
                    Caption ="Add Amendment"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add Amendment"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1800
                    LayoutCachedTop =2772
                    LayoutCachedWidth =5232
                    LayoutCachedHeight =3024
                    RowStart =6
                    RowEnd =6
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =2779219
                    BorderThemeColorIndex =6
                    BorderTint =60.0
                    ThemeFontIndex =1
                    HoverThemeColorIndex =6
                    HoverShade =75.0
                    PressedColor =2779219
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1800
                    Top =3096
                    Width =3432
                    Height =252
                    FontSize =9
                    FontWeight =700
                    TabIndex =12
                    Name ="SaveUpdate"
                    Caption ="Save Update"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save this Amendment"
                    GroupTable =1
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1800
                    LayoutCachedTop =3096
                    LayoutCachedWidth =5232
                    LayoutCachedHeight =3348
                    RowStart =7
                    RowEnd =7
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Gradient =24
                    BackThemeColorIndex =4
                    BackShade =75.0
                    BorderColor =10040879
                    ThemeFontIndex =1
                    HoverColor =12349952
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    GroupTable =1
                    QuickStyle =30
                    QuickStyleMask =-117
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =255
                    Left =60
                    Width =7290
                    Height =495
                    FontSize =18
                    ForeColor =11830108
                    Name ="Label15"
                    Caption ="Amendments                              Description"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =7350
                    LayoutCachedHeight =495
                    BackThemeColorIndex =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =4500
                    Width =1680
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =13
                    Name ="DeleteSelected"
                    Caption ="Delete Selected"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Delete Amendment"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =60
                    LayoutCachedTop =4500
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =4755
                    ColumnEnd =1
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    BackColor =1117820
                    BorderThemeColorIndex =6
                    BorderTint =60.0
                    ThemeFontIndex =1
                    HoverColor =2366701
                    PressedColor =1117820
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =30
                    QuickStyleMask =-117
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2652
                    Top =888
                    Width =2580
                    Height =276
                    TabIndex =5
                    Name ="AmendmentName"
                    ControlSource ="AmendmentName"
                    Format ="Short Date"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Revision Date Required"
                    ControlTipText ="Amendment Document Name"
                    GroupTable =1

                    LayoutCachedLeft =2652
                    LayoutCachedTop =888
                    LayoutCachedWidth =5232
                    LayoutCachedHeight =1164
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1800
                            Top =888
                            Width =780
                            Height =276
                            FontSize =9
                            FontWeight =700
                            Name ="Label69"
                            Caption ="Name:"
                            GroupTable =1
                            LayoutCachedLeft =1800
                            LayoutCachedTop =888
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =1164
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2652
                    Top =2004
                    Width =2580
                    Height =312
                    TabIndex =8
                    Name ="AmendmentSent"
                    ControlSource ="AmendmentSent"
                    Format ="Short Date"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Revision Date Required"
                    ControlTipText ="Date the Amendment was sent out"
                    GroupTable =1

                    LayoutCachedLeft =2652
                    LayoutCachedTop =2004
                    LayoutCachedWidth =5232
                    LayoutCachedHeight =2316
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1800
                            Top =2004
                            Width =780
                            Height =312
                            FontSize =9
                            FontWeight =700
                            Name ="Label81"
                            Caption ="Sent:"
                            GroupTable =1
                            LayoutCachedLeft =1800
                            LayoutCachedTop =2004
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =2316
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2652
                    Top =1620
                    Width =2580
                    Height =312
                    TabIndex =7
                    Name ="AmendmentEffective"
                    ControlSource ="AmendmentEffective"
                    Format ="Short Date"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Revision Date Required"
                    ControlTipText ="Date the Amendment became Effective"
                    GroupTable =1

                    LayoutCachedLeft =2652
                    LayoutCachedTop =1620
                    LayoutCachedWidth =5232
                    LayoutCachedHeight =1932
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1800
                            Top =1620
                            Width =780
                            Height =312
                            FontSize =9
                            FontWeight =700
                            Name ="Label87"
                            Caption ="Effective:"
                            GroupTable =1
                            LayoutCachedLeft =1800
                            LayoutCachedTop =1620
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =1932
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2652
                    Top =2388
                    Width =2580
                    Height =312
                    TabIndex =9
                    Name ="AmendmentBack"
                    ControlSource ="AmendmentBack"
                    Format ="Short Date"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Revision Date Required"
                    ControlTipText ="Date the Amendment was returned"
                    GroupTable =1

                    LayoutCachedLeft =2652
                    LayoutCachedTop =2388
                    LayoutCachedWidth =5232
                    LayoutCachedHeight =2700
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1800
                            Top =2388
                            Width =780
                            Height =312
                            FontSize =9
                            FontWeight =700
                            Name ="Label93"
                            Caption ="Back:"
                            GroupTable =1
                            LayoutCachedLeft =1800
                            LayoutCachedTop =2388
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =2700
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6180
                    Top =4440
                    Width =9660
                    Height =315
                    TabIndex =10
                    Name ="AmendmentDocLink"
                    ControlSource ="AmendmentDocLink"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Revision Date Required"
                    ControlTipText ="hyperlink to Amended Doc"

                    LayoutCachedLeft =6180
                    LayoutCachedTop =4440
                    LayoutCachedWidth =15840
                    LayoutCachedHeight =4755
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =10
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5340
                            Top =4440
                            Width =780
                            Height =315
                            FontSize =9
                            FontWeight =700
                            Name ="Label99"
                            Caption ="Doc Link:"
                            LayoutCachedLeft =5340
                            LayoutCachedTop =4440
                            LayoutCachedWidth =6120
                            LayoutCachedHeight =4755
                            RowStart =6
                            RowEnd =6
                            ForeThemeColorIndex =2
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
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
' See "ScheduleAmends.cls"
